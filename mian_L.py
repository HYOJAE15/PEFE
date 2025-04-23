import os
import sys
import xlwings as xw
from xlwings.constants import Calculation
from PyPDF2 import PdfMerger
from PyQt5.QtWidgets import *
from PyQt5 import uic

from src.utils.utils import resource_path, is_file_open

# 현재 utils.py (__file__) 기준으로 한 단계 위(src/)로 올라간 뒤 interface 폴더로
ui_relative = os.path.join('..', 'interface', 'PE_main.ui')
form = resource_path(ui_relative)

# 또는 절대 경로로 바로 지정하고 싶다면 resource_path를 쓰지 않고:
# script_dir = os.path.dirname(os.path.abspath(__file__))  # 만약 mian_L.py가 src/에 있으면
# form = os.path.join(script_dir, 'interface', 'PE_main.ui')

form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.dialog = QFileDialog()

        # 버튼 연결
        self.fileButton.clicked.connect(self.forFile)
        self.folderButton.clicked.connect(self.forFolder)

    def forFile(self):
        """ 필요 시 단일 파일만 선택해서 처리할 경우 사용 """
        self.file = self.dialog.getOpenFileName(
            caption="Select File", 
            filter="Excel Files (*.xls *.xlsx *.xlsm)"
        )
        print(f"Selected File: {self.file[0]}")

    def forFolder(self):
        """ 폴더를 선택해 내부의 모든 Excel 파일 처리 """
        self.folder = self.dialog.getExistingDirectory(caption="Select Directory")

        if not self.folder:
            print("⚠️ No folder selected.")
            return

        # 아웃풋 폴더 생성
        self.out_folder = os.path.join(self.folder, "output")
        os.makedirs(self.out_folder, exist_ok=True)

        # 폴더 내부 Excel 파일 처리
        self.processExcelFiles(self.folder, self.out_folder)

    def processExcelFiles(self, folder_path, out_folder):
        valid_extensions = ('.xls', '.xlsx', '.xlsm')
        excel_files = [
            os.path.join(folder_path, f) for f in os.listdir(folder_path)
            if f.lower().endswith(valid_extensions) and not f.startswith("~$")
        ]

        if not excel_files:
            print("⚠️ 유효한 Excel 파일이 없습니다.")
            return

        print(f"유효한 Excel 파일 목록: {excel_files}")

        # (1) App 객체를 한 번만 생성 → 성능 개선
        app = xw.App(visible=False)
        # 속도 최적화: 화면 업데이트, 경고, 자동계산, 이벤트 비활성화
        app.screen_updating = False
        app.display_alerts = False
        app.api.Calculation = Calculation.xlCalculationManual
        app.api.EnableEvents = False

        try:
            for file_path in excel_files:
                if is_file_open(file_path):
                    print(f"⚠️ 열려있는 파일은 스킵: {file_path}")
                    continue

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                pdf_list = []

                try:
                    wb = app.books.open(file_path)
                    sheet_names = [s.name for s in wb.sheets]
                    print(f"처리 대상: {os.path.basename(file_path)} → 시트: {sheet_names}")

                    for idx, sheet_name in enumerate(sheet_names):
                        sheet = wb.sheets[sheet_name]
                        pdf_filename = f"{base_name}_{idx}_{sheet_name}.pdf"
                        pdf_path = os.path.join(out_folder, pdf_filename)

                        if sheet.api.Visible == -1:
                            try:
                                sheet.to_pdf(pdf_path)
                                pdf_list.append(pdf_path)
                                print(f"✅ PDF 생성 완료: {pdf_path}")
                            except Exception as e:
                                print(f"⚠️ PDF 생성 중 오류 ({sheet_name}): {e}")
                                # 오류 발생해도 다음 시트로 계속 진행
                                continue
                        else:
                            print(f"🙈 숨김 시트 스킵: {sheet_name}")

                    wb.close()
                except Exception as e:
                    print(f"❌ Excel 처리 중 오류 ({base_name}): {e}")
                    continue  # 다음 파일로 이동

                # 파일 단위로 PDF 병합
                if pdf_list:
                    merged_pdf_name = f"{base_name}_merged.pdf"
                    merged_pdf_path = os.path.join(out_folder, merged_pdf_name)
                    self.mergePdfs(pdf_list, merged_pdf_path)
                else:
                    print(f"⚠️ 변환된 PDF가 없어 병합하지 않음: {base_name}")

        finally:
            app.quit()

    def mergePdfs(self, pdf_list, output_path):
        if not pdf_list:
            print("⚠️ 병합할 PDF 목록이 비어 있습니다.")
            return

        merger = PdfMerger()
        print(f"📂 병합할 PDF 총 {len(pdf_list)}개…")
        for pdf in pdf_list:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()
        print(f"✅ 병합 완료: {output_path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
