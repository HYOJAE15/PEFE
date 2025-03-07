import os
import sys
import xlwings as xw
from PyPDF2 import PdfMerger
from PyQt5.QtWidgets import *
from PyQt5 import uic

from utils import resource_path, is_file_open

form = resource_path('PE_main.ui')
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
        """
        폴더 내의 모든 Excel 파일을 찾아 각 파일의 시트를 개별 PDF로 저장하고,
        파일 단위로 PDF들을 병합한다.
        """
        valid_extensions = ('.xls', '.xlsx', '.xlsm')
        excel_files = [
            os.path.join(folder_path, f) for f in os.listdir(folder_path)
            if f.lower().endswith(valid_extensions) and not f.startswith("~$")
        ]

        if not excel_files:
            print("⚠️ 유효한 Excel 파일이 없습니다.")
            return

        print(f"유효한 Excel 파일 목록: {excel_files}")

        # (1) App 객체를 한 번만 생성해 놓고 모든 파일을 처리 → 성능 개선
        app = xw.App(visible=False)
        try:
            for file_path in excel_files:
                if is_file_open(file_path):
                    print(f"⚠️ 열려있는 파일은 스킵: {file_path}")
                    continue

                base_name = os.path.splitext(os.path.basename(file_path))[0]

                # 현재 파일에 대한 PDF 경로 리스트
                pdf_list = []

                try:
                    # 한 파일에 대해서만 Workbook 오픈
                    wb = app.books.open(file_path)
                    sheet_names = [s.name for s in wb.sheets]
                    print(f"처리 대상: {os.path.basename(file_path)} → 시트: {sheet_names}")

                    for idx, sheet_name in enumerate(sheet_names):
                        sheet = wb.sheets[sheet_name]

                        # PDF 파일명 예: "파일명_인덱스_시트이름.pdf"
                        pdf_filename = f"{base_name}_{idx}_{sheet_name}.pdf"
                        pdf_path = os.path.join(out_folder, pdf_filename)

                        # 시트가 표시되어 있을 때만 PDF 변환
                        if sheet.api.Visible == -1:
                            # (2) 기존 파일이 있을 경우 굳이 삭제 후 기다리지 않고 바로 덮어쓰기 시도
                            # 필요시 try-except로 PermissionError를 잡아서 처리 가능
                            sheet.to_pdf(pdf_path)
                            pdf_list.append(pdf_path)
                            print(f"✅ PDF 생성 완료: {pdf_path}")
                        else:
                            print(f"🙈 숨김 시트 스킵: {sheet_name} (Visible={sheet.api.Visible})")

                    wb.close()

                except Exception as e:
                    print(f"❌ Excel 처리 중 오류 발생: {e}")
                    continue  # 다음 파일 처리로 이동

                # 파일 단위로 PDF 병합
                if pdf_list:
                    merged_pdf_name = f"{base_name}_merged.pdf"
                    merged_pdf_path = os.path.join(out_folder, merged_pdf_name)
                    self.mergePdfs(pdf_list, merged_pdf_path)
                else:
                    print(f"⚠️ 변환된 PDF가 없으므로 병합을 진행하지 않음: {base_name}")

        finally:
            # (3) 모든 파일 처리를 마친 뒤 app 종료
            app.quit()

    def mergePdfs(self, pdf_list, output_path):
        """ pdf_list에 주어진 PDF들을 순서대로 하나로 병합 """
        if not pdf_list:
            print("⚠️ 병합할 PDF 목록이 비어 있습니다.")
            return

        merger = PdfMerger()
        print(f"📂 병합할 PDF 총 {len(pdf_list)}개...")

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
