import os
import sys

import xlwings as xw
from xlwings.constants import Calculation

from PyPDF2 import PdfMerger

from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

from src.utils.utils import resource_path, is_file_open

# UI 로드
form = resource_path(os.path.join('..', 'interface', 'PE_main.ui'))
form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 아이콘 설정
        icon_path = resource_path(os.path.join('..', '..', 'icons', 'cikw.png'))
        self.setWindowIcon(QIcon(icon_path))

        self.dialog = QFileDialog()

        # 버튼 연결
        self.fileButton.clicked.connect(self.forFile)
        self.folderButton.clicked.connect(self.forFolder)

    def forFile(self):
        self.file = self.dialog.getOpenFileName(
            caption="Select File",
            filter="Excel Files (*.xls *.xlsx *.xlsm)"
        )
        print(f"Selected File: {self.file[0]}")

    def forFolder(self):
        self.folder = self.dialog.getExistingDirectory(caption="Select Directory")
        if not self.folder:
            print("⚠️ No folder selected.")
            return

        self.out_folder = os.path.join(self.folder, "output")
        os.makedirs(self.out_folder, exist_ok=True)

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

        total_files = len(excel_files)
        progress = QProgressDialog("PDF 변환 중...", "취소", 0, total_files, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setValue(0)

        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts = False
        app.api.Calculation = Calculation.xlCalculationManual
        app.api.EnableEvents = False

        cancelled = False
        try:
            for idx_file, file_path in enumerate(excel_files):
                if progress.wasCanceled():
                    cancelled = True
                    print("🚫 사용자가 작업을 취소했습니다.")
                    break

                if is_file_open(file_path):
                    print(f"⚠️ 열려있는 파일은 스킵: {file_path}")
                    progress.setValue(idx_file + 1)
                    continue

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                pdf_list = []

                # 각 파일별 폴더 생성
                file_out = os.path.join(out_folder, base_name)
                os.makedirs(file_out, exist_ok=True)

                try:
                    wb = app.books.open(file_path)
                    sheet_names = [s.name for s in wb.sheets]
                    print(f"처리 대상: {os.path.basename(file_path)} → 시트: {sheet_names}")

                    for idx_sheet, sheet_name in enumerate(sheet_names):
                        sheet = wb.sheets[sheet_name]
                        pdf_filename = f"{base_name}_{idx_sheet}_{sheet_name}.pdf"
                        pdf_path = os.path.join(file_out, pdf_filename)

                        if sheet.api.Visible == -1:
                            try:
                                sheet.to_pdf(pdf_path)
                                pdf_list.append(pdf_path)
                                print(f"✅ PDF 생성 완료: {pdf_path}")
                            except Exception as e:
                                print(f"⚠️ PDF 생성 중 오류 ({sheet_name}): {e}")
                                continue
                        else:
                            print(f"🙈 숨김 시트 스킵: {sheet_name}")

                    wb.close()
                except Exception as e:
                    print(f"❌ Excel 처리 중 오류 ({base_name}): {e}")
                    progress.setValue(idx_file + 1)
                    continue

                # PDF 병합
                if pdf_list:
                    merged_pdf = os.path.join(file_out, f"{base_name}_merged.pdf")
                    self.mergePdfs(pdf_list, merged_pdf)
                else:
                    print(f"⚠️ 변환된 PDF가 없어 병합하지 않음: {base_name}")

                progress.setValue(idx_file + 1)
        finally:
            app.quit()
            progress.close()

        if not cancelled:
            QMessageBox.information(self, "완료", "모든 PDF 변환이 완료되었습니다.")

    def mergePdfs(self, pdf_list, output_path):
        if not pdf_list:
            print("⚠️ 병합할 PDF 목록이 비어 있습니다.")
            return

        merger = PdfMerger()
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
