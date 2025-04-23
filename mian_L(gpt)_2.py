import os
import sys
import re

import xlwings as xw
from xlwings.constants import Calculation

from PyPDF2 import PdfMerger

from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

from src.utils.utils import resource_path, is_file_open

# UI 로드

# # 코드
# # 현재 utils.py (__file__) 기준으로 한 단계 위(src/)로 올라간 뒤 interface 폴더로
# ui_relative = os.path.join('..', 'interface', 'PE_main.ui')
# form = resource_path(ui_relative)

# EXE
form = resource_path(os.path.join('interface', 'PE_main.ui'))

form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 아이콘 설정

        # # 코드
        # # utils.resource_path 기준으로 icons 폴더를 찾아서 QIcon에 전달
        # icon_rel = os.path.join('..', '..', 'icons', 'cikw.png')

        # EXE
        icon_rel = os.path.join('icons', 'cikw.png')

        icon_path = resource_path(icon_rel)
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

        # ── 1) 전체 시트 개수 미리 계산 ─────────────────────────
        # Excel 앱 띄워서 시트 이름만 읽고 닫기
        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts  = False
        app.api.Calculation  = Calculation.xlCalculationManual
        app.api.EnableEvents = False

        file_sheets = []
        total_sheets = 0
        for file_path in excel_files:
            if is_file_open(file_path):
                continue
            wb = app.books.open(file_path)
            names = [s.name for s in wb.sheets]
            file_sheets.append((file_path, names))
            total_sheets += len(names)
            wb.close()

        if total_sheets == 0:
            print("⚠️ 처리할 시트가 없습니다.")
            app.quit()
            return

        # ── 2) 프로그래스 다이얼로그 띄우기 ─────────────────────────
        progress = QProgressDialog("PDF 변환 중...", None, 0, total_sheets, self)
        progress.setWindowTitle("진행 상태")
        progress.setWindowModality(Qt.WindowModal)
        progress.setCancelButton(None)
        progress.setWindowFlags(progress.windowFlags() | Qt.WindowStaysOnTopHint)
        progress.show()

        # ── 3) 본격 변환 루프 ───────────────────────────────────
        processed = 0
        for file_path, sheet_names in file_sheets:
            if is_file_open(file_path):
                print(f"⚠️ 열려있는 파일 스킵: {file_path}")
                processed += len(sheet_names)
                progress.setValue(processed)
                QApplication.processEvents()
                continue

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            pdf_list = []
            try:
                wb = app.books.open(file_path)
                for idx, sheet_name in enumerate(sheet_names):
                    sheet = wb.sheets[sheet_name]
                    # PDF 저장 경로: 각각 output/파일명/Sheets/파일명_idx_시트.pdf
                    dest_dir = os.path.join(out_folder, base_name, 'Sheets')
                    os.makedirs(dest_dir, exist_ok=True)
                    pdf_path = os.path.join(dest_dir, f"{base_name}_{idx}_{sheet_name}.pdf")

                    if sheet.api.Visible == -1:
                        try:
                            sheet.to_pdf(pdf_path)
                            pdf_list.append(pdf_path)
                        except Exception as e:
                            print(f"⚠️ PDF 변환 오류 ({sheet_name}): {e}")
                    # 숨김 시트든 오류든, 일단 '처리된 시트'로 간주
                    processed += 1
                    progress.setValue(processed)
                    QApplication.processEvents()

                wb.close()
            except Exception as e:
                print(f"❌ 파일 처리 실패 ({base_name}): {e}")
                # 남은 시트 수치 채우기
                processed += len(sheet_names) - (processed - (total_sheets - len(sheet_names)))
                progress.setValue(processed)
                QApplication.processEvents()
                continue

            # 병합
            if pdf_list:
                merged_dir = os.path.join(out_folder, base_name, "Merged")
                os.makedirs(merged_dir, exist_ok=True)
                merged_path = os.path.join(merged_dir, f"{base_name}_merged.pdf")
                self.mergePdfs(pdf_list, merged_path)

        # ── 4) 마무리 ────────────────────────────────────
        progress.close()
        app.quit()
        QMessageBox.information(self, "완료", "모든 PDF 추출이 완료되었습니다.")


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
