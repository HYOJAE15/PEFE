import os
import sys

import xlwings as xw
from xlwings.constants import Calculation
from PyPDF2 import PdfMerger

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QProgressDialog, QMessageBox
)
from PyQt5 import uic
from PyQt5.QtGui import QIcon, QMovie
from PyQt5.QtCore import Qt

from src.utils.utils import resource_path, is_file_open

# UI 로드

# 코드
# 현재 utils.py (__file__) 기준으로 한 단계 위(src/)로 올라간 뒤 interface 폴더로
ui_relative = os.path.join('..', 'interface', 'PE_main_gpt_V7.ui')
form = resource_path(ui_relative)

# # EXE
# form = resource_path(os.path.join('interface', 'PE_main.ui'))

form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 아이콘 설정

        # 코드
        # utils.resource_path 기준으로 icons 폴더를 찾아서 QIcon에 전달
        icon_rel = os.path.join('..', '..', 'icons', 'cikw.png')

        # # EXE
        # icon_rel = os.path.join('icons', 'cikw.png')

        icon_path = resource_path(icon_rel)
        self.setWindowIcon(QIcon(icon_path))

        #── 춤추는 Chiikawa 애니 GIF 준비 ──
        gif_path = resource_path(os.path.join('..', '..', 'icons','cikw.gif'))
        self.danceMovie = QMovie(gif_path)
        self.danceLabel.setMovie(self.danceMovie)
        self.danceLabel.setVisible(False)

        # 버튼 시그널
        # UI에는 fileButton이 없으므로 folderButton만 연결
        self.folderButton.clicked.connect(self.forFolder)

    def forFolder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Directory",
            "",
            QFileDialog.DontUseNativeDialog
        )
        if not folder:
            self.statusLabel.setText("폴더 선택 취소됨")
            return

        # 초기 상태
        self.statusLabel.setText("작업 준비 중…")
        QApplication.processEvents()

        out_folder = os.path.join(folder, "output")
        os.makedirs(out_folder, exist_ok=True)

        self.processExcelFiles(folder, out_folder)

    def processExcelFiles(self, folder_path, out_folder):

        # 변환 시작 직전에
        self.danceLabel.setVisible(True)
        self.danceMovie.start()

        valid_ext = ('.xls', '.xlsx', '.xlsm')
        excel_files = [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.lower().endswith(valid_ext) and not f.startswith("~$")
        ]
        if not excel_files:
            self.statusLabel.setText("유효한 Excel 파일이 없습니다")
            return

        # Excel 앱 셋업 (속도 최적화)
        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts = False
        app.api.Calculation = Calculation.xlCalculationManual
        app.api.EnableEvents = False

        # 전체 시트 개수 계산
        file_sheets = []
        total_sheets = 0
        for fp in excel_files:
            if is_file_open(fp):
                continue
            wb = app.books.open(fp)
            names = [s.name for s in wb.sheets]
            file_sheets.append((fp, names))
            total_sheets += len(names)
            wb.close()

        if total_sheets == 0:
            self.statusLabel.setText("처리할 시트가 없습니다")
            app.quit()
            return

        # 진행 다이얼로그
        progress = QProgressDialog("PDF 변환 중…", None, 0, total_sheets, self)
        progress.setWindowTitle("진행 상태")
        progress.setWindowModality(Qt.WindowModal)
        progress.setCancelButton(None)
        progress.setWindowFlags(progress.windowFlags() | Qt.WindowStaysOnTopHint)
        progress.show()

        processed = 0

        # 파일별, 시트별 처리
        for file_path, sheets in file_sheets:
            base = os.path.splitext(os.path.basename(file_path))[0]
            self.statusLabel.setText(f"처리 중: {base}")
            QApplication.processEvents()

            pdf_list = []
            try:
                wb = app.books.open(file_path)
                for idx, name in enumerate(sheets, start=1):
                    self.statusLabel.setText(f"{base} – 시트 {idx}/{len(sheets)}: {name}")
                    QApplication.processEvents()

                    dest = os.path.join(out_folder, base, 'Sheets')
                    os.makedirs(dest, exist_ok=True)
                    pdf_path = os.path.join(dest, f"{base}_{idx-1}_{name}.pdf")

                    if wb.sheets[name].api.Visible == -1:
                        try:
                            wb.sheets[name].to_pdf(pdf_path)
                            pdf_list.append(pdf_path)
                        except Exception as e:
                            print(f"⚠️ PDF 변환 오류 ({name}): {e}")

                    processed += 1
                    progress.setValue(processed)
                    QApplication.processEvents()

                wb.close()
            except Exception as e:
                print(f"❌ 파일 처리 실패 ({base}): {e}")
                # 남은 시트 건너뛰기
                processed += len(sheets) - (processed - (total_sheets - len(sheets)))
                progress.setValue(processed)
                QApplication.processEvents()
                continue

            # Merge 체크박스 상태에 따라 병합 여부 결정
            if self.mergeCheckBox.isChecked() and pdf_list:
                merged_dir = os.path.join(out_folder, base, "Merged")
                os.makedirs(merged_dir, exist_ok=True)
                merged_path = os.path.join(merged_dir, f"{base}_merged.pdf")

                self.statusLabel.setText(f"{base} 병합 중…")
                QApplication.processEvents()

                self.mergePdfs(pdf_list, merged_path)

                self.statusLabel.setText(f"{base} 병합 완료")
            else:
                self.statusLabel.setText(f"{base} 병합 건너뜀")
            QApplication.processEvents()

        # 마무리
        # 모든 작업 완료 후
        self.danceMovie.stop()
        self.danceLabel.setVisible(False)
        progress.close()
        app.quit()

        self.statusLabel.setText("모든 작업이 완료되었습니다")
        QMessageBox.information(self, "완료", "모든 PDF 추출 작업이 완료되었습니다.")

    def mergePdfs(self, pdf_list, output_path):
        merger = PdfMerger()
        for p in pdf_list:
            merger.append(p)
        merger.write(output_path)
        merger.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = WindowClass()
    win.show()
    sys.exit(app.exec_())
