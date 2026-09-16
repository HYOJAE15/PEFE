import os
import sys

import xlwings as xw
from xlwings.constants import Calculation
from PyPDF2 import PdfMerger

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog,
    QProgressDialog, QMessageBox, QSystemTrayIcon
)
from PyQt5 import uic
from PyQt5.QtGui import QIcon, QMovie
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from src.utils.utils import resource_path, is_file_open

# 1) UI 로드
ui_path = resource_path(os.path.join('..', 'interface', 'PE_main_gpt_V8.ui'))
FormClass, _ = uic.loadUiType(ui_path)


# 2) 변환 작업을 백그라운드에서 수행할 Worker
class ConverterWorker(QThread):
    progress = pyqtSignal(int)
    status   = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, file_sheets, out_folder, do_merge, do_split):
        super().__init__()
        self.file_sheets = file_sheets
        self.out_folder  = out_folder
        self.do_merge    = do_merge
        self.do_split    = do_split  # 시트별 저장 On/Off 옵션 추가

    def run(self):
        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts  = False
        app.api.Calculation  = Calculation.xlCalculationManual
        app.api.EnableEvents = False

        processed = 0

        for file_path, sheets in self.file_sheets:
            base = os.path.splitext(os.path.basename(file_path))[0]
            self.status.emit(f"처리 중: {base}")

            pdf_list = []
            try:
                wb = app.books.open(file_path)

                # --- [A] 시트별 개별 저장 (do_split = True) ---
                if self.do_split:
                    dest = os.path.join(self.out_folder, base, 'Sheets')
                    os.makedirs(dest, exist_ok=True)

                    for idx, name in enumerate(sheets, start=1):
                        self.status.emit(f"{base} – 시트 {idx}/{len(sheets)}: {name}")
                        pdf_path = os.path.join(dest, f"{base}_{idx-1}_{name}.pdf")

                        if wb.sheets[name].api.Visible == -1:
                            try:
                                wb.sheets[name].to_pdf(pdf_path)
                                pdf_list.append(pdf_path)
                            except Exception as e:
                                print(f"⚠️ PDF 변환 오류 ({name}): {e}")

                        processed += 1
                        self.progress.emit(processed)

                # --- [B] 시트별 개별 저장 미사용 (do_split = False) ---
                else:
                    self.status.emit(f"{base} – 전체 파일 변환 중")
                    # 통합 저장용 임시 디렉토리
                    temp_dir = os.path.join(self.out_folder, base, 'Temp')
                    os.makedirs(temp_dir, exist_ok=True)
                    full_pdf_path = os.path.join(temp_dir, f"{base}_full.pdf")

                    wb.to_pdf(full_pdf_path)
                    pdf_list.append(full_pdf_path)

                    processed += len(sheets)
                    self.progress.emit(processed)

                wb.close()

            except Exception as e:
                print(f"❌ 파일 처리 실패 ({base}): {e}")
                processed += len(sheets)
                self.progress.emit(processed)
                continue

            # --- [C] 병합 처리 (do_merge = True) ---
            if self.do_merge and pdf_list:
                merge_dir = os.path.join(self.out_folder, base, 'Merged')
                os.makedirs(merge_dir, exist_ok=True)
                merged_path = os.path.join(merge_dir, f"{base}_merged.pdf")

                self.status.emit(f"{base} 병합 중…")
                self._merge_pdfs(pdf_list, merged_path)
                self.status.emit(f"{base} 병합 완료")
            else:
                self.status.emit(f"{base} 병합 건너뜀")

        app.quit()
        self.finished.emit()

    def _merge_pdfs(self, pdf_list, output_path):
        merger = PdfMerger()
        for p in pdf_list:
            merger.append(p)
        merger.write(output_path)
        merger.close()


# 3) 메인 윈도우
class WindowClass(QMainWindow, FormClass):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 아이콘 설정
        icon_path = resource_path(os.path.join('..', '..', 'icons', 'cikw.png'))
        self.setWindowIcon(QIcon(icon_path))

        # 트레이 알림용 아이콘
        self.tray = QSystemTrayIcon(QIcon(icon_path), self)
        self.tray.show()

        # Chiikawa GIF 설정
        gif_path = resource_path(os.path.join('..', '..', 'icons', 'cikw.gif'))
        self.danceMovie = QMovie(gif_path)
        self.danceLabel.setMovie(self.danceMovie)
        self.danceLabel.setVisible(False)

        # 버튼 연결
        self.folderButton.clicked.connect(self.onSelectFolder)

    def onSelectFolder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "폴더 선택", "", QFileDialog.DontUseNativeDialog
        )
        if not folder:
            self.statusLabel.setText("폴더 선택 취소됨")
            return

        self.statusLabel.setText("작업 준비 중…")
        QApplication.processEvents()

        out_folder = os.path.join(folder, "output")
        os.makedirs(out_folder, exist_ok=True)

        # 엑셀 파일 목록 & 시트 수집
        valid_ext = ('.xls', '.xlsx', '.xlsm')
        excel_files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.lower().endswith(valid_ext) and not f.startswith("~$")
        ]
        file_sheets = []
        total_sheets = 0
        tmp_app = xw.App(visible=False)
        tmp_app.screen_updating = False
        tmp_app.display_alerts  = False
        tmp_app.api.Calculation  = Calculation.xlCalculationManual
        tmp_app.api.EnableEvents = False

        for fp in excel_files:
            if is_file_open(fp):
                continue
            wb = tmp_app.books.open(fp)
            names = [s.name for s in wb.sheets]
            wb.close()
            file_sheets.append((fp, names))
            total_sheets += len(names)
        tmp_app.quit()

        if total_sheets == 0:
            self.statusLabel.setText("처리할 시트가 없습니다")
            return

        # 프로그래스바 설정
        self.progressDialog = QProgressDialog("PDF 변환 중…", None, 0, total_sheets, self)
        self.progressDialog.setWindowTitle("진행 상태")
        self.progressDialog.setWindowModality(Qt.WindowModal)
        self.progressDialog.setCancelButton(None)
        self.progressDialog.show()

        # 위치 이동
        dlg_size = self.progressDialog.sizeHint()
        x = self.x() + self.width() - dlg_size.width() - 20
        y = self.y() + 20
        self.progressDialog.move(x, y)

        # GIF 애니메이션 시작
        self.danceLabel.setVisible(True)
        self.danceMovie.start()

        # UI 요소 상태 가져오기 (.ui 파일 내 splitCheckBox 이름 기준)
        do_merge = self.mergeCheckBox.isChecked()
        
        # splitCheckBox가 체크박스 형태인 경우 (없으면 기본값 True 설정)
        do_split = getattr(self, 'splitCheckBox', None)
        do_split_val = do_split.isChecked() if do_split else True

        # Worker 실행
        self.worker = ConverterWorker(file_sheets, out_folder, do_merge, do_split_val)
        self.worker.progress.connect(self.progressDialog.setValue)
        self.worker.status.connect(self.statusLabel.setText)
        self.worker.finished.connect(self.onFinished)
        self.worker.start()

    def onFinished(self):
        # GIF 정지
        self.danceMovie.stop()
        self.danceLabel.setVisible(False)
        self.progressDialog.close()

        # 최종 상태 표시
        self.statusLabel.setText("모든 작업이 완료되었습니다")

        # 트레이 알림
        self.tray.showMessage(
            "PEFE 완료",
            "PDF 추출 작업이 완료되었습니다!",
            QSystemTrayIcon.Information,
            5000
        )

        # 알림 메시지 박스
        msg = QMessageBox(self)
        msg.setWindowTitle("작업 완료 🎉")
        msg.setText("🎉 모든 PDF 추출 작업이 완료되었습니다! 🎉")
        msg.setIcon(QMessageBox.Information)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setWindowFlags(msg.windowFlags() | Qt.WindowStaysOnTopHint)
        msg.setStyleSheet(
            "QLabel{min-width:250px; font-size:14pt; color:#186F9A;} "
            "QPushButton{min-width:80px; font-size:12pt; padding:8px;}"
        )
        msg.exec_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = WindowClass()
    win.show()
    sys.exit(app.exec_())