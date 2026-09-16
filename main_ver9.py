import os
import sys
import shutil

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
ui_path = resource_path(os.path.join('..', 'interface', 'PE_main_gpt_V7.ui'))
FormClass, _ = uic.loadUiType(ui_path)


# 2) 변환 작업을 백그라운드에서 수행할 Worker
class ConverterWorker(QThread):
    progress = pyqtSignal(int)
    status   = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, root_folder, file_sheets, out_folder, do_merge, do_split):
        super().__init__()
        self.root_folder = os.path.abspath(root_folder)
        self.file_sheets = file_sheets
        self.out_folder  = os.path.abspath(out_folder)
        self.do_merge    = do_merge
        self.do_split    = do_split

    def run(self):
        # xlwings App 생성 및 옵션 최적화
        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts  = False
        
        try:
            app.api.Calculation = Calculation.xlCalculationManual
            app.api.EnableEvents = False
            app.api.Interactive = False
        except Exception:
            pass

        processed = 0

        for file_path, sheets in self.file_sheets:
            abs_file_path = os.path.abspath(file_path)
            base = os.path.splitext(os.path.basename(abs_file_path))[0]
            
            # 절대 경로 기반 상대 경로 계산
            rel_dir = os.path.relpath(os.path.dirname(abs_file_path), self.root_folder)
            if rel_dir == ".":
                target_base_dir = self.out_folder
            else:
                target_base_dir = os.path.abspath(os.path.join(self.out_folder, rel_dir))

            self.status.emit(f"처리 중: {base}")
            pdf_list = []

            try:
                wb = app.books.open(abs_file_path)

                # --------------------------------------------------------
                # CASE 1: 시트별 분리 저장 (splitCheckBox ON)
                # --------------------------------------------------------
                if self.do_split:
                    dest = os.path.abspath(os.path.join(target_base_dir, base, 'Sheets'))
                    os.makedirs(dest, exist_ok=True)

                    for idx, name in enumerate(sheets, start=1):
                        self.status.emit(f"{base} – 시트 {idx}/{len(sheets)}: {name}")
                        
                        # 시트 이름 특수문자 정제 (파일명 오류 방지)
                        safe_sheet_name = "".join(c for c in name if c not in r'\/:*?"<>|')
                        pdf_path = os.path.abspath(os.path.join(dest, f"{base}_{idx-1}_{safe_sheet_name}.pdf"))

                        try:
                            sht = wb.sheets[name]
                            if sht.api.Visible == -1:  # xlSheetVisible (-1)
                                # Type=0 (xlTypePDF), Quality=0 (xlQualityStandard)
                                sht.api.ExportAsFixedFormat(
                                    Type=0,
                                    Filename=pdf_path,
                                    Quality=0,
                                    IncludeDocProperties=True,
                                    IgnorePrintAreas=False,
                                    OpenAfterPublish=False
                                )
                                if os.path.exists(pdf_path):
                                    pdf_list.append(pdf_path)
                        except Exception as e:
                            print(f"⚠️ PDF 변환 오류 ({name}): {e}")

                        processed += 1
                        self.progress.emit(processed)

                # --------------------------------------------------------
                # CASE 2: 통째로 파일 저장 (splitCheckBox OFF)
                # --------------------------------------------------------
                else:
                    self.status.emit(f"{base} – 전체 파일 저장 중")
                    dest = os.path.abspath(os.path.join(target_base_dir, base))
                    os.makedirs(dest, exist_ok=True)
                    
                    pdf_path = os.path.abspath(os.path.join(dest, f"{base}_full.pdf"))

                    try:
                        wb.api.ExportAsFixedFormat(
                            Type=0,
                            Filename=pdf_path,
                            Quality=0,
                            IncludeDocProperties=True,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                        if os.path.exists(pdf_path):
                            pdf_list.append(pdf_path)
                    except Exception as e:
                        print(f"⚠️ PDF 변환 오류 ({base}): {e}")

                    processed += len(sheets)
                    self.progress.emit(processed)

                wb.close()

            except Exception as e:
                print(f"❌ 파일 처리 실패 ({base}): {e}")
                processed += len(sheets)
                self.progress.emit(processed)
                continue

            # --------------------------------------------------------
            # 병합 처리 (mergeCheckBox ON)
            # --------------------------------------------------------
            if self.do_merge and pdf_list:
                merge_dir = os.path.abspath(os.path.join(target_base_dir, base, 'Merged'))
                os.makedirs(merge_dir, exist_ok=True)
                merged_path = os.path.abspath(os.path.join(merge_dir, f"{base}_merged.pdf"))

                self.status.emit(f"{base} 병합 중…")
                self._merge_pdfs(pdf_list, merged_path)
                self.status.emit(f"{base} 병합 완료")

                # split OFF + merge ON 일 때 생성했던 임시 full.pdf 제거
                if not self.do_split:
                    for temp_p in pdf_list:
                        if os.path.exists(temp_p) and temp_p != merged_path:
                            try:
                                os.remove(temp_p)
                            except Exception as e:
                                print(f"⚠️ 임시 파일 삭제 실패: {e}")

        app.quit()
        self.finished.emit()

    def _merge_pdfs(self, pdf_list, output_path):
        merger = PdfMerger()
        for p in pdf_list:
            if os.path.exists(p):
                merger.append(p)
        merger.write(output_path)
        merger.close()

# 3) 메인 윈도우
class WindowClass(QMainWindow, FormClass):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # 아이콘
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

        self.statusLabel.setText("작업 준비 중… (하위 폴더 검색 중)")
        QApplication.processEvents()

        out_folder = os.path.join(folder, "output")
        os.makedirs(out_folder, exist_ok=True)

        # --- [핵심] 하위 폴더 재귀적 전체 탐색 (os.walk 사용) ---
        valid_ext = ('.xls', '.xlsx', '.xlsm')
        excel_files = []

        for root, dirs, files in os.walk(folder):
            # output 폴더 내부로 계속 타고 들어가는 무한 루프 방지
            if os.path.commonpath([root, out_folder]) == out_folder:
                continue

            for f in files:
                if f.lower().endswith(valid_ext) and not f.startswith("~$"):
                    excel_files.append(os.path.join(root, f))

        file_sheets = []
        total_sheets = 0

        tmp_app = xw.App(visible=False)
        tmp_app.screen_updating = False
        tmp_app.display_alerts  = False
        tmp_app.api.Calculation  = Calculation.xlCalculationManual
        tmp_app.api.EnableEvents = False
        tmp_app.api.Interactive  = False

        for fp in excel_files:
            if is_file_open(fp):
                continue
            try:
                wb = tmp_app.books.open(fp)
                names = [s.name for s in wb.sheets]
                wb.close()
                file_sheets.append((fp, names))
                total_sheets += len(names)
            except Exception as e:
                print(f"⚠️ 파일 열기 실패 ({fp}): {e}")

        tmp_app.quit()

        if total_sheets == 0:
            self.statusLabel.setText("처리할 엑셀 파일/시트가 없습니다")
            return

        # 프로그래스바 설정
        self.progressDialog = QProgressDialog("PDF 변환 중…", None, 0, total_sheets, self)
        self.progressDialog.setWindowTitle("진행 상태")
        self.progressDialog.setWindowModality(Qt.WindowModal)
        self.progressDialog.setCancelButton(None)
        self.progressDialog.show()

        # 메인 윈도우 우측 상단 배치
        dlg_size = self.progressDialog.sizeHint()
        x = self.x() + self.width() - dlg_size.width() - 20
        y = self.y() + 20
        self.progressDialog.move(x, y)

        # GIF 애니메이션 시작
        self.danceLabel.setVisible(True)
        self.danceMovie.start()

        # UI 요소 값 가져오기
        do_merge = self.mergeCheckBox.isChecked()
        do_split = self.splitCheckBox.isChecked()

        # Worker 실행 (선택된 최상위 root_folder 경로 포함 전달)
        self.worker = ConverterWorker(folder, file_sheets, out_folder, do_merge, do_split)
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

        # 완료 메시지 박스
        msg = QMessageBox(self)
        msg.setWindowTitle("작업 완료 🎉")
        msg.setText("🎉 모든 하위 폴더의 PDF 추출 작업이 완료되었습니다! 🎉")
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