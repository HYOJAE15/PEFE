import os
import sys
import time
import xlwings as xw
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5 import uic

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path('PE_main.ui')
form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.dialog = QFileDialog()

        # Connect buttons
        self.fileButton.clicked.connect(self.forFile)
        self.folderButton.clicked.connect(self.forFolder)

    def forFile(self):
        self.file = self.dialog.getOpenFileName(caption="Select File", filter="Excel Files (*.xls *.xlsx *.xlsm)")
        print(self.file)

    def forFolder(self):
        self.folder = self.dialog.getExistingDirectory(caption="Select Directory")

        self.excelInfo(self.folder)

        self.out_folder = os.path.join(self.folder, "output")
        os.makedirs(self.out_folder, exist_ok=True)

        self.transPDF(self.result, self.out_folder)

    def excelInfo(self, filepath):
        valid_extensions = ('.xls', '.xlsx', '.xlsm')

        # Exclude temporary Excel files (~$)
        excel_list_valid = [
            os.path.join(self.folder, f) for f in os.listdir(self.folder)
            if f.lower().endswith(valid_extensions) and not f.startswith("~$")
        ]
        print(f"Valid Excel Files: {excel_list_valid}")

        self.result = []
        for file in excel_list_valid:
            if self.is_file_open(file):
                print(f"⚠️ File is open, skipping: {file}")
                continue  # Skip open files

            # Read sheet names using pandas
            try:
                sheet_names = pd.ExcelFile(file, engine="openpyxl").sheet_names
            except Exception as e:
                print(f"❌ Error reading sheets: {e}")
                continue

            filename, _ = os.path.splitext(os.path.basename(file))

            for sheet in sheet_names:
                temp_tuple = (file, filename, sheet)
                self.result.append(temp_tuple)

            print(f"Processed: {filename} - Sheets: {sheet_names}")

    def transPDF(self, fileinfo, savepath):
        app = xw.App(visible=False)  # Start Excel in the background

        for i, info in enumerate(fileinfo):
            file_path = os.path.abspath(info[0])
            sheet_name = info[2]

            if self.is_file_open(file_path):
                print(f"⚠️ File is open, skipping: {file_path}")
                continue

            try:
                wb = app.books.open(file_path)
                ws = wb.sheets[sheet_name]

                # PDF output path
                pdf_filename = f"{i}_{info[1]}_{info[2]}.pdf"
                pdf_path = os.path.join(savepath, pdf_filename)

                if os.path.exists(pdf_path):
                    try:
                        os.remove(pdf_path)  # Remove existing file
                        time.sleep(1)
                    except PermissionError:
                        print(f"❌ {pdf_path} is open. Close it and try again.")
                        wb.close()
                        continue

                # Export as PDF
                ws.api.ExportAsFixedFormat(0, pdf_path)

                print(f"✅ Saved: {pdf_path}")

                wb.close()
            except Exception as e:
                print(f"❌ Error converting to PDF: {e}")
                continue

        app.quit()
        print("📄 PDF conversion complete.")

    def is_file_open(self, file_path):
        """Check if a file is open."""
        try:
            with open(file_path, "r+"):
                return False
        except IOError:
            return True  # File is open

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
