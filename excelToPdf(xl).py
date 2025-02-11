import os
import sys
import time
import xlwings as xw
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

        # Connect buttons
        self.fileButton.clicked.connect(self.forFile)
        self.folderButton.clicked.connect(self.forFolder)

    def forFile(self):
        self.file = self.dialog.getOpenFileName(caption="Select File", filter="Excel Files (*.xls *.xlsx *.xlsm)")
        print(f"Selected File: {self.file[0]}")

    def forFolder(self):
        self.folder = self.dialog.getExistingDirectory(caption="Select Directory")

        if not self.folder:
            print("⚠️ No folder selected.")
            return

        self.excelInfo(self.folder)

        self.out_folder = os.path.join(self.folder, "output")
        os.makedirs(self.out_folder, exist_ok=True)

        self.transPDF(self.result, self.out_folder)

    def excelInfo(self, filepath):
        valid_extensions = ('.xls', '.xlsx', '.xlsm')

        # Exclude temporary Excel files (~$)
        excel_list_valid = [
            os.path.join(filepath, f) for f in os.listdir(filepath)
            if f.lower().endswith(valid_extensions) and not f.startswith("~$")
        ]
        print(f"Valid Excel Files: {excel_list_valid}")

        self.result = []
        for file in excel_list_valid:
            if is_file_open(file):
                print(f"⚠️ File is open, skipping: {file}")
                continue  

            with xw.App(visible=False) as app:
                wb = app.books.open(file)
                sheet_names = [s.name for s in wb.sheets]

                print(f"Processed: {os.path.basename(file)} - Sheets: {sheet_names}")

                for sheet in sheet_names:
                    self.result.append((file, os.path.splitext(os.path.basename(file))[0], sheet))

                wb.close()

    def transPDF(self, fileinfo, savepath):
        os.makedirs(savepath, exist_ok=True)

        for i, (file_path, filename, sheet_name) in enumerate(fileinfo):
            if is_file_open(file_path):
                print(f"⚠️ File is open, skipping: {file_path}")
                continue

            with xw.App(visible=False) as app:
                try:
                    wb = app.books.open(file_path)
                    sheet = wb.sheets[sheet_name]

                    pdf_filename = f"{i}_{filename}_{sheet_name}.pdf"
                    pdf_path = os.path.join(savepath, pdf_filename)

                    if os.path.exists(pdf_path):
                        try:
                            os.remove(pdf_path)
                            time.sleep(1)
                        
                        except PermissionError:
                            print(f"❌ PDF file is open, skipping: {pdf_path}")
                            continue

                    print(f"📄 Saving: {pdf_path}")

                    if sheet.api.Visible == -1 : 
                        sheet.to_pdf(pdf_path)
                        wb.close()
                        print(f"✅ Successfully saved: {pdf_path}")
                    else :
                        print(f"🙈 Sheet: {sheet_name} → Status: {sheet.api.Visible}(0=hidden, 2=very hidden)")
                    
                except Exception as e:
                    print(f"❌ Error converting to PDF: {e}")
                    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
