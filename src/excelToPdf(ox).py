import os
import sys

from PyQt5.QtWidgets import *
from PyQt5 import uic

from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3
import natsort
import win32com.client
import openpyxl as op
import time
import pythoncom
import ctypes


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))    
    return os.path.join(base_path, relative_path)

form = resource_path('PE_main.ui')
form_class = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super( ).__init__( )
        self.setupUi(self)

        self.dialog = QFileDialog()

        # sig
        self.fileButton.clicked.connect(self.forFile)
        self.folderButton.clicked.connect(self.forFolder)  
    
    # def
    def forFile (self):
        self.file = self.dialog.getOpenFileName(caption = "Select File", filter="excel (*.xls *.xlsx *.xlsm *.xlsb *.xltx *.xltm)")
        print(self.file)
        
    def forFolder (self):
        self.folder = self.dialog.ShowDirsOnly
        self.folder = self.dialog.getExistingDirectory(caption = "Select Directory")
        
        self.excelInfo(self.folder)
        
        self.out_folder = os.path.join(self.folder, "output")
        os.makedirs(self.out_folder, exist_ok=True)
        
        self.transPDF(self.result, self.out_folder)
        

    def excelInfo(self, filepath):
        valid_extensions = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm')

        # "~$" 임시 파일 제거
        excel_list_valid = [
            os.path.join(self.folder, f) for f in os.listdir(self.folder)
            if f.lower().endswith(valid_extensions) and not f.startswith("~$")
        ]
        print(f"excel_list_valid: {excel_list_valid}")

        self.result = []
        for file in excel_list_valid:
            if self.is_file_open(file):
                print(f"⚠️ 파일이 열려 있습니다. 닫고 다시 시도하세요: {file}")
                continue  # 파일이 열려 있으면 건너뛰기

            wb = op.load_workbook(file)
            ws_list = wb.sheetnames
            
            filename, extension = os.path.splitext(os.path.basename(file))

            print(f"{filename}_sheet_name{ws_list}")

            for sht in ws_list:
                temp_tuple = (file, filename, sht)
                self.result.append(temp_tuple)

            wb.close()
            print(f"temp_tuple: {self.result}")

    def transPDF(self, fileinfo, savepath):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        i = 0
        os.makedirs(savepath, exist_ok=True)

        for info in fileinfo:
            file_path = os.path.abspath(info[0])
            file_name = info[1]
            sheet_name = info[2]

            if self.is_file_open(file_path):
                print(f"⚠️ 파일이 열려 있습니다. 닫고 다시 시도하세요: {file_path}")
                continue
            wb = excel.Workbooks.Open(file_path)

            try:
                ws = wb.Worksheets(sheet_name)
                ws.Select()
            except Exception as e:
                print(f"❌ 시트 '{sheet_name}'를 찾을 수 없습니다: {e}")
                wb.Close(False)
                continue
            pdf_filename = f"{i}_{file_name}_{sheet_name}.pdf"
            pdf_path = os.path.join(savepath, pdf_filename)

            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                    time.sleep(1)
                except PermissionError:
                    print(f"❌ {pdf_path} 파일이 열려 있어 삭제할 수 없습니다. 닫고 다시 시도하세요.")
                    wb.Close(False)
                    continue

            try:
                print(f"저장시작!")
                print(f"저장시작!: {pdf_path}")
                wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
                wb.Close(False)
                excel.Quit()
            except Exception as e:
                print(f"❌ PDF 변환 중 오류 발생: {e}")
                wb.Close(False)
                continue

            print(f"✅ 저장 완료: {pdf_path}")

            i += 1
            wb.Close(False)

        excel.Quit()
        print("📄 변환이 완료되었습니다.")

    def is_file_open(self, file_path):
        """파일이 현재 사용 중인지 확인"""
        try:
            fh = open(file_path, "r+")
            fh.close()
        except IOError:
            return True  # 파일이 열려 있음
        return False  # 파일이 닫혀 있음
    

        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = WindowClass( )
    myWindow.show( )
    app.exec_( )