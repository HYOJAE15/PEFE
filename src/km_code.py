import os
from pathlib import Path
from win32com import client as win32
from openpyxl import load_workbook

def excel_to_pdf(excel_file, output_folder):
    # Excel 파일의 각 시트를 개별적인 PDF로 변환하는 함수
    excel = win32.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_file)
    
    # 각 시트를 개별적인 PDF로 저장
    for ws in wb.Sheets:
        ws.Visible = 1
        pdf_filename = f"{ws.Name}.pdf"
        pdf_output_path = os.path.join(output_folder, pdf_filename)
        ws.ExportAsFixedFormat(0, pdf_output_path)
    
    wb.Close()
    excel.Quit()

def convert_folder_to_pdf(folder_path):
    # 폴더 내의 모든 엑셀 파일을 PDF로 변환하는 함수
    folder = Path(folder_path)
    excel_files = list(folder.glob("*.xlsx"))
    
    # 폴더 내의 모든 엑셀 파일에 대해 변환 수행
    for excel_file in excel_files:
        output_folder = folder / excel_file.stem
        output_folder.mkdir(exist_ok=True)
        
        # 엑셀 파일의 각 시트를 개별적인 PDF로 변환
        excel_to_pdf(str(excel_file), str(output_folder))
        
        print(f"{excel_file}을(를) {output_folder} 폴더에 각 시트별로 PDF로 변환했습니다.")

# 폴더 경로 지정
folder_path = r"D:\구\23.06\230619 수량산출서\09.강천3교 수량산출서\3-1. 강천3교 교대 수량산출서\새 폴더\\"

# 폴더 내의 엑셀 파일을 PDF로 변환
convert_folder_to_pdf(folder_path)
