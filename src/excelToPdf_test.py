import win32com.client
import openpyxl as op
import os


def excelInfo(filepath):
    excel_list = os.listdir(filepath) #폴더안에 잇는 엑셀파일 명을 리스트로 저장
    result = [] #빈 리스트 생성
    for file in excel_list: #엑셀파일명 리스트를 for문을 통해 반복
        wb = op.load_workbook(filepath+"/"+file) #openpyxl workbook 생성
        ws_list = wb.sheetnames #해당 workbook의 시트명을 리스트로 받음
        filename = file.replace(".xlsx","") #파일명을 저장하기 위해 문자열에서 확장자 제거
        for sht in ws_list: #시트명 리스트를 for문을 통해 반복
            temp_tuple = (filepath+"/"+file, filename, sht) #파일경로, 파일명, sht를 튜플에 저장
            result.append(temp_tuple) #위 튜플을 빈 리스트에 추가
    print(result)
    return result # 튜플로 이루어진 리스트 리턴


def transPDF(fileinfo, savepath):

    excel = win32com.client.Dispatch("Excel.Application")
    i=0 #파일명 중복을 방지하기 위한 인덱싱 번호

    #excelinfo를 입력받아 for문 실행
    for info in fileinfo:
        wb = excel.Workbooks.Open(info[0]) #info가 튜플이므로 인덱싱으로 접근(0번째는 파일경로)
        ws = wb.Worksheets(info[2]) #튜플의 2번째 요소는 시트명임. 

        ws.Select() #위 설정한 시트 선택
        wb.ActiveSheet.ExportAsFixedFormat(0, savepath+"/"+str(i)+"_"+info[1]+"_"+info[2]+".pdf") #파일명, 시트명으로 pdf 파일명 저장
        i=i+1
        wb.Close(False) #workbook 닫기, True일 경우 그 상태를 저장한다.
        excel.Quit()  # excel application 닫기


#실행부
if __name__ == "__main__":
    filepath = "D:/04. Test/pdf/2. 진입Ramp"
    pdfpath = "D:/04. Test/pdf/2. 진입Ramp/output"

    #excelinfo 실행(결과값은 튜플임)
    excelinfo = excelInfo(filepath)

    #transPDF 실행(실제 pdf파일로 변환하는 부분)
    transPDF(excelinfo, pdfpath)


