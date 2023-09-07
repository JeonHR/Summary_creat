## MADE by HR.Jeon 23.09.02

import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl

# CSV 파일이 있는 폴더 경로


folder_path = './'

# 폴더 내 모든 CSV 파일에 대해 반복합니다
for file_name in os.listdir(folder_path):
    if file_name.lower().endswith('.csv'): ## 모든 문자 소문자로 바꿈 대문자 CSV 로 입력 시 read가 안됨
        csv_path = os.path.join(folder_path, file_name)
        basename = os.path.basename(csv_path)
        name , exit = os.path.splitext(basename) ### File name 설정 exit는 확장자 명이라고 생각하면 됨

        # CSV 파일을 pandas DataFrame으로 읽어옵니다
        csv_data = pd.read_csv(csv_path)
        
        
        # 덮어쓰기를 원하는 엑셀 파일과 시트를 지정합니다.
        excel_file_path = './Tanami_FT1_summary.xlsx'
        sheet_name1 = 'happy'
        

        
        # 엑셀 파일을 열고 특정 시트를 가져옵니다.
        workbook = openpyxl.load_workbook(excel_file_path)
       
        sheet = workbook[sheet_name1]
        

        # 기존 시트의 내용을 모두 삭제합니다.
        for row in sheet:
            sheet.delete_rows(1, sheet.max_row)  ### 1행부터 지우기 시작함 --> 0으로 못한 이유는 csv data 자체가 깨지는 문제 확인
        
        # data Frame을 excel sheet 에 넣어주는 기능 --> 가장 시간이 오래걸린 부분으로 CSV data file을 excel sheet 에 넣어주는 기능
        for row in dataframe_to_rows(csv_data,index=True, header=False):
            sheet.append(row)
            

        # 변경사항을 저장하고 파일을 닫습니다.
        workbook.save(excel_file_path)
        workbook.close()


print("모든 CSV 파일을 처리하여 Excel 파일에 저장하였습니다.")
######################## Windows real save하는 것을 통해서 

import win32com.client as win32
import os
import time

# Excel 객체 생성
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Excel 창을 보이도록 설정 (False로 설정하면 숨김)

try:
    # 현재 작업 디렉토리를 기준으로 상대 경로를 절대 경로로 변환
    relative_path = "./Tanami_FT1_summary.xlsx"
    absolute_path = os.path.abspath(relative_path)
    

    # 열기
    wb = excel.Workbooks.Open(absolute_path)  # 상대 경로 대신 절대 경로를 사용

    # 원하는 작업 수행 (예: 데이터 수정)

    # 기존 파일에 덮어쓰기로 저장
    wb.Save()

    # 무조건 저장 (변경 사항이 없어도 저장)
    wb.Close(SaveChanges=True)

    # Excel 작업 완료 후 충분한 시간 대기 (예: 2초)
    time.sleep(2)

except Exception as e:
    print(f"오류 발생: {e}")
finally:
    # Excel 종료
    excel.Quit()




######################## Sheet --> TxT



# xlsx 파일 경로와 시트 이름 지정
sheet_name2 = 'Tanami_FT1' 


# xlsx 파일에서 시트 데이터를 데이터프레임으로 읽어옴
df = pd.read_excel(excel_file_path, sheet_name=sheet_name2) ## sheet name 을 불러오도록 하는 명령어


# 데이터프레임을 txt 파일로 저장
txt_file_path = f"{name}_Lotsummary.txt" ## CSV 생성되는 Name과 동일하게 사용
df.to_csv(txt_file_path, sep='\t', index=False,header=None)  
### Header None 값을 해주는 것을 통해서 첫 행의 열 없는 값 Unname을 지울 수 있음
### 마지막 구현하는 과정 중 오류가 발생한 부분으로는 colum data 에 영어 값이 하나도 없는 경우에 error 발생함
## excel sheet 열 기준으로 처음 확인되는 값이 열의 type을 결정함


print("시트를 {txt_file_path} 파일로 저장했습니다.")
