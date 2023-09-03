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
    if file_name.lower().endswith('.csv'):
        csv_path = os.path.join(folder_path, file_name)
        basename = os.path.basename(csv_path)
        name , exit = os.path.splitext(basename)

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
            sheet.delete_rows(1, sheet.max_row)
        
        # data Frame을 excel sheet 에 넣어주는 기능입니다.
        for row in dataframe_to_rows(csv_data,index=True, header=False):
            sheet.append(row)
            

        # 변경사항을 저장하고 파일을 닫습니다.
        workbook.save(excel_file_path)
        workbook.close()


print("모든 CSV 파일을 처리하여 Excel 파일에 저장하였습니다.")
######################## Windows real save

import win32com.client as win32

# Excel 객체 생성
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Excel 창을 보이도록 설정 (False로 설정하면 숨김)

try:
    # 열기
    wb = excel.Workbooks.Open("./Tanami_FT1_summary.xlsx")  # 엑셀 파일 경로 및 이름 설정

    # 저장
    wb.Save()

    # 닫기
    wb.Close(SaveChanges=True)  # SaveChanges=True로 설정하면 변경사항 저장, False로 설정하면 저장하지 않음
except Exception as e:
    print(f"오류 발생: {e}")
finally:
    # Excel 종료
    excel.Quit()





######################## Sheet --> TxT



# xlsx 파일 경로와 시트 이름 지정

sheet_name2 = 'Tanami_FT1'


# xlsx 파일에서 시트 데이터를 데이터프레임으로 읽어옴
df = pd.read_excel(excel_file_path, sheet_name=sheet_name2)
df2 = df.drop(['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 4','Unnamed: 5'],axis= 1)

# 데이터프레임을 txt 파일로 저장
txt_file_path = f"{name}_Lotsummary.txt"
df2.to_csv(txt_file_path, sep='\t', index=False)  # 탭으로 구분된 txt 파일로 저장 (sep='\t'를 사용하여 탭으로 구분)

print("시트를 {txt_file_path} 파일로 저장했습니다.")
