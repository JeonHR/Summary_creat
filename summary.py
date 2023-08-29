import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl

# CSV 파일이 있는 폴더 경로
folder_path = './'

# 폴더 내 모든 CSV 파일에 대해 반복합니다
for file_name in os.listdir(folder_path):
    if file_name.endswith('.CSV'):
        csv_path = os.path.join(folder_path, file_name)
        
        # CSV 파일을 pandas DataFrame으로 읽어옵니다
        csv_data = pd.read_csv(csv_path)
        
        # 덮어쓰기를 원하는 엑셀 파일과 시트를 지정합니다.
        excel_file_path = './Tanami_FT1_summary.xlsx'
        sheet_name = 'happy'
            

        # 엑셀 파일을 열고 특정 시트를 가져옵니다.
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook[sheet_name]
        

        # 기존 시트의 내용을 모두 삭제합니다.
        for row in sheet:
            sheet.delete_rows(1, sheet.max_row)
        
        for row in dataframe_to_rows(csv_data,index=False, header=True):
            sheet.append(row)
        

        # 변경사항을 저장하고 파일을 닫습니다.
        workbook.save(excel_file_path)
        workbook.close()


print("모든 CSV 파일을 처리하여 Excel 파일에 저장하였습니다.")
