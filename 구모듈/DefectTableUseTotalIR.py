# -*- coding: utf-8 -*-


import pandas as pd
import os
import csv
grade = {5: 'A', 4: 'B', 3: 'C', 2: 'D', 1: 'E'}
def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

def numberlist(defect):
    codex = defect['Defect_ID'].str.split('_').str[-1]
    # 문자열을 정수로 변환하여 정렬
    sorted_codex = sorted(codex.astype(int))

    # 리스트를 문자열로 변환 (쉼표로 구분)
    result_string = ','.join(map(str, sorted_codex))
    return result_string
def defectTable(structurename, file, dist,stage1_5_result,today):
    # print('손상물량표작성')
  
    csvData = pd.read_csv(file)
    # Defect_ID에서 '_' 앞부분 추출하여 새로운 컬럼 생성
    csvData['Prefix'] = csvData['Defect_ID'].apply(lambda x: '_'.join(x.split('_')[:-1]))

    # 특정 Prefix로 Defect_ID 검색
    unique_prefixes = csvData['Prefix'].unique().tolist()
   
    for prefix in unique_prefixes:
       
        if  stage1_5_result[stage1_5_result['Unit'] == prefix].empty:
            continue
        filtered_df = csvData[csvData['Prefix'] == prefix]
        
            
        OverflowData = filtered_df [filtered_df ["structure"] == structurename]
        # print('overflowdata_1', OverflowData)
        totalDataFrame = []

   

      
        


        ## leakage
        Defectleakage = OverflowData[OverflowData["Type"] == 'leakage']
        leakagesum = Defectleakage["Defect_A"]
        leakagesum = round(leakagesum.sum(), 3)
        leakageCount = len(Defectleakage)
        datalist = ['1', str('누수'), '㎡',  numberlist(Defectleakage), str(leakagesum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['leakageTotalGrade'].values[0]]
        totalDataFrame.append(datalist)

    





 

        totalFrame = pd.DataFrame(totalDataFrame, columns=('구분', '주요결함 및 손상', '기준', '수량', '손상물량','등급'))

        # print(totalDataFrame)

        a,postfix= prefix.split('-',1)
        structurename2 = f'{postfix}'


        totalFrame.to_csv(os.path.join(dist, str(structurename2)+'.csv'), encoding='cp949', index=False)
        # totalFrame.to_excel(os.path.join(dist, str(structurename2)+'.xlsx'),  index=False)

    # totalFrame.to_csv(os.path.join(dist, str(structurename) + 'table.csv'))
    # totalFrame.to_excel(os.path.join(dist, str(structurename) + 'table.xlsx'))




# if __name__ == '__main__':
#     outputpath = 'table_dist_231221'
#     defectTable('Overflow', './특구/27SYD-5YSR-231205_일부.csv', outputpath)
#     # reporttable(defectTable, outputpath)

#     csv_file_path = 'merged_file.csv'
#     docx_file_name = "output_document.docx"


#     # DOCX 문서 생성
#     doc = Document()

#     # CSV 파일을 열고 내용을 DOCX 표로 추가
#     with open(csv_file_path, "r", encoding="cp949", errors="replace", newline="") as csv_file:
#         csv_reader = csv.reader(csv_file)

#         # 첫 번째 행에서 최대 열 수 추출
#         max_columns = max(len(row) for row in csv_reader)
#         print(max_columns)

#         # max_rows = max(len(col) for col in csv_reader)
#         # print(max_rows)

#         # DOCX 표 생성
#         table = doc.add_table(rows=1, cols=max_columns)

#         # 첫 번째 행에 표 제목 추가
#         table.rows[0].cells[0].text = "CSV 데이터"

#         # CSV 파일 내용을 DOCX 표로 복사
#         for row in csv_reader:
#             cells = table.add_row().cells  # 새로운 행 추가
#             for i, cell_value in enumerate(row):
#                 cells[i].text = cell_value

#     # DOCX 파일 저장
#     doc.save(docx_file_name)