# -*- coding: utf-8 -*-


import pandas as pd
import os
import re
import csv
grade = {5: 'A', 4: 'B', 3: 'C', 2: 'D', 1: 'E'}
def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")




def numberlist(defect):
    if len(defect) ==0:
        return ''
    codex = defect['Defect_ID'].str.split('_').str[-1]
    # 문자열을 정수로 변환하여 정렬
    sorted_codex = sorted(codex.astype(int))

    # 리스트를 문자열로 변환 (쉼표로 구분)
    result_string = ','.join(map(str, sorted_codex))
    return result_string
def defectTable(structurename, file, dist,stage1_5_result,today,attribute,name):
    # print('손상물량표작성')
   

    csvData = pd.read_csv(file)
    # Defect_ID에서 '_' 앞부분 추출하여 새로운 컬럼 생성
    csvData = csvData.dropna(subset=['Defect_ID'])
    csvData['Prefix'] =csvData['Defect_ID'].apply(lambda x: '_'.join(re.sub(r'^U\d+-', '', x).split('_')[:-1]))
    
    # 특정 Prefix로 Defect_ID 검색
    unique_prefixes = csvData['Prefix'].unique().tolist()
   
    for prefix in name:
       
        if  stage1_5_result[stage1_5_result['Unit'] == prefix].empty:
            continue
        filtered_df = csvData[csvData['Prefix'] == prefix]
        
            
        OverflowData = filtered_df [filtered_df ["structure"] == structurename]
        # print('overflowdata_1', OverflowData)
        totalDataFrame = []

   

        DefectCount = len(OverflowData)

        ## crack
  

        

        
     
        if attribute == 'C':
            DefectCrack = OverflowData[OverflowData["Type"] == 'crack']
        
   
            cracksum = DefectCrack["Defect_A"]
            cracksum = round(cracksum.sum(), 3)

            datalist = ["1", str('균열'), str('㎡'), numberlist(DefectCrack), str(cracksum) + str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['crackTotalGrade'].values[0]]

            totalDataFrame.append(datalist)


            ## RE
            DefectRE = OverflowData[OverflowData["Type"] == 're']
            REsum = DefectRE["Defect_A"]
            REsum = round(REsum.sum(), 3)
            RECount = len(DefectRE)
            datalist = ['2', str('철근노출'), '㎡', numberlist(DefectRE), str(REsum) + str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['rebarExposureTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## Peeling
            Defectpeeling = OverflowData[OverflowData["Type"] == 'desqu']
            peelingsum = Defectpeeling["Defect_A"]
            peelingsum = round(peelingsum.sum(), 3)
            peelingCount = len(Defectpeeling)
            datalist = ['3', str('박락'), '㎡', numberlist(Defectpeeling), str(peelingsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['peelingTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## Desqu
            DefectDesqu = OverflowData[OverflowData["Type"] == 'peeling']
            defectsum = DefectDesqu["Defect_A"]
            defectsum = round(defectsum.sum(), 3)
            DesquCount = len(DefectDesqu)
            datalist = ['4', str('박리'), '㎡', numberlist(DefectDesqu), str(defectsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['desquamationTotalGrade'].values[0]]
            totalDataFrame.append(datalist)


            ## leakage
            Defectleakage = OverflowData[OverflowData["Type"] == 'leakage']
            leakagesum = Defectleakage["Defect_A"]
            leakagesum = round(leakagesum.sum(), 3)
            leakageCount = len(Defectleakage)
            datalist = ['5', str('누수'), '㎡',  numberlist(Defectleakage), str(leakagesum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['leakageTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## efflor
            Defectefflor = OverflowData[OverflowData["Type"] == 'efflor']
            efflorsum = Defectefflor["Defect_A"]
            efflorsum = round(efflorsum.sum(), 3)
            efflorCount = len(Defectefflor)
            datalist = ['6', str('백태'), '㎡', numberlist(Defectefflor), str(efflorsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['efflorescenceTotalGrade'].values[0]]
            totalDataFrame.append(datalist)
        elif attribute == 'F':
            
            ## leakage
            Defectleakage = OverflowData[OverflowData["Type"] == 'leakage']
            leakagesum = Defectleakage["Defect_A"]
            leakagesum = round(leakagesum.sum(), 3)
            leakageCount = len(Defectleakage)
            datalist = ['1', str('누수'), '㎡',  numberlist(Defectleakage), str(leakagesum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['leakageTotalGrade'].values[0]]
            totalDataFrame.append(datalist)
              ## leakage
            Defectdeform = OverflowData[OverflowData["Type"] == 'deform']
            deformsum = Defectdeform["Defect_A"]
            deformsum = round(deformsum.sum(), 3)
            deformCount = len(Defectdeform)
            datalist = ['2', str('패임'), '㎡',  numberlist(Defectdeform), str(deformsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['deformTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

        elif attribute == 'D':
            DefectCrack = OverflowData[OverflowData["Type"] == 'crack']
        
   
            cracksum = DefectCrack["Defect_L"]
            cracksum = round(cracksum.sum(), 3)

            datalist = ["1", str('균열'), str('mm'), numberlist(DefectCrack), str(cracksum) + str('mm'),stage1_5_result[stage1_5_result['Unit'] == prefix]['crackTotalGrade'].values[0]]

            totalDataFrame.append(datalist)


            ## RE
            DefectRE = OverflowData[OverflowData["Type"] == 're']
            REsum = DefectRE["Defect_A"]
            REsum = round(REsum.sum(), 3)
            RECount = len(DefectRE)
            datalist = ['2', str('철근노출'), '㎡', numberlist(DefectRE), str(REsum) + str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['rebarExposureTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## Peeling
            Defectpeeling = OverflowData[OverflowData["Type"] == 'peeling']
            peelingsum = Defectpeeling["Defect_A"]
            peelingsum = round(peelingsum.sum(), 3)
            peelingCount = len(Defectpeeling)
            datalist = ['3', str('박리'), '㎡', numberlist(Defectpeeling), str(peelingsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['peelingTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## Desqu
            DefectDesqu = OverflowData[OverflowData["Type"] == 'desqu']
            defectsum = DefectDesqu["Defect_A"]
            defectsum = round(defectsum.sum(), 3)
            DesquCount = len(DefectDesqu)
            datalist = ['4', str('박락'), '㎡', numberlist(DefectDesqu), str(defectsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['desquamationTotalGrade'].values[0]]
            totalDataFrame.append(datalist)


            ## leakage
            Defectleakage = OverflowData[OverflowData["Type"] == 'leakage']
            leakagesum = Defectleakage["Defect_A"]
            leakagesum = round(leakagesum.sum(), 3)
            leakageCount = len(Defectleakage)
            datalist = ['5', str('누수'), '㎡',  numberlist(Defectleakage), str(leakagesum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['leakageTotalGrade'].values[0]]
            totalDataFrame.append(datalist)

            ## efflor
            Defectefflor = OverflowData[OverflowData["Type"] == 'efflor']
            efflorsum = Defectefflor["Defect_A"]
            efflorsum = round(efflorsum.sum(), 3)
            efflorCount = len(Defectefflor)
            datalist = ['6', str('백태'), '㎡', numberlist(Defectefflor), str(efflorsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['efflorescenceTotalGrade'].values[0]]
            totalDataFrame.append(datalist)
            
            Defectdeform = OverflowData[OverflowData["Type"] == 'deform']
            deformsum = Defectdeform["Defect_A"]
            deformsum = round(deformsum.sum(), 3)
            deformCount = len(Defectdeform)
            datalist = ['7', str('패임'), '㎡',  numberlist(Defectdeform), str(deformsum)+str('㎡'),stage1_5_result[stage1_5_result['Unit'] == prefix]['deformTotalGrade'].values[0]]
            totalDataFrame.append(datalist)



 

        totalFrame = pd.DataFrame(totalDataFrame, columns=('구분', '주요결함 및 손상', '기준', '결함번호', '손상물량','등급'))

        # print(totalDataFrame)

     
        structurename2 = f'{prefix}'


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