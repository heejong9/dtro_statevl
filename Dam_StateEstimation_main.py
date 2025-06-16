import DamFloor_StateEstimation
import pandas as pd
import os
import argparse
from datetime import datetime
 
"""
    상태평가 등급 생성기 
    
    1단계와 1.5단계 생성용도 모듈
    
    계산식은 모두  DamFloor_StateEstimation에서 작업
"""

def createDirectory(directory):
    """_summary_
        폴더 비존재시 생성
    Args:
        directory (_type_): _폴더경로_
    """
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

 
def  get_unique_filenames_without_extension(directory_path):
    """
        파일이름 목록 모음
    Args:
        directory_path (str): _결함 검출 폴더의 경로_

    Returns:
        set[str]: _결함검출폴더 내부의 파일이름 목록_
    """
    # 파일 이름 (확장자 제거) 저장을 위한 set 생성
    unique_filenames = set()
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
    # 지정된 경로의 파일들을 순회
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            # 파일 이름에서 확장자를 제거
            filename_without_extension = os.path.splitext(file)[0]
            if os.path.splitext(file)[1].lower() in image_extensions:
                allPath[filename_without_extension] = os.path.join(root, file)
                unique_filenames.add(filename_without_extension)  # set에 추가

    return unique_filenames
     


bujeaList = {
    "YSR" :"overflow",
    "DMR" : "DamFloor",
    "SRM" : "UpStream",
    "HRM" : "DownStream",
    "BYR" : "Subflow",
    "BYI" : "SubflowIN",
    "BYO" : "SubflowOUT",
    'CSA' : "CSA",
    'CSB' : "CSB",
    'CSC' : "CSC",
    'CSD' : "CSD",
     
}

def checkBujea(bujea):
    """" 각 부재 폴더 이름 생성 함수"""
    actions = {
        'HRM': ['1R_HRM','1Y_HRM','1B_HRM'],
        'DMR': ['2R_DMR','2Y_DMR','2B_DMR'],
        'SRM': ['3R_SRM','3Y_SRM','3B_SRM'],
        'CSA': ['4R_CSA','4Y_CSA','4B_CSA'],
        'CSB': ['4R_CSB','4Y_CSB','4B_CSB'],
        'CSC': ['4R_CSC','4Y_CSC','4B_CSC'],
        'CSD': ['4R_CSD','4Y_CSD','4B_CSD'],
        'YSR': ['5R_YSR','5Y_YSR','5B_YSR'],
        'BYR': ['6R_BYR','6Y_BYR','6B_BYR'],
        'BYI': ['6R_BYI','6Y_BYI','6B_BYI'],
        'BYO': ['6R_BYO','6Y_BYO','6B_BYO'],
    }
    return actions.get(bujea)

if __name__ == '__main__':
    #231228 1100
    parser = argparse.ArgumentParser(description="Read text from an image and filter by specific conditions")

    # parser.add_argument('--measure-csv', default='../measure/27SYD-5YSR.csv', help="검출 결과 measure 파일")
    parser.add_argument('--measure-csv', default='./SYD/YSR.csv', help="검출 결과 measure 파일")

    # # parser.add_argument('--measure-csv', help="검출 결과 measure 파일")
    parser.add_argument('--output-path', default='241007', help="상태 등급 output 경로")
    #
 

    parser.add_argument('--dam', default='27SYD', help="상태평가보고서 작성용 파일들")
    
    parser.add_argument('--bujea', default='YSR', help="부재 이름")
    parser.add_argument('--attribute', default='', help="부재 속성(필 = F 콘크리트 = C 대청댐 = DCD)")
    parser.add_argument('--shotingday', default='', help="촬영시간")
    parser.add_argument('--mainpath', default='', help="메인 경로")
    args = parser.parse_args()

    file = args.measure_csv
    outputpath = args.output_path
    damName = args.dam
    bujeaName =args.bujea

    attribute = args.attribute
    mainPath = args.mainpath
    shotingDay = args.shotingday
    # print('resultname', resultname)



 
 
   
  
    stage4 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage04\\04"
    stage5 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage05\\05"
    stage6 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage06\\06"
    bujeaFolderNames = checkBujea(bujeaName)
    allPath = {}
    fileName = set()
    ##PATH 노드 본체까지 C:\Users\user\Desktop\codelist\(스토리지 루트)  댐이름 부재이름 촬영일자(20250106) 'R, Y, B'  '외관조사망도 생성경로' '부분생성여부(0,1)'
    ## 1 6대결함 경로 2  6대 결함 아웃풋 3 6대결함 JSON 4 누수 경로 5 누수 아웃풋 6 누수 JSON 7 변형 경로 8 변형 아웃풋 9 변형원본 las파일 (4)
   # 10 변형 검출 LAS파일(5) 11 GCP CSV 경로(4) 12 머지 경로(6) 13 들어갈 데이터 14 댐이름 15 부재이름 16 부분생성 여부(1,0)
   # 시작 시간

    if 'C' == attribute:   
        fileName.update(get_unique_filenames_without_extension(stage5+bujeaFolderNames[0]))

    if 'F' == attribute:     
        fileName.update(get_unique_filenames_without_extension(stage5+bujeaFolderNames[1]))
        fileName.update(get_unique_filenames_without_extension(stage6+bujeaFolderNames[2]))
    if 'D' == attribute:
        fileName.update(get_unique_filenames_without_extension(stage5+bujeaFolderNames[0]))
        fileName.update(get_unique_filenames_without_extension(stage5+bujeaFolderNames[1]))
        fileName.update(get_unique_filenames_without_extension(stage6+bujeaFolderNames[2]))
      


    structurename =bujeaList.get(bujeaName)

    createDirectory(outputpath)
    
    # file = 'CJD_merge_test2.csv'
    csvData = pd.read_csv(file)
    totalDataFrame = []
    # print('--------------------------------------------')
    gradeScore = {5: 'a', 4: 'b', 3: 'c', 2: 'd', 1: 'e'}
    Data = csvData[csvData["structure"] == structurename]
    

    structureLen = len(Data)
  
    DamFloor_StateEstimation.step1Grade(structurename, Data, structureLen, totalDataFrame,attribute)


  

   # 여기 스테이션별 결함등급 필요
    if attribute == 'C':
        onecolums= ( 'StructureType', 'Unit',
            'crackWidthGrade', 'crackInfulenceCoefficient',
            'desquamationGrade', 'desquamationInfulenceCoefficient',
            'leakageGrade', 'leakageInfulenceCoefficient',
            'efflorescenceGrade', 'efflorescenceInfulenceCoefficient',
            'peelingGrade', 'peelingInfulenceCoefficient',
            'rebarExposureGrade', 'rebarExposureInfulenceCoefficient')
    elif attribute == 'F':
        onecolums= ('StructureType', 'Unit',
            'leakageGrade', 'leakageInfulenceCoefficient',
            'deformGrade', 'deformInfulenceCoefficient',)
    elif attribute == 'D':
        
        onecolums= ( 'StructureType', 'Unit',
            'crackWidthGrade', 'crackInfulenceCoefficient',
            'desquamationGrade', 'desquamationInfulenceCoefficient',
            'leakageGrade', 'leakageInfulenceCoefficient',
            'efflorescenceGrade', 'efflorescenceInfulenceCoefficient',
            'peelingGrade', 'peelingInfulenceCoefficient',
            'rebarExposureGrade', 'rebarExposureInfulenceCoefficient',
            'deformGrade', 'deformInfulenceCoefficient',)
   
    # totalFrame = pd.DataFrame(totalDataFrame, columns=('StructureType', 'Unit', 'crackWidthGrade', 'crackInfulenceCoefficient', 'contractionWidthGrade', 'contractionInfluenceCoefficeient', 'desquamationGrade', 'desquamationInfulenceCoefficient'))
    totalFrame = pd.DataFrame(totalDataFrame, columns=onecolums)
    # totalFrame.to_csv(str(outputpath)+'/1단계결과_231023.csv', index=False)   peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient

    totalFrame.to_csv(str(outputpath) + '/1단계결과.csv', index=False)
    stage1_result = os.path.join(outputpath, '1단계결과.csv')

    csvDataSub = pd.read_csv(file)
    totalDataFrameSub = []

    SubData = csvDataSub[csvDataSub["structure"] == structurename]
    

    structureLenSub = len(SubData)
    
    
    DamFloor_StateEstimation.stepSubGrade(structurename, SubData, totalDataFrameSub,attribute,list(fileName))
    
    if attribute == 'C':
        subcolums= ('Unit',
    'crackWidthGrade', 'crackArea','crackName','crackCount',
    'desquamationGrade', 'desquamationArea','desquamationName','desquamationCount',
    'leakageGrade', 'leakageArea','leakageName','leakageCount',
    'efflorescenceGrade', 'efflorescenceArea','efflorescenceName','efflorescenceCount',
    'peelingGrade', 'peelingArea', 'peelingName','peelingCount',
    'rebarExposureGrade','rebarExposureArea','rebarExposureName','reberExposureCount',
    'crackTotalGrade','desquamationTotalGrade','leakageTotalGrade',
    'efflorescenceTotalGrade','peelingTotalGrade','rebarExposureTotalGrade')
    elif attribute == 'F':
        subcolums= ('Unit',
    'leakageGrade', 'leakageArea','leakageName','leakageCount',
    'deformGrade','deformArea','deformName','deformCount',
    'leakageTotalGrade','deformTotalGrade')
    elif attribute == 'D':
        subcolums= ('Unit',
    'crackWidthGrade', 'crackArea','crackName','crackCount',
    'desquamationGrade', 'desquamationArea','desquamationName','desquamationCount',
    'leakageGrade', 'leakageArea','leakageName','leakageCount',
    'efflorescenceGrade', 'efflorescenceArea','efflorescenceName','efflorescenceCount',
    'peelingGrade', 'peelingArea', 'peelingName','peelingCount',
    'rebarExposureGrade','rebarExposureArea','rebarExposureName','reberExposureCount',
    'deformGrade','deformArea','deformName','deformCount',
    'crackTotalGrade','desquamationTotalGrade','leakageTotalGrade',
    'efflorescenceTotalGrade','peelingTotalGrade','rebarExposureTotalGrade','deformTotalGrade')
  
    totalFrameSub = pd.DataFrame(totalDataFrameSub, columns=subcolums)

    totalFrameSub.to_csv(str(outputpath) + '/1.5단계결과.csv', index=False)
    stage1_5_result = os.path.join(outputpath, '1.5단계결과.csv')

    # totalDataFrame2 = []
    # file2 = stage1_result
    # csvData2 = pd.read_csv(file2)
    
    

    