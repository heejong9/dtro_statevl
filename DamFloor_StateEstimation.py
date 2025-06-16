import numpy
import pandas as pd
import re 

typedict = {
   "overflow":"YSR",
    "subOverFlow" :"BYR",
    "DamFloor" : "DMR",
    "DownStream" : "HRM",
    "UpStream" : "SRM" 
}

def grade5Score(grade) :

    if grade is not None and grade < 1.1:
        return 'A'
    elif grade is not None and grade < 1.19 :
        return 'D'
    elif grade is not None and grade < 1.29:
        return 'C'
    elif grade is not None and grade < 1.49:
        return 'D'
    else:
        return 'E'



def gradeScore(grade) :

    if grade is not None and grade < 1.5:
        return 'e'
    elif grade is not None and grade < 2.5 :
        return 'd'
    elif grade is not None and grade < 3.5:
        return 'c'
    elif grade is not None and grade < 4.5:
        return 'D'
    elif grade is None or grade <= 5:
        return 'a'
    else:
        return 'e'

def adjustmentCoefficientScore(grade):

    if grade < 1.5 :
        return 6
    elif grade < 2.5 :
        return 6
    elif grade < 3.5:
        return 3
    elif grade < 4.5:
        return 2
    elif grade <= 5 :
        return 1

    else:
        return 1
    

def weighttScore(grade):

    if grade == 5 :
        return 1
    elif grade == 4 :
        return 1.1
    elif grade == 3:
        return 1.2
    elif grade == 2 :
        return 1.4
    elif grade <= 1 :
        return 2

def gradeName(grade):

    if grade == 5 :
        return 'A'
    elif grade == 4 :
        return 'D'
    elif grade == 3:
        return 'C'
    elif grade == 2 :
        return 'D'
    elif grade == 1 :
        return 'E'
    else:
        return 'F'

#결함정보
def getCrackInfo(crackAreaRate, crackWidth,flag=1):
    
    waterCrackTable = [
        # (crackAreaRate, crackWidth(mm), crackWidthGrade, crackInfulenceCoefficient) 수처리 구조물 균열
        (0.05, 0.1, 5, 1.0),
        (0.05, 0.3, 5, 1.0),
        (0.05, 0.5, 4, 1.1),
        (0.05, 1.0, 3, 1.2),
        (0.05, float('inf'), 2, 1.4),

        (0.2, 0.1, 5, 1.0),
        (0.2, 0.3, 4, 1.1),
        (0.2, 0.5, 3, 1.2),
        (0.2, 1.0, 2, 1.4),
        (0.2, float('inf'), 1, 2.0),

        (float('inf'), 0.1, 4, 1.1),
        (float('inf'), 0.3, 3, 1.2),
        (float('inf'), 0.5, 2, 1.4),
        (float('inf'), 1.0, 1, 2.0),
        (float('inf'), float('inf'), 1, 2.0)
    ]
   
    
    
   
    
    
    if crackAreaRate == 0 and flag==1:
        return None, None
     
    for area_max, width_max, grade, influence in waterCrackTable: #수처리인지 일반구조물인지 확인부탁
        if crackAreaRate <= area_max and crackWidth <= width_max:
            return grade, influence
    
    return None, None  # 기본값 설정
#누수정보
def getLeakageInfo(leakageAreaRate,flag=1):
    leakageTable = [
        # (leakAreaRate, leakageWidthGrade, leakageInfulenceCoefficient) 
         (0.01, 5, 1.0),   
        (0.05, 4, 1.1),       
        (0.1, 3, 1.2),
        (0.2, 2, 1.4),
        (float('inf'), 1, 2.0),
        
    ]
    if leakageAreaRate == 0 and flag==1:
        return None, None
    for area_max, grade, influence in leakageTable:
        if leakageAreaRate <= area_max:
            return grade, influence
#백태정보
def getEffloreInfo(effloreAreaRate,flag=1):
    effloreTable = [
        # (effloreAreaRate, effloreWidthGrade, effloreInfulenceCoefficient) 
        (0.01, 5, 1.0), 
        (0.05, 4, 1.1),       
        (0.1, 3, 1.3),
        (0.2, 2, 1.7),
        (float('inf'), 1, 3.0), 
       ]
    if effloreAreaRate == 0 and flag==1:
        return None, None
    for area_max, grade, influence in effloreTable: 
        if effloreAreaRate <= area_max:
            return grade, influence
#박리정보
def getPeelingInfo(peelingArea,flag=1):
    peelingTable = [
        # (peelingAreaRate, peelingWidthGrade, peelingInfulenceCoefficient) 
        (2500.0, 5, 1.0),
        (7500.0, 4, 1.1),       
        (15000.0, 3, 1.2),
        (30000.0, 2, 1.4),
        (float('inf'), 1, 2.0),    
       ]
    if peelingArea == 0 and flag==1:
        return None, None
    for area_max, grade, influence in peelingTable: 
        if peelingArea <= area_max:
            return grade, influence
#박락정보
def getDesquInfo(desquArea,flag=1):
    desquTable = [
        # (desquAreaRate, desquWidthGrade, desquInfulenceCoefficient) 
        (2500.0, 5, 1.0),
        (7500.0, 4, 1.1),       
        (15000.0, 3, 1.2),
        (30000.0, 2, 1.4),
        (float('inf'), 1, 2.0),      
       ]
  
    if desquArea == 0 and flag==1:
        return None, None
    for area_max, grade, influence in desquTable: 
        if desquArea <= area_max:
            return grade, influence
#철근노출정보
def getRebarExpInfo(rebarExpAreaRate,flag=1):
    rebarExpTable = [
        # (rebarExpAreaRate, rebarExpWidthGrade, rebarExpInfulenceCoefficient) 
        (0.05, 5, 1.0),
        (0.1, 4, 1.1),       
        (0.3, 3, 1.2),
        (0.5, 2, 1.4),
        (float('inf'), 1, 2.0),    
       ]
    if rebarExpAreaRate == 0 and flag==1:
        return None, None
    for area_max, grade, influence in rebarExpTable: 
        if rebarExpAreaRate <= area_max:
            return grade, influence
##침하등급
def getDeformInfo(deformAreaRate,flag=1):
    deformTable = [
        # (leakAreaRate, leakageWidthGrade, leakageInfulenceCoefficient) 
         (5, 5, 1.0),   
        (10, 4, 1.0),       
        (50, 3, 1.0),
        (100, 2, 1.0),
        (float('inf'), 1, 1.0),
        
    ]
    if deformAreaRate == 0 and flag==1:
        return None, None
    for area_max, grade, influence in deformTable:
        if deformAreaRate <= area_max:
            return grade, influence    
def step1Grade(structureName, data ,structureLen,totalDataFrame,attribute) :
    # print('Stage 1')
   
    ##스테이션 별 등급
    dataList = []
    for i in range(0, len(data)) :
       
        
            
        unit,crackWidthGrade, crackInfulenceCoefficient,  desquamationGrade, desquamationInfulenceCoefficient, \
        leakageGrade, leakageInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient, \
        peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient,deformGrade,deformInfulenceCoefficient = damFloodrOneStep(data[i:i+1])
    
        if attribute == 'C':
            dataList = [structureName,unit, crackWidthGrade, crackInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, \
                leakageGrade, leakageInfulenceCoefficient,  efflorescenceGrade, efflorescenceInfulenceCoefficient, \
                peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient]
        elif attribute == 'F':
            dataList = [structureName,unit, 
                leakageGrade, leakageInfulenceCoefficient, deformGrade,deformInfulenceCoefficient]
        elif attribute == 'D':
            dataList = [structureName,unit, crackWidthGrade, crackInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, \
                leakageGrade, leakageInfulenceCoefficient,  efflorescenceGrade, efflorescenceInfulenceCoefficient, \
                peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient,\
                deformGrade,deformInfulenceCoefficient   ]
        
        totalDataFrame.append(dataList)
   

def stepSubGrade(structureName, data ,totalDataFrame,attribute,fileList):
      # ##부재 별 등급
   
    # data['Station'] = data['Defect_ID'].apply(lambda x: x.split('-', 1)[1].rsplit('_', 1)[0])
    data['Station'] = data['Defect_ID'].apply(lambda x: '_'.join(re.sub(r'^U\d+-', '', x).split('_')[:-1]))
    ##스테이션 별 등급
    # for i in range(1,structureLen+1) :
    for station in fileList:
     
        if not data[data['Station'] == station].empty:
     
     
            unit,crackWidthGrade, crackInfulenceCoefficient,crackArea,crackName,  desquamationGrade, desquamationInfulenceCoefficient,desquamationArea,desquamationName, \
                leakageGrade, leakageInfulenceCoefficient,leakageArea,leakageName, efflorescenceGrade, efflorescenceInfulenceCoefficient,efflorescenceArea,efflorescenceName, \
                peelingGrade,peelingInfulenceCoefficient,PeelingArea,PeelingName, rebarExposureGrade,rebarExposureInfulenceCoefficient,REArea,REName, \
                deformGrade,deformInfulenceCoefficient,deformArea,deformName,crackCount,desquamationCount,leakageCount,efflorescenceCount,peelingCount,rebarExposureCount,deformCount= damFloodSubStep(data[data['Station'] == station])
          
        else:
          unit,crackWidthGrade, crackInfulenceCoefficient,crackArea,crackName,  desquamationGrade, desquamationInfulenceCoefficient,desquamationArea,desquamationName, \
                leakageGrade, leakageInfulenceCoefficient,leakageArea,leakageName, efflorescenceGrade, efflorescenceInfulenceCoefficient,efflorescenceArea,efflorescenceName, \
                peelingGrade,peelingInfulenceCoefficient,PeelingArea,PeelingName, rebarExposureGrade,rebarExposureInfulenceCoefficient,REArea,REName,deformGrade,\
                    deformInfulenceCoefficient,deformArea,deformName,crackCount,desquamationCount,leakageCount,efflorescenceCount,peelingCount,rebarExposureCount,deformCount   = [station,"A",0,0,'',"A",0,0,'',"A",0,0,'',"A",0,0,'',"A",0,0,'',"A",0,0,'',"A",0,0,'','0','0','0','0','0','0','0']
        if attribute == 'F':
             dataList = [unit, \
                leakageGrade, leakageArea,leakageName,leakageCount,  \
                deformGrade,deformArea,deformName,deformCount,\
                   '' ,'' ]
        elif attribute == 'C':
             dataList = [unit,crackWidthGrade, crackArea,crackName,crackCount,  desquamationGrade,desquamationArea,desquamationName,desquamationCount, \
                leakageGrade, leakageArea,leakageName,leakageCount, efflorescenceGrade, efflorescenceArea,efflorescenceName,efflorescenceCount, \
                peelingGrade,PeelingArea,PeelingName,peelingCount, rebarExposureGrade,REArea,REName,rebarExposureCount,\
                   '' ,'' ,'' ,'' ,'' , '',]
        elif attribute == 'D':
             dataList = [unit,crackWidthGrade, crackArea,crackName,crackCount,  desquamationGrade,desquamationArea,desquamationName,desquamationCount, \
                leakageGrade, leakageArea,leakageName,leakageCount, efflorescenceGrade, efflorescenceArea,efflorescenceName,efflorescenceCount, \
                peelingGrade,PeelingArea,PeelingName,peelingCount, rebarExposureGrade,REArea,REName,rebarExposureCount,deformGrade,deformArea,deformName,deformCount,\
                   '' ,'' ,'' ,'' ,'' , '','',]
        totalDataFrame.append(dataList)
def damFloodSubStep(csvData):
    dataLen = len(csvData)

    if dataLen != 0 :
        totalSpanArea = csvData["filter_pixel"] #m단위를 mm단위로 그리고 merge파일에서 뽑아쓸것

        REAreaRate,leakageAreaRate,crackAreaRate,crackWidth,efflorescenceAreaRate,PeelingArea,desquamationArea,deformAreaMax,defactlist,defectlistcount = areaRateDefectAll(csvData,totalSpanArea)
        ## 철근노출 파트
        rebarExposureGrade,rebarExposureInfulenceCoefficient = getRebarExpInfo(REAreaRate,0)
        
        ## 누수 파트
        leakageGrade,leakageInfulenceCoefficient = getLeakageInfo(leakageAreaRate,0)
       
        ## 균열 파트
        crackWidthGrade,crackInfulenceCoefficient = getCrackInfo(crackAreaRate,crackWidth,0)
        
        ## 백태 파트
        efflorescenceGrade,efflorescenceInfulenceCoefficient = getEffloreInfo(efflorescenceAreaRate,0)
        

        ## 박리 파트
        peelingGrade,peelingInfulenceCoefficient = getPeelingInfo(PeelingArea,0)
        

        ##박락 파트
        desquamationGrade,desquamationInfulenceCoefficient = getDesquInfo(desquamationArea,0) 
        
        ##변형 파트
        
        deformGrade,deformInfulenceCoefficient = getDeformInfo(deformAreaMax,0)   
    # print('step1', crackWidthGrade, crackInfulenceCoefficient,  desquamationGrade, desquamationInfulenceCoefficient, \
    #        leakageGrade, leakageInfulenceCoefficient,  efflorescenceGrade, efflorescenceInfulenceCoefficient, \
    #         peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient,deformGrade)
    
    csvData['Defect_ID'].apply(lambda x: "_".join(x.split('_')[0:3])).iloc[0]
    # 균열 , 균열 영향 계수
    def count_nonempty_elements(text):
        return len(list(filter(None, text.split(","))))
    crackCount = count_nonempty_elements(defectlistcount[1])
    desquamationCount = count_nonempty_elements(defectlistcount[4])
    leakageCount = count_nonempty_elements(defectlistcount[2])
    efflorescenceCount = count_nonempty_elements(defectlistcount[5])
    peelingCount = count_nonempty_elements(defectlistcount[3])
    rebarExposureCount = count_nonempty_elements(defectlistcount[0])
    deformCount = count_nonempty_elements(defectlistcount[6])
    
    return csvData['Station'].iloc[0],gradeName(crackWidthGrade), crackInfulenceCoefficient,defactlist[1],defectlistcount[1],  gradeName(desquamationGrade), desquamationInfulenceCoefficient,defactlist[4],defectlistcount[4], \
           gradeName(leakageGrade), leakageInfulenceCoefficient,defactlist[2],defectlistcount[2],  gradeName(efflorescenceGrade), efflorescenceInfulenceCoefficient,defactlist[5],defectlistcount[5], \
           gradeName(peelingGrade),peelingInfulenceCoefficient,defactlist[3],defectlistcount[3], gradeName(rebarExposureGrade),rebarExposureInfulenceCoefficient,defactlist[0],defectlistcount[0],\
           gradeName(deformGrade),deformInfulenceCoefficient,defactlist[6],defectlistcount[6],crackCount,desquamationCount,leakageCount,efflorescenceCount,peelingCount,rebarExposureCount,deformCount

## 댐마루의 총 등급 작성
def damFloodrOneStep(csvData)  -> object:
    dataLen = len(csvData)
    
   
  
    # print(csvData)
    if dataLen != 0 :
        totalSpanArea = csvData["filter_pixel"] #m단위 그리고 merge파일에서 뽑아쓸것

        REAreaRate,leakageAreaRate,crackAreaRate,crackWidth,efflorescenceAreaRate,PeelingArea,desquamationArea,deformAreaMax = areaRateDefect(csvData,totalSpanArea)
       
        ## 철근노출 파트
        rebarExposureGrade,rebarExposureInfulenceCoefficient = getRebarExpInfo(REAreaRate)
        ## 누수 파트
        leakageGrade,leakageInfulenceCoefficient = getLeakageInfo(leakageAreaRate)
       
        ## 균열 파트
        crackWidthGrade,crackInfulenceCoefficient = getCrackInfo(crackAreaRate,crackWidth)
      
        ## 백태 파트
        efflorescenceGrade,efflorescenceInfulenceCoefficient = getEffloreInfo(efflorescenceAreaRate)
       

        ## 박리 파트
        peelingGrade,peelingInfulenceCoefficient = getPeelingInfo(PeelingArea)

        ##박락 파트
        desquamationGrade,desquamationInfulenceCoefficient = getDesquInfo(desquamationArea)
       
       ##침하 파트
       
        deformGrade,deformInfulenceCoefficient = getDeformInfo(deformAreaMax)
       

 


    # print('step1', crackWidthGrade, crackInfulenceCoefficient,  desquamationGrade, desquamationInfulenceCoefficient, \
    #        leakageGrade, leakageInfulenceCoefficient,  efflorescenceGrade, efflorescenceInfulenceCoefficient, \
    #         peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient)
   
    # # 균열 , 균열 영향 계수
  
    return csvData['Defect_ID'].iloc[0],crackWidthGrade, crackInfulenceCoefficient,  desquamationGrade, desquamationInfulenceCoefficient, \
           leakageGrade, leakageInfulenceCoefficient,  efflorescenceGrade, efflorescenceInfulenceCoefficient, \
           peelingGrade,peelingInfulenceCoefficient, rebarExposureGrade,rebarExposureInfulenceCoefficient,deformGrade,deformInfulenceCoefficient

def step2Grade(structureName, data, structureLen, totalDataFrame,attribute,fileName):
    import re
    print('Stage 2')

    # Station 추출
    data['Station'] = data['Unit'].apply(lambda x: '_'.join(re.sub(r'^U\d+-', '', x).split('_')[:-1]))
    unique_stations = data['Station'].unique().tolist()

    # Station별 최소값 계산
    for station in fileName:
        # Station별 데이터 필터링
        unitData = data[data['Station'] == station]

        # Station별 최소값 계산 (열별 계산 후 전체 최소값)
        min_value = float('inf')  # 초기값을 무한대로 설정
        if attribute == 'C':
            for grade, coef in [
                ("crackWidthGrade", "crackInfulenceCoefficient"),
                ("desquamationGrade", "desquamationInfulenceCoefficient"),
                ("leakageGrade", "leakageInfulenceCoefficient"),
                ("peelingGrade", "peelingInfulenceCoefficient"),
                ("rebarExposureGrade", "rebarExposureInfulenceCoefficient"),
                ("efflorescenceGrade", "efflorescenceInfulenceCoefficient")
            ]:
                min_value = min(min_value, (unitData[grade] * unitData[coef]).min())
            # 각 결함 등급과 영향 계수를 곱한 최소값 계산
          
        elif attribute == 'F':
            for grade, coef in [
                ("leakageGrade", "leakageInfulenceCoefficient"),
                ("deformGrade", "deformInfulenceCoefficient"),
            ]:
                min_value = min(min_value, (unitData[grade] * unitData[coef]).min())
        elif attribute == 'D':
            for grade, coef in [
                ("crackWidthGrade", "crackInfulenceCoefficient"),
                ("desquamationGrade", "desquamationInfulenceCoefficient"),
                ("leakageGrade", "leakageInfulenceCoefficient"),
                ("peelingGrade", "peelingInfulenceCoefficient"),
                ("rebarExposureGrade", "rebarExposureInfulenceCoefficient"),
                ("efflorescenceGrade", "efflorescenceInfulenceCoefficient")
            ]:
            # 각 결함 등급과 영향 계수를 곱한 최소값 계산
                min_value = min(min_value, (unitData[grade] * unitData[coef]).min())
        # 등급 계산 및 결과 저장
        if pd.isna(min_value) or numpy.isinf(min_value) :
            dataList = [structureName, station, gradeScore(5), 5]
        else:
            dataList = [structureName, station, gradeScore(min_value), round(min_value, 1)]
        
        totalDataFrame.append(dataList)

def step3Grade(structureName, e2List, structureLen, totalDataFrame) :
    dataLen = len(e2List)
    if dataLen != 0 :
        # print('------------------------------')
        # print('Stage 3')

        # print(e2List)
        # print(structureLen)
        if structureLen==0:
            return
        importanceWeight = 100 / structureLen
        # print(structureName)
        # print('가중치:' + str(importanceWeight))
        denominator =0 # 분모 계산식 합
        numerator = 0 # 분자 계산식 합
        count = 1
        for e2 in e2List :
            adjustmentCoefficient = adjustmentCoefficientScore(e2)
            awCalc = importanceWeight * adjustmentCoefficient
            eawCalc = e2 * awCalc
            # print(structureName +'-'+str(count).zfill(3)+ ' E2 : '+ str(e2) + ' , A : ' + str(adjustmentCoefficient)+
            #       ' , W : '+str(importanceWeight) +' , aw : '+str(awCalc) +' , eaw : '+str(eawCalc) )
            denominator += awCalc
            numerator += eawCalc
            count +=1
            # print('numerator:' + str(numerator))

        if denominator != 0:
            e3 = round(numerator / denominator, 2)
            e3grade = gradeScore(e3)
        else:
            e3 = 1

        # print('1.복합부재의 상태 평가 지수(E3) 값 = %s'%(e3))
        # print('2.복합부재의 상태평가 등급 = %s등급'%(gradeScore(e3)))
        dataList = [structureName, e3, gradeScore(e3)]

        totalDataFrame.append(dataList)
        # print('3.분자합  = %s' % (numerator))
        # print('4.분모합 = %s' % (denominator))

        
# 4단계(5-52p)
def step4Grade(structureName, e3Max, scaleList, e3Min, txtList, totalDataFrame,e3SList):
    print('Stage 4')


    # e3Max = e3List
    # e3Min = e3List
    scaleSum = sum(scaleList)
    
    
    e3ScaleSum = e3SList

    v1 = round(0.3 * (e3Max - e3Min), 2)
    v2 = round(e3ScaleSum / (5 * scaleSum), 2)
    ec = round(e3Min + (v1 * v2), 2)

    # grade = gradeScore(ec)
    grade = gradeScore(ec)
    print('------------------------------')
    print(structureName + ' - 4단계 상태평가 결과')
    print('1.상태평가지수(E3) 최댓값(Max) = %s' % (e3Max))
    print('2.상태평가지수(E3) 최댓값(Min) = %s' % (e3Min))
    print('3.V1 = 0.3 × (Max -Min) = %s' % (v1))
    print('4.V2 = Σ(E3 × S) / (5 × ΣS)  = %s' % (v2))
    print('5.개별 시설물의 상태평가지수(Ec) = Min + V1 × V2 = %s' % (ec))
    print('6.개별시설물의 상태평가 등급 = %s등급' % (grade))

    txtList.append('------------------------------')
    txtList.append(structureName + ' - 4단계 상태평가 결과')
    txtList.append('1.상태평가지수(E3) 최댓값(Max) = %s' % (e3Max))
    txtList.append('2.상태평가지수(E3) 최댓값(Min) = %s' % (e3Min))
    txtList.append('3.V1 = 0.3 × (Max -Min) = %s' % (v1))
    txtList.append('4.V2 = Σ(E3 × S) / (5 × ΣS)  = %s' % (v2))
    txtList.append('5.개별 시설물의 상태평가지수(Ec) = Min + V1 × V2 = %s' % (ec))
    txtList.append('6.개별시설물의 상태평가 등급 = %s등급' % (grade))


    dataList = [structureName, ec, gradeScore(ec)]

    totalDataFrame.append(dataList)

def step5Grade(structureName, data, structureLen, totalDataFrame) :
    dataLen = len(data)
    if dataLen != 0 :
        # print('------------------------------')
        # print('Stage 3')

        # print(e2List)
        if structureLen==0:
            return
        efflorescenceTotal= leakageTotal=rebarExposureTotal=peelingTotal=desquaTotal=crackTotal =0
        
        for index,rows in data.iterrows():
            crackTotal += weighttScore(rows["crackWidthGrade"])
            desquaTotal += weighttScore(rows["desquamationGrade"])
            peelingTotal += weighttScore(rows["leakageGrade"])
            efflorescenceTotal += weighttScore(rows["efflorescenceGrade"])
            leakageTotal += weighttScore(rows["peelingGrade"])
            rebarExposureTotal += weighttScore(rows["rebarExposureGrade"])
                    
                

        # 각 행에서 값이 있는 부분 (NaN이 아닌 값) 개수 계산
        non_nan_counts = data.notnull().sum(axis=0)
 
        totalDataFrame.append(grade5(crackTotal,desquaTotal,peelingTotal,
                                     rebarExposureTotal,leakageTotal,efflorescenceTotal,non_nan_counts,structureName))
        # print('3.분자합  = %s' % (numerator))
        # print('4.분모합 = %s' % (denominator))
        
        ## 여기서 부재의 결함종류 등급도 나와야함
        
def grade5(crackTotal,desquaTotal,peelingTotal,rebarExposureTotal,leakageTotal,efflorescenceTotal,length,structureName):
    crackValue = crackTotal/length["crackWidthGrade"]
    desquValue = desquaTotal/length["desquamationGrade"]
    peelingValue =peelingTotal/length["leakageGrade"]
    reberExposureValue = rebarExposureTotal/length["rebarExposureGrade"]
    leakageValue = leakageTotal/length["leakageGrade"]
    efflorescenceValue = efflorescenceTotal/length["efflorescenceGrade"]
    
    crackGrade= grade5Score(crackValue)
    desquGrade = grade5Score(desquValue)
    peelingGrade = grade5Score(peelingValue)
    reberGrade = grade5Score(reberExposureValue)
    leakageGrade = grade5Score(leakageValue)
    efflorescenceGrade = grade5Score(efflorescenceValue)
    
    return (structureName,crackValue,crackGrade,desquValue,desquGrade,peelingValue,peelingGrade,reberExposureValue,reberGrade,leakageValue,leakageGrade,efflorescenceValue,efflorescenceGrade )
    


def areaRateDefect(csvData,totalSpanArea ):
 
    totalSpanArea=totalSpanArea.iloc[0]
  
    crackData = csvData[csvData['Type'] == "crack"]
    crackWidth = crackData['Defect_W'].max()
    crackArea = crackData['Defect_A']
    crackAreaSum = crackArea.sum()
    crackAreaRate = crackAreaSum / totalSpanArea 


    efflorescenceArea = csvData[csvData['Type'] == "efflor"]
    efflorescenceArea = efflorescenceArea['Defect_A']
    efflorescenceAreaSum = efflorescenceArea.sum()
    efflorescenceAreaRate = efflorescenceAreaSum / totalSpanArea 


    desquamationArea = csvData[csvData['Type'] == "desqu"]
    desquamationArea = desquamationArea['Defect_A']
    desquamationAreaSum = desquamationArea.sum()
    desquamationAreaRate = desquamationAreaSum / totalSpanArea 

    leakageArea = csvData[csvData['Type'] == "leakage"]
    leakageArea = leakageArea['Defect_A']
    leakageAreaSum = leakageArea.sum()
    leakageAreaRate = leakageAreaSum / totalSpanArea
    

    REArea = csvData[csvData['Type'] == "re"]
    REArea = REArea['Defect_A']
    REAreaSum = REArea.sum()
    REAreaRate = REAreaSum / totalSpanArea
    

    PeelingArea = csvData[csvData['Type'] == "peeling"]
    PeelingArea = PeelingArea['Defect_A']
    PeelingAreaSum = PeelingArea.sum()
    PeelingAreaSumRate = PeelingAreaSum / totalSpanArea
    
    DeformArea = csvData[csvData['Type'] == "deform"]
    
    DeformArea = DeformArea['Defect_D']
    DeformAreaSum = DeformArea.sum()
    DeformAreaSumRate = DeformAreaSum / totalSpanArea
    DeformAreaMax = DeformArea.max() 
    if  DeformAreaMax != DeformAreaMax:
        DeformAreaMax = 0
    return  REAreaRate,leakageAreaRate,crackAreaRate,crackWidth,efflorescenceAreaRate,PeelingAreaSum,desquamationAreaSum,DeformAreaMax
        
def areaRateDefectAll(csvData,totalSpanArea ):
    def calculateCountName(Data):
      
        lastParts =   Data['Defect_ID'].apply(lambda x: x.split("_")[-1])
        result = ",".join(lastParts)
        if result == 0:
            result = ''
        
        return result
    totalSpanArea=totalSpanArea.iloc[0]
    defectlistSum =[]
    # 균열 처리
    crackData = csvData[csvData['Type'] == "crack"]

    crackWidth = crackData['Defect_W'].dropna().max() if not crackData['Defect_W'].dropna().empty else 0
    crackArea = crackData['Defect_A']
    crackName = calculateCountName(crackData)
    
    crackAreaSum = crackArea.dropna().sum() if not crackArea.dropna().empty else 0
    crackAreaRate = crackArea.dropna().max() / totalSpanArea if not crackArea.dropna().empty else 0

    # 백태 처리
    efflorescenceArea = csvData[csvData['Type'] == "efflor"]['Defect_A']
    
    efflorescenceAreaSum =  efflorescenceArea.dropna().sum()  if not efflorescenceArea.dropna().empty else 0
    efflorescenceAreaMax  = efflorescenceArea.dropna().max() if not efflorescenceArea.dropna().empty else 0
    efflorescenceAreaRate = efflorescenceAreaMax / totalSpanArea if efflorescenceAreaMax else 0.0001
    efflorescenceName = calculateCountName(csvData[csvData['Type'] == "efflor"])
    # 박락 처리
    desquamationArea = csvData[csvData['Type'] == "desqu"]['Defect_A']
    desquamationAreaSum = desquamationArea.dropna().sum() if not desquamationArea.dropna().empty else 0
    desquamationAreaMax = desquamationArea.dropna().max() if not desquamationArea.dropna().empty else 0
    desquamationAreaRate = desquamationAreaMax / totalSpanArea if desquamationAreaMax else 0
    desquamationName = calculateCountName(csvData[csvData['Type'] == "desqu"])
    # 누수 처리
    leakageArea = csvData[csvData['Type'] == "leakage"]['Defect_A']
    leakageAreaSum = leakageArea.dropna().sum() if not leakageArea.dropna().empty else 0
    leakageAreaMax = leakageArea.dropna().max() if not leakageArea.dropna().empty else 0
    leakageAreaRate = leakageAreaMax / totalSpanArea if leakageAreaMax else 0.0001
    leakageName = calculateCountName(csvData[csvData['Type'] == "leakage"])
    # 철근 노출 처리
    REArea = csvData[csvData['Type'] == "re"]['Defect_A']
    REAreaSum = REArea.dropna().sum() if not REArea.dropna().empty else 0
    REAreaMax = REArea.dropna().max() if not REArea.dropna().empty else 0
    REAreaRate = REAreaMax / totalSpanArea if REAreaMax else 0.0001
  
    REName = calculateCountName(csvData[csvData['Type'] == "re"])
    # 박리 처리
    PeelingArea = csvData[csvData['Type'] == "peeling"]['Defect_A']
    PeelingAreaSum = PeelingArea.dropna().sum() if not PeelingArea.dropna().empty else 0
    PeelingAreaMax = PeelingArea.dropna().max() if not PeelingArea.dropna().empty else 0
    PeelingAreaSumRate = PeelingAreaMax / totalSpanArea if PeelingAreaMax else 0.0001
    PeelingName = calculateCountName(csvData[csvData['Type'] == "peeling"])
     # 침하 처리
    deformArea = csvData[csvData['Type'] == "deform"]['Defect_D']
    deformAreaSum = deformArea.dropna().sum() if not deformArea.dropna().empty else 0
    deformAreaMax = deformArea.dropna().max() if not deformArea.dropna().empty else 0
    deformAreaRate = deformAreaMax / totalSpanArea if deformAreaMax else 0.0001
    deformName = calculateCountName(csvData[csvData['Type'] == "deform"])
    # print( crackAreaSum,efflorescenceAreaSum,desquamationAreaSum,leakageAreaSum,REAreaSum,PeelingAreaSum)
    defectlistSum =[REAreaSum,crackAreaSum,leakageAreaSum,PeelingAreaSum,desquamationAreaSum,efflorescenceAreaSum,deformAreaSum]
    defectlistcount =[REName,crackName,leakageName,PeelingName,desquamationName,efflorescenceName,deformName]
    return  REAreaRate,leakageAreaRate,crackAreaRate,crackWidth,efflorescenceAreaRate,PeelingAreaMax,desquamationAreaMax,deformAreaMax,defectlistSum,defectlistcount
        
    