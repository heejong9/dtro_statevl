import os
import pandas as pd
import numpy as np

def gradeScore(grade) :

    if grade < 1.5 :
        return 'e'
    elif grade < 2.5 :
        return 'd'
    elif grade < 3.5:
        return 'c'
    elif grade < 4.5:
        return 'b'
    elif grade <= 5 :
        return 'a'

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

def step3Grade(structureName, e2List, structureLen, totalDataFrame) :
    print('------------------------------')

    importanceWeight = 100 / structureLen
    print(structureName)
    denominator = 0 # 분모 계산식 합
    numerator =  0 # 분자 계산식 합
    count = 1
    for e2 in e2List :

        adjustmentCoefficient = adjustmentCoefficientScore(e2)

        awCalc = importanceWeight * adjustmentCoefficient
        eawCalc = e2 * awCalc
        print(structureName +'-'+str(count).zfill(3)+ ' E2 : '+ str(e2) + ' , A : ' + str(adjustmentCoefficient)+' , W : '+str(importanceWeight) +' , aw : '+str(awCalc) +' , eaw : '+str(eawCalc) )

        denominator += awCalc
        numerator += eawCalc
        count +=1

    e3 = round(numerator / denominator, 2)
    gradeScore(e3)

    print('1.복합부재의 상태 평가 지수(E3) 값 = %s'%(e3))
    print('2.복합부재의 상태평가 등급 = %s등급'%(gradeScore(e3)))
    dataList = [structureName, e3, gradeScore(e3)]

    totalDataFrame.append(dataList)
    # print('3.분자합  = %s' % (numerator))
    # print('4.분모합 = %s' % (denominator))


if __name__ == "__main__":

    structureTypeList = ['DamFloor',
                        'UpStream',
                        'DownStream',
                        'Footing',
                        'Gallery',
                        'Overflow',
                        'DissipatorSurfaceWaterway',
                        'DissipatorWaterPurificationPaper',
                        'PublicBridge',
                        'WaterIntakeTower',
                        'SmallHydro'
                         ]

    # structureCountList = {'DamFloor': 4,
    #                       'UpStream': 21,
    #                       'DownStream': 10,
    #                       'Footing': 10,
    #                       'Gallery': 13,
    #                       'Overflow': 24,
    #                       'DissipatorSurfaceWaterway': 13,
    #                       'DissipatorWaterPurificationPaper': 13,
    #                       'PublicBridge': 2,
    #                       'WaterIntakeTower': 10,
    #                       'SmallHydro': 3
    #                       }

    structureCountList = {'DamFloor': 0,
                          'UpStream': 0,
                          'DownStream': 0,
                          'Footing': 0,
                          'Gallery': 0,
                          'Overflow': 1,
                          'DissipatorSurfaceWaterway': 0,
                          'DissipatorWaterPurificationPaper': 0,
                          'WaterIntakeTower': 0,
                          'SmallHydro': 0
                          }

    file = 'KPD_230710_02.csv'
    saveFile = 'KPD_230710_03.csv'
    totalDataFrame = []

    csvData = pd.read_csv(file)

    damfloorData = csvData[csvData["StructureType"] == 'DamFloor'].Evalue.tolist()
    upStreamData = csvData[csvData["StructureType"] == 'UpStream'].Evalue.tolist()
    downStreamData = csvData[csvData["StructureType"] == 'DownStream'].Evalue.tolist()
    footingData = csvData[csvData["StructureType"] == 'Footing'].Evalue.tolist()
    galleryData = csvData[csvData["StructureType"] == 'Gallery'].Evalue.tolist()
    overflowData = csvData[csvData["StructureType"] == 'Overflow'].Evalue.tolist()
    dissipatorSurfaceWaterwayData = csvData[csvData["StructureType"] == 'DissipatorSurfaceWaterway'].Evalue.tolist()
    dissipatorWaterPurificationPaperData = csvData[csvData["StructureType"] == 'DissipatorWaterPurificationPaper'].Evalue.tolist()
    publicBridgeData = csvData[csvData["StructureType"] == 'PublicBridge'].Evalue.tolist()
    waterIntakeTowerData = csvData[csvData["StructureType"] == 'WaterIntakeTower'].Evalue.tolist()
    smallHydroData = csvData[csvData["StructureType"] == 'SmallHydro'].Evalue.tolist()


    # step3Grade(structureTypeList[0],damfloorData,structureCountList[structureTypeList[0]],totalDataFrame)
    # step3Grade(structureTypeList[1],upStreamData,structureCountList[structureTypeList[1]],totalDataFrame)
    # step3Grade(structureTypeList[2],downStreamData,structureCountList[structureTypeList[2]],totalDataFrame)
    # step3Grade(structureTypeList[3],footingData,structureCountList[structureTypeList[3]],totalDataFrame)
    # step3Grade(structureTypeList[4],galleryData,structureCountList[structureTypeList[4]],totalDataFrame)
    step3Grade(structureTypeList[5],overflowData,structureCountList[structureTypeList[5]],totalDataFrame)
    # step3Grade(structureTypeList[6],dissipatorSurfaceWaterwayData,structureCountList[structureTypeList[6]],totalDataFrame)
    # step3Grade(structureTypeList[7],dissipatorWaterPurificationPaperData,structureCountList[structureTypeList[7]],totalDataFrame)
    # step3Grade(structureTypeList[8],publicBridgeData,structureCountList[structureTypeList[8]],totalDataFrame)
    # step3Grade(structureTypeList[9],waterIntakeTowerData,structureCountList[structureTypeList[9]],totalDataFrame)
    # step3Grade(structureTypeList[10],smallHydroData,structureCountList[structureTypeList[10]],totalDataFrame)

    totalFrame = pd.DataFrame(totalDataFrame, columns=('StructureType', 'Evalue', 'Grade'))
    totalFrame.to_csv(saveFile, index=False)