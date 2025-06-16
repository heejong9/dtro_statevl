import os
import pandas as pd
import numpy as np

# Leakage = 누수
# Fail = 손상
# Spalling = 박리
# Desquamation = 층분리 / 박락
# Efflorescence = 백태
# Segregation = 재료 분리
# RebarExposure = 철근 노출
# Crack Repair = 균열(보수 후)


def damFloodrOneStep(csvData) :

    dataLen = len(csvData)
    print('----DamFLoor Start----')

    #균열
    crackWidthGrade = 5
    crackInfulenceCoefficient = 1.0

    # 박리
    spallingGrade = 5
    spallingInfulenceCoefficient = 1.0

    # 박락
    desquamationGrade = 5
    desquamationInfulenceCoefficient = 1.0

    # 누수
    leakageGrade = 5
    leakageInfulenceCoefficient = 1.0

    # 파손 및 손상
    failGrade = 5
    failInfulenceCoefficient = 1.0

    #백태
    efflorescenceGrade = 5
    efflorescenceInfulenceCoefficient = 1.0



    if dataLen != 0 :
        totalSpanArea = 692.30

        crackWidth = csvData['Cr_W'].max() * 1000
        crackInfulenceCoefficient = 1.0
        crackArea = csvData['Cr_A']
        crackAreaSum = crackArea.sum()
        crackAreaRate = crackAreaSum / totalSpanArea

        failArea = csvData[csvData['Type'] == "Fail"]
        failDepth = csvData['Defect_Depth'].max() * 1000
        failInfulenceCoefficient = 1.0

        failArea = failArea['Defect_A']
        failAreaSum = failArea.sum()
        failAreaRate = failAreaSum / totalSpanArea




        print('crackAreaRate = ' + str(crackAreaRate))
        print('failAreaRate = ' + str(failAreaRate))


        if crackAreaRate  <= 5 :

            if crackWidth < 0.1:
                crackWidthGrade = 5

            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 5
            elif crackWidth < 0.5 :
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            else:
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2

        elif crackAreaRate  <= 20 :

            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.5 :
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2
            else:
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4

        else :

            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.3:
                crackWidthGrade = 3
                crackInfulenceCoefficient= 1.2
            elif crackWidth < 0.5 :
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4
            else:
                crackWidthGrade = 1
                crackInfulenceCoefficient = 2.0

        if failDepth == 0 :
            failGrade = 5

        elif failDepth < 20 and failAreaRate < 10 :
            failGrade = 4
            failInfulenceCoefficient = 1.1

        elif 20<= failDepth < 50  and failAreaRate < 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif failDepth <20  and failAreaRate > 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif 50<= failDepth <80  and failAreaRate < 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth < 50  and failAreaRate > 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7
        else :
            failGrade = 1
            failInfulenceCoefficient = 3.0

    # 균열 , 균열 영향 계수
            return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

    return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

def downStreamOneStep(csvData) :

    dataLen = len(csvData)
    print('----downStream Start----')

    #균열
    crackWidthGrade = 5
    crackInfulenceCoefficient = 1.0

    # 박리
    spallingGrade = 5
    spallingInfulenceCoefficient = 1.0

    # 박락
    desquamationGrade = 5
    desquamationInfulenceCoefficient = 1.0

    # 누수
    leakageGrade = 5
    leakageInfulenceCoefficient = 1.0

    # 파손 및 손상
    failGrade = 5
    failInfulenceCoefficient = 1.0

    #백태
    efflorescenceGrade = 5
    efflorescenceInfulenceCoefficient = 1.0



    if dataLen != 0 :
        totalSpanArea = 692.30

        crackWidth = csvData['Cr_W'].max() * 1000
        crackInfulenceCoefficient = 1.0
        crackArea = csvData['Cr_A']
        crackAreaSum = crackArea.sum()
        crackAreaRate = crackAreaSum / totalSpanArea

        failArea = csvData[csvData['Type'] == "Fail"]
        failDepth = csvData['Defect_Depth'].max() * 1000
        failInfulenceCoefficient = 1.0

        failArea = failArea['Defect_A']
        failAreaSum = failArea.sum()
        failAreaRate = failAreaSum / totalSpanArea

        efflorescenceArea = csvData[csvData['Type'] == "Efflorescence"]
        efflorescenceArea = efflorescenceArea['Defect_A']
        efflorescenceAreaSum = efflorescenceArea.sum()
        efflorescenceAreaRate = efflorescenceAreaSum / totalSpanArea


        spallingArea = csvData[csvData['Type'] == "Spalling"]
        spallingArea = spallingArea['Defect_A']
        spallingAreaSum = spallingArea.sum()
        spallingAreaRate = spallingAreaSum / totalSpanArea


        print('crackAreaRate = ' + str(crackAreaRate))
        print('failAreaRate = ' + str(failAreaRate))


        ## 균열 파트
        if crackAreaRate  <= 5 :

            if crackWidth < 0.1:
                crackWidthGrade = 5

            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 5
            elif crackWidth < 0.5 :
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            else:
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2

        elif crackAreaRate  <= 20 :
            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.5 :
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2
            else:
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4

        else :

            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.3:
                crackWidthGrade = 3
                crackInfulenceCoefficient= 1.2
            elif crackWidth < 0.5 :
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4
            else:
                crackWidthGrade = 1
                crackInfulenceCoefficient = 2.0

        ## 손상 파트
        if failDepth == 0 :
            failGrade = 5

        elif failDepth < 20 and failAreaRate < 10 :
            failGrade = 4
            failInfulenceCoefficient = 1.1

        elif 20<= failDepth < 50  and failAreaRate < 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif failDepth <20  and failAreaRate > 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif 50<= failDepth <80  and failAreaRate < 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth < 50  and failAreaRate > 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7
        else :
            failGrade = 1
            failInfulenceCoefficient = 3.0


        ## 백태 파트
        if efflorescenceAreaRate == 0:
            efflorescenceGrade = 5

        elif efflorescenceAreaRate < 5 :
            efflorescenceGrade = 4
            failInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 10:
            efflorescenceGrade = 3
            failInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 20:
            efflorescenceGrade = 2
            efflorescenceInfulenceCoefficient = 1.7
        else:
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 3.0

        ## 박리 파트

        if spallingAreaRate == 0:
            spallingGrade = 5

        elif spallingAreaRate < 5:
            spallingGrade = 4
            spallingInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 10:
            spallingGrade = 3
            spallingInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 20:
            spallingGrade = 2
            spallingInfulenceCoefficient = 1.7
        else:
            spallingGrade = 1
            spallingInfulenceCoefficient = 3.0



    # 균열 , 균열 영향 계수
            return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

    return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

def overflowOneStep(csvData):

    dataLen = len(csvData)
    print('----overflowOneStep Start----')


    #균열
    crackWidthGrade = 5
    crackInfulenceCoefficient = 1.0

    # 박리
    spallingGrade = 5
    spallingInfulenceCoefficient = 1.0

    # 박락
    desquamationGrade = 5
    desquamationInfulenceCoefficient = 1.0

    # 누수
    leakageGrade = 5
    leakageInfulenceCoefficient = 1.0

    # 파손 및 손상
    failGrade = 5
    failInfulenceCoefficient = 1.0

    #백태
    efflorescenceGrade = 5
    efflorescenceInfulenceCoefficient = 1.0



    if dataLen != 0 :
        totalSpanArea = 228.58

        crackData = csvData[csvData['Type'] == "Crack"]

        crackWidth = crackData['Defect_W'].max()
        crackInfulenceCoefficient = 1.0
        crackArea = crackData['Defect_A']
        crackAreaSum = crackArea.sum()
        crackAreaRate = crackAreaSum / totalSpanArea

        failArea = csvData[csvData['Type'] == "Fail"]
        failDepth = failArea['Defect_Depth'].max()
        failInfulenceCoefficient = 1.0

        failArea = failArea['Defect_A']
        failAreaSum = failArea.sum()
        failAreaRate = failAreaSum / totalSpanArea

        efflorescenceArea = csvData[csvData['Type'] == "Efflorescence"]
        efflorescenceArea = efflorescenceArea['Defect_A']
        efflorescenceAreaSum = efflorescenceArea.sum()
        efflorescenceAreaRate = efflorescenceAreaSum / totalSpanArea

        spallingArea = csvData[csvData['Type'] == "Spalling"]
        spallingArea = spallingArea['Defect_A']
        spallingAreaSum = spallingArea.sum()
        spallingAreaRate = spallingAreaSum / totalSpanArea

        desquamationArea = csvData[csvData['Type'] == "Desquamation"]
        desquamationArea2 = csvData[csvData['Type'] == "Segregation"]
        desquamationArea = pd.concat([desquamationArea, desquamationArea2])
        desquamationDepth = desquamationArea['Defect_Depth'].max()

        desquamationArea = desquamationArea['Defect_A']
        desquamationAreaSum = desquamationArea.sum()
        desquamationAreaRate = desquamationAreaSum / totalSpanArea

        leakageArea = csvData[csvData['Type'] == "Leakage"]
        leakageArea = leakageArea['Defect_A']
        leakageAreaSum = leakageArea.sum()
        leakageAreaRate = leakageAreaSum / totalSpanArea
        leakageAreaRate = leakageAreaRate * 100000
        print('crackAreaRate = ' + str(crackAreaRate))
        print('crackWidth = '+str(crackWidth))
        print('failAreaRate = ' + str(failAreaRate))
        print('efflorescenceAreaRate = ' + str(efflorescenceAreaRate))
        print('spallingAreaRate = ' + str(spallingAreaRate))
        print('failAreaRate = ' + str(failAreaRate))
        print('desquamationAreaRate = ' + str(desquamationAreaRate))
        print('leakageAreaRate = ' + str(leakageAreaRate))

        if leakageAreaRate < 1:
            leakageGrade = 5

        elif leakageAreaRate < 10:
            leakageGrade = 4

        elif leakageAreaRate < 20:
            leakageGrade = 3

        elif leakageAreaRate < 30:
            leakageGrade = 2

        elif leakageAreaRate >= 40:
            leakageGrade = 1


        ## 균열 파트
        if crackAreaRate  <= 5 :

            if crackWidth < 0.1:
                crackWidthGrade = 5

            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 5
            elif crackWidth < 0.5 :
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth >=0.5:
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2

        elif crackAreaRate  <= 20 :
            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 5
            elif crackWidth < 0.3:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.5 :
                crackWidthGrade = 3
                crackInfulenceCoefficient = 1.2
            elif crackWidth >= 0.5:
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4

        else :

            if crackWidth < 0.1:
                crackWidthGrade = 5
            elif crackWidth < 0.2:
                crackWidthGrade = 4
                crackInfulenceCoefficient = 1.1
            elif crackWidth < 0.3:
                crackWidthGrade = 3
                crackInfulenceCoefficient= 1.2
            elif crackWidth < 0.5 :
                crackWidthGrade = 2
                crackInfulenceCoefficient = 1.4

            elif crackWidth >= 0.5:
                crackWidthGrade = 1
                crackInfulenceCoefficient = 2.0

        ## 손상 파트
        if failDepth == 0 :
            failGrade = 5

        elif failDepth < 20 and failAreaRate < 10 :
            failGrade = 4
            failInfulenceCoefficient = 1.1

        elif 20<= failDepth < 50  and failAreaRate < 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif failDepth <20  and failAreaRate > 10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif 50<= failDepth <80  and failAreaRate < 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth < 50  and failAreaRate > 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth > 50  and failAreaRate > 10 :
            failGrade = 1
            failInfulenceCoefficient = 3.0


        ## 백태 파트
        if efflorescenceAreaRate == 0:
            efflorescenceGrade = 5

        elif efflorescenceAreaRate < 5 :
            efflorescenceGrade = 4
            failInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 10:
            efflorescenceGrade = 3
            failInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 20:
            efflorescenceGrade = 2
            efflorescenceInfulenceCoefficient = 1.7
        elif efflorescenceAreaRate > 20 :
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 3.0

        ## 박리 파트

        if spallingAreaRate == 0:
            spallingGrade = 5

        elif spallingAreaRate < 5:
            spallingGrade = 4
            spallingInfulenceCoefficient = 1.1

        elif spallingAreaRate < 10:
            spallingGrade = 3
            spallingInfulenceCoefficient = 1.3

        elif spallingAreaRate < 20:
            spallingGrade = 2
            spallingInfulenceCoefficient = 1.7
        elif spallingAreaRate > 20 :
            spallingGrade = 1
            spallingInfulenceCoefficient = 3.0

        #박락
        if desquamationDepth == 0 :
            desquamationGrade = 5

        elif desquamationDepth < 15 and desquamationAreaRate < 10 :
            desquamationGrade = 4
            desquamationInfulenceCoefficient = 1.1

        elif 15<= desquamationDepth < 20  and desquamationAreaRate < 10 :
            desquamationGrade = 3
            desquamationInfulenceCoefficient = 1.2

        elif desquamationDepth <15  and desquamationAreaRate > 10 :
            desquamationGrade = 3
            desquamationInfulenceCoefficient = 1.2

        elif 20<= desquamationDepth <25  and desquamationAreaRate < 10 :
            desquamationGrade = 2
            desquamationInfulenceCoefficient = 1.4

        elif desquamationDepth < 20  and desquamationAreaRate > 10 :
            desquamationGrade = 2
            desquamationInfulenceCoefficient = 1.4
        elif desquamationDepth > 20  and desquamationAreaRate < 10  :
            desquamationGrade = 1
            desquamationInfulenceCoefficient = 2.0

        print('a')

        if efflorescenceAreaRate == 0:
            efflorescenceGrade = 5

        elif efflorescenceAreaRate < 5 :
            efflorescenceGrade = 4
            failInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 10:
            efflorescenceGrade = 3
            failInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 20:
            efflorescenceGrade = 2
            efflorescenceInfulenceCoefficient = 1.7
        elif efflorescenceAreaRate > 20 :
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 3.0




    # 균열 , 균열 영향 계수
            return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

    return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient

def step1Grade(structureName, data ,structureLen,totalDataFrame) :
    print(structureLen)

    for i in range(1,structureLen+1) :
        print('--------------------------------------------')
        print(structureName,i)
        unitNum = i
        unitData = data[data["unit1"] == unitNum]

        crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient,\
        leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient\
            = overflowOneStep(unitData)

        dataList = [structureName,unitNum,crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient]
        totalDataFrame.append(dataList)

if __name__ == "__main__":


    structureTypeList = ['DamFloor',
                        'UpStream',
                        'DownStream',
                        'Footing',
                        'Gallery',
                        'Overflow',
                        'DissipatorSurfaceWaterway',
                        'DissipatorWaterPurificationPaper',
                        'WaterIntakeTower',
                        'SmallHydro'
                         ]

    structureCountList = {'DamFloor': 4,
                          'UpStream': 21,
                          'DownStream': 10,
                          'Footing': 10,
                          'Gallery': 13,
                          'Overflow': 24,
                          'DissipatorSurfaceWaterway': 8,
                          'DissipatorWaterPurificationPaper': 5,
                          'WaterIntakeTower': 10,
                          'SmallHydro': 3
                          }

    file = 'CJD_merge_test.csv'



    csvData = pd.read_csv(file)

    totalDataFrame = []
    print('--------------------------------------------')
    gradeScore = {5: 'a', 4: 'b', 3: 'c', 2: 'd', 1: 'e'}


    damfloorData = csvData[csvData["structure"] == 'damfloor']
    upStreamData = csvData[csvData["structure"] == 'upstream']
    downStreamData = csvData[csvData["structure"] == 'downstream']
    footingData = csvData[csvData["structure"] == 'footing']
    galleryData = csvData[csvData["structure"] == 'gallery']
    overflowData = csvData[csvData["structure"] == 'overflow']
    dissipatorsurfaceWaterwayData = csvData[csvData["structure"] == 'dissipatorsurfaceWaterway']
    dissipatorWaterPurificationPaperData = csvData[csvData["structure"] == 'dissipatorwaterpurificationPaper']
    waterIntakeTowerData = csvData[csvData["structure"] == 'waterIntaketower']
    smallHydroData = csvData[csvData["structure"] == 'smallHydro']

    #1단계

    # step1Grade(structureTypeList[0],damfloorData,structureCountList[structureTypeList[0]],totalDataFrame)
    # step1Grade(structureTypeList[1], upStreamData, structureCountList[structureTypeList[1]], totalDataFrame)
    # step1Grade(structureTypeList[2], downStreamData, structureCountList[structureTypeList[2]], totalDataFrame)
    # step1Grade(structureTypeList[3], footingData, structureCountList[structureTypeList[3]], totalDataFrame)
    # step1Grade(structureTypeList[4], galleryData, structureCountList[structureTypeList[4]], totalDataFrame)
    step1Grade(structureTypeList[5], overflowData, structureCountList[structureTypeList[5]], totalDataFrame)
    # step1Grade(structureTypeList[6], overflowData, structureCountList[structureTypeList[6]], totalDataFrame)
    # step1Grade(structureTypeList[7], dissipatorWaterPurificationPaperData, structureCountList[structureTypeList[7]], totalDataFrame)
    # step1Grade(structureTypeList[8], waterIntakeTowerData, structureCountList[structureTypeList[8]], totalDataFrame)
    # step1Grade(structureTypeList[9], smallHydroData, structureCountList[structureTypeList[9]], totalDataFrame)

    totalFrame = pd.DataFrame(totalDataFrame,columns=('StructureType', 'Unit',"crackWidthGrade","crackInfulenceCoefficient","spallingGrade","spallingInfulenceCoefficient","desquamationGrade","desquamationInfulenceCoefficient","leakageGrade","leakageInfulenceCoefficient","failGrade","failInfulenceCoefficient","efflorescenceGrade","efflorescenceInfulenceCoefficient" ))

    totalFrame.to_csv('1단계결과.csv', index=False)
