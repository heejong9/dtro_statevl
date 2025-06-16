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

def step2Grade(structureName, data ,structureLen,totalDataFrame) :

    for i in range(1,structureLen+1) :
        print('--------------------------------------------')

        unitNum = i
        unitData = data[data["Unit"] == unitNum]

        print(structureName,unitNum)

        crackGrade = float(unitData["crackWidthGrade"] * unitData["crackInfulenceCoefficient"])
        spallingGrade = float(unitData["spallingGrade"] * unitData["spallingInfulenceCoefficient"])
        desquamationGrade = float(unitData["desquamationGrade"] * unitData["desquamationInfulenceCoefficient"])
        leakageGrade = float(unitData["leakageGrade"] * unitData["leakageInfulenceCoefficient"])
        failGrade = float(unitData["failGrade"] * unitData["failInfulenceCoefficient"])
        efflorescenceGrade = float(unitData["efflorescenceGrade"] * unitData["efflorescenceInfulenceCoefficient"])

        minValue = min(crackGrade,spallingGrade,desquamationGrade,leakageGrade,failGrade,efflorescenceGrade)
        print('Unit : ' + str(i))
        print(gradeScore(minValue))
        print(minValue)
        dataList = [structureName,unitNum,gradeScore(minValue),minValue]

        totalDataFrame.append(dataList)



if __name__ == "__main__":

    file = 'KPD_230710_01.csv'

    csvData = pd.read_csv(file)
    totalDataFrame = []
    print('--------------------------------------------')

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

    # structureCountList = {'DamFloor': 4,
    #                       'UpStream': 21,
    #                       'DownStream': 10,
    #                       'Footing': 10,
    #                       'Gallery': 13,
    #                       'Overflow': 24,
    #                       'DissipatorSurfaceWaterway': 8,
    #                       'DissipatorWaterPurificationPaper': 5,
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

    damfloorData = csvData[csvData["StructureType"] == 'DamFloor']
    upStreamData = csvData[csvData["StructureType"] == 'UpStream']
    downStreamData = csvData[csvData["StructureType"] == 'DownStream']
    footingData = csvData[csvData["StructureType"] == 'Footing']
    galleryData = csvData[csvData["StructureType"] == 'Gallery']
    overflowData = csvData[csvData["StructureType"] == 'Overflow']
    dissipatorSurfaceWaterwayData = csvData[csvData["StructureType"] == 'DissipatorSurfaceWaterway']
    dissipatorWaterPurificationPaperData = csvData[csvData["StructureType"] == 'DissipatorWaterPurificationPaper']
    waterIntakeTowerData = csvData[csvData["StructureType"] == 'WaterIntakeTower']
    smallHydroData = csvData[csvData["StructureType"] == 'SmallHydro']

    step2Grade(structureTypeList[0], damfloorData,structureCountList[structureTypeList[0]],totalDataFrame)
    step2Grade(structureTypeList[1], upStreamData, structureCountList[structureTypeList[1]], totalDataFrame)
    step2Grade(structureTypeList[2], downStreamData, structureCountList[structureTypeList[2]], totalDataFrame)
    step2Grade(structureTypeList[3], footingData, structureCountList[structureTypeList[3]], totalDataFrame)
    step2Grade(structureTypeList[4], galleryData, structureCountList[structureTypeList[4]], totalDataFrame)
    step2Grade(structureTypeList[5], overflowData, structureCountList[structureTypeList[5]], totalDataFrame)

    step2Grade(structureTypeList[6], dissipatorSurfaceWaterwayData, structureCountList[structureTypeList[6]], totalDataFrame)
    step2Grade(structureTypeList[7], dissipatorWaterPurificationPaperData, structureCountList[structureTypeList[7]], totalDataFrame)

    step2Grade(structureTypeList[8], waterIntakeTowerData, structureCountList[structureTypeList[8]], totalDataFrame)
    step2Grade(structureTypeList[9], smallHydroData, structureCountList[structureTypeList[9]], totalDataFrame)

    # tmpDataList = ["PublicBridge", 1, 'b', 3.79]
    # tmp2DataList = ["PublicBridge", 2, 'b', 4.19]
    #
    # totalDataFrame.append(tmpDataList)
    # totalDataFrame.append(tmp2DataList)

    totalFrame = pd.DataFrame(totalDataFrame, columns=('StructureType', 'Unit', 'Grade', 'Evalue'))
    totalFrame.to_csv('KPD_230710_02.csv', index=False)

