import os
import pandas as pd
import numpy as np

def gradeScore(grade) :

    if grade < 1.5 :
        return 'E'
    elif grade < 2.5 :
        return 'D'
    elif grade < 3.5:
        return 'C'
    elif grade < 4.5:
        return 'B'
    elif grade <= 5 :
        return 'A'

def step4Grade(structureName,e3List,scaleList,e3SList,txtList) :

    e3Max = max(e3List)
    e3Min = min(e3List)
    scaleSum = sum(scaleList)
    e3ScaleSum = sum(e3SList)

    v1 = round(0.3 * (e3Max - e3Min),2)
    v2 = round(e3ScaleSum / (5*scaleSum),2)
    ec = round(e3Min + (v1 * v2),2)

    grade = gradeScore(ec)
    print('------------------------------')
    print(structureName+ ' - 4단계 상태평가 결과')
    print('1.상태평가지수(E3) 최댓값(Max) = %s'%(e3Max))
    print('2.상태평가지수(E3) 최댓값(Min) = %s'%(e3Min))
    print('3.V1 = 0.3 × (Max -Min) = %s'%(v1))
    print('4.V2 = Σ(E3 × S) / (5 × ΣS)  = %s'%(v2))
    print('5.개별 시설물의 상태평가지수(Ec) = Min + V1 × V2 = %s'%(ec))
    print('6.개별시설물의 상태평가 등급 = %s등급'%(grade))


    txtList.append('------------------------------')
    txtList.append(structureName+ ' - 4단계 상태평가 결과')
    txtList.append('1.상태평가지수(E3) 최댓값(Max) = %s'%(e3Max))
    txtList.append('2.상태평가지수(E3) 최댓값(Min) = %s'%(e3Min))
    txtList.append('3.V1 = 0.3 × (Max -Min) = %s'%(v1))
    txtList.append('4.V2 = Σ(E3 × S) / (5 × ΣS)  = %s'%(v2))
    txtList.append('5.개별 시설물의 상태평가지수(Ec) = Min + V1 × V2 = %s'%(ec))
    txtList.append('6.개별시설물의 상태평가 등급 = %s등급'%(grade))


if __name__ == "__main__":
    #본댐 제체(댐마루, 상류면, 하류면 비월류부)
    #본댐 푸팅부(푸팅부)
    #본댐 갤러리(갤러리)
    #여수로 월류부(월류부)
    #여수로 감세공(감세공(측수로식),감세공(정수지식)
    #교량 (공도교 및 유지관리 교량)
    #취수탑 (취수탑)
    #소수력 발전소(소수력 발전소)

    #region 규모
    structureScaleList = {'DamFloor': 904.8,
                          'UpStream': 8839.6,
                          'DownStream': 1920.2,
                          'Footing': 1728.9,
                          'Gallery': 2185.9,
                          'Overflow': 5486.0,
                          'DissipatorSurfaceWaterway': 2446.0,
                          'DissipatorWaterPurificationPaper': 1625.0,
                          'PublicBridge': 166.0,
                          'WaterIntakeTower': 58.0,
                          'SmallHydro': 121.0
                          }

    #endregion

    inputFile = 'KPD_230710_03.csv'
    saveFile = 'KPD_230710_04' + '.txt'
    csvData = pd.read_csv(inputFile)
    txtList = []

    #
    #region 평가지수
    # damFloorE3 = 4.49
    # upStreamE3 = 3.67
    # downStreamE3 = 4.41
    # footingE3 = 4.55
    # galleryE3 = 4.20
    # overflowE3 = 3.60
    # dissipatorSurfaceWaterwayE3 = 3.60
    # dissipatorWaterPurificationPaperE3 = 3.51
    # publicBridgeE3 = 3.87
    # waterIntakeTowerE3 = 4.80
    # smallHydroE3 = 4.00

    # damFloorE3 = csvData[csvData["StructureType"] == 'DamFloor'].Evalue.tolist()[0]
    # upStreamE3 = csvData[csvData["StructureType"] == 'UpStream'].Evalue.tolist()[0]
    # downStreamE3 = csvData[csvData["StructureType"] == 'DownStream'].Evalue.tolist()[0]
    # footingE3 = csvData[csvData["StructureType"] == 'Footing'].Evalue.tolist()[0]
    # galleryE3 = csvData[csvData["StructureType"] == 'Gallery'].Evalue.tolist()[0]
    overflowE3 = csvData[csvData["StructureType"] == 'Overflow'].Evalue.tolist()[0]
    # dissipatorSurfaceWaterwayE3 = csvData[csvData["StructureType"] == 'DissipatorSurfaceWaterway'].Evalue.tolist()[0]
    # dissipatorWaterPurificationPaperE3 = csvData[csvData["StructureType"] == 'DissipatorWaterPurificationPaper'].Evalue.tolist()[0]
    # publicBridgeE3 = csvData[csvData["StructureType"] == 'PublicBridge'].Evalue.tolist()[0]
    # waterIntakeTowerE3 = csvData[csvData["StructureType"] == 'WaterIntakeTower'].Evalue.tolist()[0]
    # smallHydroE3 = csvData[csvData["StructureType"] == 'SmallHydro'].Evalue.tolist()[0]

    #endregion 평가지수

    #region 규모
    damFloorScale = 904.8
    upStreamScale = 8839.6
    downStreamScale = 1920.2
    footingScale = 1728.9
    galleryScale = 2185.9
    overflowScale = 5486.0
    dissipatorSurfaceWaterwayScale = 2446.0
    dissipatorWaterPurificationPaperScale = 1625.0
    publicBridgeScale = 166.0
    waterIntakeTowerScale = 58.0
    smallHydroScale = 121.0
    #endregion 규모

    #region E3 * S 계산값
    # damFloorE3S = damFloorE3 * damFloorScale
    # upStreamE3S = upStreamE3 * upStreamScale
    # downStreamE3S = downStreamE3 * downStreamScale
    # footingE3S = footingE3 * footingScale
    # galleryE3S = galleryE3 * galleryScale
    overflowE3S = overflowE3 * overflowScale
    # dissipatorSurfaceWaterwayE3S = dissipatorSurfaceWaterwayE3 * dissipatorSurfaceWaterwayScale
    # dissipatorWaterPurificationPaperE3S = dissipatorWaterPurificationPaperE3 * dissipatorWaterPurificationPaperScale
    # publicBridgeE3S = publicBridgeE3 * publicBridgeScale
    # waterIntakeTowerE3S = waterIntakeTowerE3 * waterIntakeTowerScale
    # smallHydroE3S = smallHydroE3 * smallHydroScale
    #endregion

    #region 상태평가 4단계 계산
    # step4Grade("본댐 제체",[damFloorE3,upStreamE3,downStreamE3],[damFloorScale,upStreamScale,downStreamScale],[damFloorE3S,upStreamE3S,downStreamE3S],txtList)
    #
    # step4Grade("본댐 푸팅부",[footingE3],[footingScale],[footingE3S],txtList)
    #
    # step4Grade("본댐 갤러리",[galleryE3],[galleryScale],[galleryE3S],txtList)
    #
    step4Grade("여수로 월류부",[overflowE3],[overflowScale],[overflowE3S],txtList)
    #
    # step4Grade("여수로 감세공",[dissipatorSurfaceWaterwayE3,dissipatorWaterPurificationPaperE3],[dissipatorSurfaceWaterwayScale,dissipatorWaterPurificationPaperScale],[dissipatorSurfaceWaterwayE3S,dissipatorWaterPurificationPaperE3S],txtList)
    #
    # step4Grade("교량",[publicBridgeE3],[publicBridgeScale],[publicBridgeE3S],txtList)
    #
    # step4Grade("취수탑",[waterIntakeTowerE3],[waterIntakeTowerScale],[waterIntakeTowerE3S],txtList)
    #
    # step4Grade("소수력 발전소",[smallHydroE3],[smallHydroScale],[smallHydroE3S],txtList)

    # step4Grade("본댐 제체",[damFloorE3],[damFloorScale],[damFloorE3S],txtList)


    #endregion 상태평가 4단계 계산

    with open(saveFile,'w') as f :

        for i in txtList :
            f.write(i+'\n')
