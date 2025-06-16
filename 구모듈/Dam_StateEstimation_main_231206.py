
import DamFloor_StateEstimation_231206
import Dam_StateEstimator_Report
import DefectTable

import pandas as pd
import os
import argparse

import json
from datetime import datetime


def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

    print(outputpath)


def mergetable(tablepath, outputpath):
    file_paths = [os.path.join(tablepath, file) for file in os.listdir(tablepath) if file.endswith(".xlsx")]
    merged_data = pd.read_excel(file_paths[0], engine='openpyxl')
    for file_path in file_paths[1:]:
        data_to_append = pd.read_excel(file_path)
        merged_data = merged_data.append(data_to_append, ignore_index=True)

    merged_data.to_csv(os.path.join(outputpath, "merged_file.csv"), index=False)
    # merged_data.to_excel(os.path.join(outputpath, "merged_file.xlsx"), index=False)
#

def csvmerge(tablepath, outputpath):
    # 모든 CSV 파일을 저장할 빈 데이터프레임 생성
    merged_df = pd.DataFrame()

    # 폴더 내 CSV 파일 목록 가져오기
    csv_files = [f for f in os.listdir(tablepath) if f.endswith('.csv')]

    print(csv_files)

    # 각 CSV 파일을 읽어서 병합
    for csv_file in csv_files:
        file_path = os.path.join(tablepath, csv_file)
        print(file_path)
        df = pd.read_csv(file_path, encoding='cp949')  # CSV 파일 읽기
        # df = pd.read_csv(file_path)  # CSV 파일 읽기


        print(df)
        merged_df = pd.concat([merged_df, df], ignore_index=True)
        print(type(merged_df))


    # 병합된 데이터프레임을 CSV 파일로 저장
    merged_df.to_csv(os.path.join(outputpath, "merged_file.csv"), encoding='cp949', index=False)

    # file_list = os.listdir(tablepath)
    # file_extension = '.csv'
    #
    # for file_name in file_list:
    #     if file_name.endswith(file_extension):
    #         file_path = os.path.join(tablepath, file_name)
    #         os.remove(file_path)  # 파일 삭제
    #
    # print(f"{file_extension} 확장자를 가진 파일을 모두 삭제했습니다.")
    #
    # os.rmdir(tablepath)




if __name__ == '__main__':
    #231206 1100
    parser = argparse.ArgumentParser(description="Read text from an image and filter by specific conditions")

    parser.add_argument('--measure-csv', default='./mergecsv/27SYD-5YSR-231205.csv', help="검출 결과 measure 파일")
    # # parser.add_argument('--measure-csv', help="검출 결과 measure 파일")
    parser.add_argument('--output-path', default='231208_test2차', help="상태 등급 output 경로")
    #
    parser.add_argument('--config', default='J:/00.User/wlchoe/수자원연구원/StateEstimatorReport/27SYD-5YSR-231205.conf', help="시설물 Conf 파일")
    #
    parser.add_argument('--stateestimatorfiles', default='J:/00.User/wlchoe/수자원연구원/StateEstimator', help="상태평가보고서 작성용 파일들")

    # parser.add_argument('--config', help="시설물 Conf 파일")


    #
    # parser.add_argument('--measure-csv', help="검출 결과 measure 파일")
    # # parser.add_argument('--measure-csv', help="검출 결과 measure 파일")
    # parser.add_argument('--output-path', help="상태 등급 output 경로")
    #
    # parser.add_argument('--config', help="시설물 Conf 파일")
    # parser.add_argument('--stateestimatorfiles', help="상태평가보고서 작성용 파일들")


    args = parser.parse_args()

    file = args.measure_csv
    outputpath = args.output_path
    stateestimatorfiles = args.stateestimatorfiles
    config = args.config

    resultname = os.path.basename(config)[:-5]
    print('resultname', resultname)

    now = datetime.now()
    Today = str(now.year) + '-' + str(now.month) + '-' + str(now.day)

    estimatorroot = 'StateEstimator'
    createDirectory(outputpath)

    # file = 'CJD_merge_test2.csv'
    csvData = pd.read_csv(file)
    totalDataFrame = []
    print('--------------------------------------------')
    gradeScore = {5: 'a', 4: 'b', 3: 'c', 2: 'd', 1: 'e'}
    DamFloorData = csvData[csvData["structure"] == 'DamFloor']
    UpStreamData = csvData[csvData["structure"] == 'UpStream']
    DownStreamData = csvData[csvData["structure"] == 'DownStream']
    OverflowData = csvData[csvData["structure"] == 'overflow']


    unit1max = csvData["unit1"].max()
    print('unit1max', unit1max)

    # print(OverflowData)

    # spilwayData = csvData[csvData["structure"] == 'Spilway']
    # print(spilwayData)
    # step1Grade('Spilway', spilwayData, 6, totalDataFrame)

    DamFloorstructureLen = len(DamFloorData)
    UpStreamstructureLen = len(UpStreamData)
    DownStreamstructureLen = len(DownStreamData)
    OverflowstructureLen = len(OverflowData)


    # DamFloor_StateEstimation_231206.step1Grade('DamFloor', DamFloorData, DamFloorstructureLen, totalDataFrame)
    # DamFloor_StateEstimation_231206.step1Grade('UpStream', UpStreamData, UpStreamstructureLen, totalDataFrame)
    # DamFloor_StateEstimation_231206.step1Grade('overflow', OverflowData, OverflowstructureLen, totalDataFrame)
    DamFloor_StateEstimation_231206.step1Grade('overflow', unit1max, OverflowData, OverflowstructureLen, totalDataFrame)


    createDirectory('table')
    tablepath = os.path.join(outputpath, 'table')

    # DefectTable.defectTable('DamFloor', file, tablepath)
    # DefectTable.defectTable('UpStream', file, tablepath)
    DefectTable.defectTable('overflow', file, tablepath)

    mergetable(tablepath, outputpath)

    csvmerge(tablepath, outputpath)
#
#
#     # DamFloor_StateEstimation_231206.step1Grade('DamFloor', DamFloorData, 110, totalDataFrame)
#
#     # totalFrame = pd.DataFrame(totalDataFrame, columns=('StructureType', 'Unit', 'crackWidthGrade', 'crackInfulenceCoefficient', 'contractionWidthGrade', 'contractionInfluenceCoefficeient', 'desquamationGrade', 'desquamationInfulenceCoefficient'))
    totalFrame = pd.DataFrame(totalDataFrame, columns=(
    'StructureType', 'Unit',
    'crackWidthGrade', 'crackInfulenceCoefficient',
    'spallingGrade', 'spallingInfulenceCoefficient',
    'desquamationGrade', 'desquamationInfulenceCoefficient',
    'leakageGrade', 'leakageInfulenceCoefficient',
    'failGrade', 'failInfulenceCoefficient',
    'efflorescenceGrade', 'efflorescenceInfulenceCoefficient'))
#     # totalFrame.to_csv(str(outputpath)+'/1단계결과_231023.csv', index=False)
#
    totalFrame.to_csv(str(outputpath) + '/1단계결과_' + Today + '.csv', index=False)
    stage1_result = os.path.join(outputpath, '1단계결과_' + Today + '.csv')
#
#
#
    totalDataFrame2 = []
    file2 = stage1_result
    csvData2 = pd.read_csv(file2)
#
#
#
#     # spilwayData2 = csvData2[csvData2["StructureType"] == 'Spilway']
#     # print(spilwayData2)
#     # step2Grade('Spilway', spilwayData2, 6, totalDataFrame2)
#
#     # DamFloorData2 = csvData2[csvData2["StructureType"] == 'DamFloor']
#     # UpStreamData2 = csvData2[csvData2["StructureType"] == 'UpStream']
#     # DownStreamData2 = csvData2[csvData2["StructureType"] == 'DownStream']
    OverflowStreamData2 = csvData2[csvData2["StructureType"] == 'overflow']
#
#
#     # print(DamFloorData2)
#     # DamFloor_StateEstimation_231206.step2Grade('DamFloor', DamFloorData2, DamFloorstructureLen, totalDataFrame2)
#     # DamFloor_StateEstimation_231206.step2Grade('DownStream', UpStreamData2, UpStreamstructureLen, totalDataFrame2)
#     # DamFloor_StateEstimation_231206.step2Grade('DownStream', DownStreamData2, DownStreamstructureLen, totalDataFrame2)
    DamFloor_StateEstimation_231206.step2Grade('overflow', OverflowStreamData2, unit1max, totalDataFrame2)
#
#
    totalFrame2 = pd.DataFrame(totalDataFrame2, columns=('StructureType', 'Unit', 'Grade', 'Evalue'))
    stage2_result = os.path.join(outputpath, '2단계결과_' + Today + '.csv')
    totalFrame2.to_csv(stage2_result, index=False)
#
#
    totalDataFrame3 = []
    file3 = stage2_result
    stage3_result = os.path.join(outputpath, '3단계결과_' + Today + '.csv')
#     # saveFile = str(outputpath) + '/3단계결과_' + Today + '.csv'
#
#     # csvData3 = pd.read_csv(file3)
#     # spilwayData3 = csvData3[csvData3["StructureType"] == 'Spilway'].Evalue.tolist()
#     # step3Grade('Spilway', spilwayData3, 6, totalDataFrame3)
#
#
    csvData3 = pd.read_csv(file3)

    # DamFloorData3 = csvData3[csvData3["StructureType"] == 'DamFloor'].Evalue.tolist()
    # UpStreamData3 = csvData3[csvData3["StructureType"] == 'UpStream'].Evalue.tolist()
    # DownStreamData3 = csvData3[csvData3["StructureType"] == 'DownStream'].Evalue.tolist()
    OverflowData3 = csvData3[csvData3["StructureType"] == 'overflow'].Evalue.tolist()
    OverflowData3 = DamFloor_StateEstimation_231206.step2Grades(OverflowData3)

#     # DamFloor_StateEstimation_231206.step3Grade('DamFloor', DamFloorData3, DamFloorstructureLen, totalDataFrame3)
#     # DamFloor_StateEstimation_231206.step3Grade('UpStream', UpStreamData3, UpStreamstructureLen, totalDataFrame3)
#     # DamFloor_StateEstimation_231206.step3Grade('DownStream', UpStreamData3, UpStreamstructureLen, totalDataFrame3)
#     DamFloor_StateEstimation_231206.step3Grade('overflow', OverflowData3, OverflowstructureLen, totalDataFrame3)
    DamFloor_StateEstimation_231206.step3Grades('overflow', OverflowData3, totalDataFrame3)

#
#
    totalFrame3 = pd.DataFrame(totalDataFrame3, columns=('StructureType', 'Evalue', 'Grade'))
    totalFrame3.to_csv(stage3_result, index=False)
#
# 4단계
    totaltextList = []
    totalDataFrame4 = []

    # inputFile = str(outputpath) + '/3단계결과_' + Today + '.csv'
    # saveFile = str(outputpath) + '/4단계결과_' + Today + '.txt'
    # saveFile2 = str(outputpath) + '/4단계결과_' + Today + '.csv'

    inputFile = stage3_result
    saveFile = os.path.join(outputpath, '4단계결과_' + Today + '.txt')
    saveFile2 = os.path.join(outputpath, '4단계결과_' + Today + '.csv')

    csvData4 = pd.read_csv(inputFile)

    print(csvData4)


    # damFloorE3 = csvData4[csvData4["StructureType"] == 'DamFloor'].Evalue.tolist()[0]
    # upStreamE3 = csvData4[csvData4["StructureType"] == 'UpStream'].Evalue.tolist()[0]
    # downStreamE3 = csvData4[csvData4["StructureType"] == 'DownStream'].Evalue.tolist()[0]
    # OverflowE3 = csvData4[csvData4["StructureType"] == 'overflow'].Evalue.tolist()[0]

    # damFloorE3 = csvData4[csvData4["StructureType"] == 'DamFloor']['Evalue'].tolist()
    # if damFloorE3:
    #     damFloorE3 = damFloorE3[0]
    # else:
    #     damFloorE3 = 0
    #
    # upStreamE3 = csvData4[csvData4["StructureType"] == 'UpStream']['Evalue'].tolist()
    # if upStreamE3:
    #     upStreamE3 = upStreamE3[0]
    # else:
    #     upStreamE3 = 0
    #
    # downStreamE3 = csvData4[csvData4["StructureType"] == 'DownStream']['Evalue'].tolist()
    # if downStreamE3:
    #     downStreamE3 = downStreamE3[0]
    # else:
    #     downStreamE3 = 0

    OverflowE3 = csvData4[csvData4["StructureType"] == 'overflow']['Evalue'].tolist()
    if OverflowE3:
        OverflowE3 = OverflowE3[0]
    else:
        OverflowE3 = 0



    #region 규모

    jsonfile = 'scale.json'
    jsonfile = os.path.join(stateestimatorfiles, jsonfile)

    try:
        with open(jsonfile, "r") as jsonfile:
            Scale = json.load(jsonfile)
    except IOError as e:
        print('Error occured:', str(e))



    damFloorScale = float(Scale["damFloorScale"])
    upStreamScale = float(Scale["upStreamScale"])
    downStreamScale = float(Scale["downStreamScale"])
    overflowScale = float(Scale["overflowScale"])

    #region E3 * S 계산값
    # damFloorE3S = damFloorE3 * damFloorScale
    # upStreamE3S = upStreamE3 * upStreamScale
    # downStreamE3S = downStreamE3 * downStreamScale
    overflowE3S = OverflowE3 * overflowScale

    #region 상태평가 4단계 계산
    # step4Grade("본댐 제체",[damFloorE3,upStreamE3,downStreamE3],[damFloorScale,upStreamScale,downStreamScale],[damFloorE3S,upStreamE3S,downStreamE3S],totaltextList)
    # DamFloor_StateEstimation_231206.step4Grade("DamFloor", [damFloorE3], [damFloorScale], [damFloorE3S], totaltextList, totalDataFrame4)
    # DamFloor_StateEstimation_231206.step4Grade("UpStream", [upStreamE3], [upStreamScale], [upStreamE3S], totaltextList, totalDataFrame4)
    # DamFloor_StateEstimation_231206.step4Grade("DownStream", [downStreamE3], [downStreamScale], [downStreamE3S], totaltextList, totalDataFrame4)
    DamFloor_StateEstimation_231206.step4Grade("overflow", [OverflowE3], [overflowScale], [overflowE3S], totaltextList, totalDataFrame4)


    print("4단계" + str(totalDataFrame4))

    # endregion 상태평가 4단계 계산

    with open(saveFile, 'w') as f:
        for i in totaltextList:
            f.write(i + '\n')

    totalFrame4 = pd.DataFrame(totalDataFrame4, columns=('StructureType', 'Ec', 'Grade'))
    totalFrame4.to_csv(saveFile2, index=False)

    # Dam_StateEstimator_Report.Dam_StateEstimator_Report(config, file, outputpath, saveFile2 , '상태평가보고서_Dam_'+Today)
    Dam_StateEstimator_Report.Dam_StateEstimator_Report(config, file, outputpath,
                                                        stateestimatorfiles, saveFile2, resultname)


