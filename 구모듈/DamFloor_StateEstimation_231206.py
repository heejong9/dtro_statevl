import pandas as pd

def gradeScore(grade) :

    if grade is not None and grade < 1.5:
        return 'e'
    elif grade is not None and grade < 2.5 :
        return 'd'
    elif grade is not None and grade < 3.5:
        return 'c'
    elif grade is not None and grade < 4.5:
        return 'b'
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

def step1Grade(structureName, unit1max, data ,structureLen,totalDataFrame) :
    print('Stage 1')
    print(structureLen)

    for i in range(1, unit1max+1) :
        if structureName == 'overflow':
            print(structureName, i)
            unitNum = i
            # print(data[i])

            # print(data[i: i+1])
            # crackData = csvData[csvData['Type'] == "Crack"]
            print("unit1", unitNum)
            unitData = data[data["unit1"] == i]

            # unitData = data[i]

            print('--------------------------------------------')
            print(unitData)
            print('--------------------------------------------')

            if unitData.empty:
                crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient = 0,0,0,0,0,0,0,0,0,0,0,0
            else:
                crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient = OverflowOneStep(unitData)


        dataList = [structureName, unitNum, crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient]

        totalDataFrame.append(dataList)

def damFloodrOneStep(csvData)  -> object:

    # print('----DamFloor Start----')
    # print(csvData)
    #dataLen = len(csvData)
    #print('type')

    smdata = csvData['Type']


    #균열

    totalSpanArea = 692.30

    crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient = 0, 0, 0, 0, 0, 0

    crackData = csvData[csvData['Type'] == "Crack"]
    crackLength = crackData['Defect_L'].max() * 1000
    crackInfulenceCoefficient = 1.0
    crackArea = crackData['Defect_A']
    crackAreaSum = crackArea.sum()
    crackAreaRate = crackAreaSum / totalSpanArea

    if crackLength == 0:
        crackWidthGrade = 5.0
    elif 0.1 > crackLength and 0.1 > crackAreaRate:
        crackWidthGrade = 4.0
    elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
        crackWidthGrade = 3.0
    elif crackLength >= 0.5 and crackAreaRate >= 0.5:
        crackWidthGrade = 2.0
    else:
        crackWidthGrade = 1.0

    if (smdata.to_string().split(' ')[-1] == 'Crack'):
        crackWidthGrade = 5.0
        crackInfulenceCoefficient = 1.0
        contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
        desquamationGrade, desquamationInfulenceCoefficient = '', ''

    elif (smdata.to_string().split(' ')[-1] == 'contraction'):
        crackWidthGrade, crackInfulenceCoefficient = '', ''
        contractionWidthGrade, contractionInfluenceCoefficeient = 5, 1
        desquamationGrade, desquamationInfulenceCoefficient = '', ''

    elif (smdata.to_string().split(' ')[-1] == 'Desqu'):
        crackWidthGrade, crackInfulenceCoefficient = '', ''
        contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
        desquamationGrade, desquamationInfulenceCoefficient = 5, 1

    else:
        crackWidthGrade, crackInfulenceCoefficient = 1, 1
        contractionWidthGrade, contractionInfluenceCoefficeient = 1, 1
        desquamationGrade, desquamationInfulenceCoefficient = 1, 1



    # print(gradeScore(min(float(crackWidthGrade)*float(crackInfulenceCoefficient),
    #                      float(contractionWidthGrade)*float(contractionInfluenceCoefficeient),
    #                      float(desquamationGrade)*float(desquamationInfulenceCoefficient))))




    # contractionData = csvData[csvData['Type'] == "contraction"]
    # contractionLength = contractionData['Defect_L'].max() * 1000
    # contractionInfulenceCoefficient = 1.0
    # contractionArea = contractionData['Defect_A']
    # contractionAreaSum = contractionArea.sum()
    # contractionAreaRate = contractionAreaSum / totalSpanArea
    #
    # if contractionLength == 0:
    #     contractionInfulenceCoefficient = 5.0
    # elif 0.1 > contractionLength and 0.1 > contractionAreaRate:
    #     contractionInfulenceCoefficient = 4.0
    # elif 0.5 > contractionLength >= 0.1 and 0.5 > contractionAreaRate >= 0.1:
    #     contractionInfulenceCoefficient = 3.0
    # elif contractionLength >= 0.5 and contractionAreaRate >= 0.5:
    #     contractionInfulenceCoefficient = 2.0
    # else:
    #     contractionInfulenceCoefficient = 1.0


    # if [csvData['Type']] == "crack":
    #     print('crack')
    #     # crackData = csvData[csvData['Type'] == "Crack"]
    #     # crackDatalen = len(crackData)
    #     # print(crackDatalen)
    #     # # print(crackData)
    #     # # print(type(crackData))
    #     # # print(crackData)
    #     # crackLength = crackData['Defect_L'].max() * 1000
    #     # crackInfulenceCoefficient = 1.0
    #     # crackArea = crackData['Defect_A']
    #     # crackAreaSum = crackArea.sum()
    #     # crackAreaRate = crackAreaSum / totalSpanArea
    #     #
    #     # if crackLength == 0:
    #     #     crackWidthGrade = 5.0
    #     # elif 0.1 > crackLength and 0.1 > crackAreaRate:
    #     #     crackWidthGrade = 4.0
    #     #
    #     # elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
    #     #     crackWidthGrade = 3.0
    #     # elif crackLength >= 0.5 and crackAreaRate >= 0.5:
    #     #     crackWidthGrade = 2.0
    #     # else:
    #     #     crackWidthGrade = 1.0
    #     crackWidthGrade = 4.0
    #     crackInfulenceCoefficient = 1.0
    #     contractionWidthGrade = 0.0
    #     contractionInfluenceCoefficeient = 0.0
    #     desquamationGrade = 0.0
    #     desquamationInfulenceCoefficient = 0.0
    #
    # elif [csvData['Type']] == "contraction":
    #     print('contraction')
    #
    #     contractiondata = csvData[csvData['Type'] == "contraction"]
    #     crackWidthGrade = 0.0
    #     crackInfulenceCoefficient = 0.0
    #     contractionWidthGrade = 4.0
    #     contractionInfluenceCoefficeient = 4.0
    #     desquamationGrade = 0.0
    #     desquamationInfulenceCoefficient = 0.0
    #
    # if [csvData['Type']] == "Desqu":
    #     print('Desqu')
    #     DesqudataData = csvData[csvData['Type'] == "Desqu"]
    #
    #     crackWidthGrade = 0.0
    #     crackInfulenceCoefficient = 0.0
    #     contractionWidthGrade = 0.0
    #     contractionInfluenceCoefficeient = 0.0
    #     desquamationGrade = 4.0
    #     desquamationInfulenceCoefficient = 4.0


    #print(crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient)

    # 균열 , 균열 영향 계수
    return crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient

def UpStreamOneStep(csvData)  -> object:

    # print('----DamFloor Start----')
    # print(csvData)
    #dataLen = len(csvData)
    #print('type')

    ## 231024
    # smdata = csvData['Type']
    #
    #
    # #균열
    #
    # totalSpanArea = 692.30
    #
    # crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient = 0, 0, 0, 0, 0, 0
    #
    # crackData = csvData[csvData['Type'] == "Crack"]
    # crackLength = crackData['Defect_L'].max() * 1000
    # crackInfulenceCoefficient = 1.0
    # crackArea = crackData['Defect_A']
    # crackAreaSum = crackArea.sum()
    # crackAreaRate = crackAreaSum / totalSpanArea
    #
    # if crackLength == 0:
    #     crackWidthGrade = 5.0
    # elif 0.1 > crackLength and 0.1 > crackAreaRate:
    #     crackWidthGrade = 4.0
    # elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
    #     crackWidthGrade = 3.0
    # elif crackLength >= 0.5 and crackAreaRate >= 0.5:
    #     crackWidthGrade = 2.0
    # else:
    #     crackWidthGrade = 1.0
    #
    # if (smdata.to_string().split(' ')[-1] == 'Crack'):
    #     crackWidthGrade = 5.0
    #     crackInfulenceCoefficient = 1.0
    #     contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
    #     desquamationGrade, desquamationInfulenceCoefficient = '', ''
    #
    # elif (smdata.to_string().split(' ')[-1] == 'contraction'):
    #     crackWidthGrade, crackInfulenceCoefficient = '', ''
    #     contractionWidthGrade, contractionInfluenceCoefficeient = 4, 1
    #     desquamationGrade, desquamationInfulenceCoefficient = '', ''
    #
    # elif (smdata.to_string().split(' ')[-1] == 'Desqu'):
    #     crackWidthGrade, crackInfulenceCoefficient = '', ''
    #     contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
    #     desquamationGrade, desquamationInfulenceCoefficient = 3, 1
    #
    # else:
    #     crackWidthGrade, crackInfulenceCoefficient = 1, 1
    #     contractionWidthGrade, contractionInfluenceCoefficeient = 1, 1
    #     desquamationGrade, desquamationInfulenceCoefficient = 1, 1
    # #
    # # print(gradeScore(min(float(crackWidthGrade)*float(crackInfulenceCoefficient), float(contractionWidthGrade)*float(contractionInfluenceCoefficeient),
    # #                      float(desquamationGrade)*float(desquamationInfulenceCoefficient))))
    #
    #
    # # print(gradeScore(min((crackWidthGrade*crackInfulenceCoefficient), (contractionWidthGrade*contractionInfluenceCoefficeient), (desquamationGrade*desquamationInfulenceCoefficient))))
    #
    #
    # # contractionData = csvData[csvData['Type'] == "contraction"]
    # # contractionLength = contractionData['Defect_L'].max() * 1000
    # # contractionInfulenceCoefficient = 1.0
    # # contractionArea = contractionData['Defect_A']
    # # contractionAreaSum = contractionArea.sum()
    # # contractionAreaRate = contractionAreaSum / totalSpanArea
    # #
    # # if contractionLength == 0:
    # #     contractionInfulenceCoefficient = 5.0
    # # elif 0.1 > contractionLength and 0.1 > contractionAreaRate:
    # #     contractionInfulenceCoefficient = 4.0
    # # elif 0.5 > contractionLength >= 0.1 and 0.5 > contractionAreaRate >= 0.1:
    # #     contractionInfulenceCoefficient = 3.0
    # # elif contractionLength >= 0.5 and contractionAreaRate >= 0.5:
    # #     contractionInfulenceCoefficient = 2.0
    # # else:
    # #     contractionInfulenceCoefficient = 1.0
    #
    #
    # # if [csvData['Type']] == "crack":
    # #     print('crack')
    # #     # crackData = csvData[csvData['Type'] == "Crack"]
    # #     # crackDatalen = len(crackData)
    # #     # print(crackDatalen)
    # #     # # print(crackData)
    # #     # # print(type(crackData))
    # #     # # print(crackData)
    # #     # crackLength = crackData['Defect_L'].max() * 1000
    # #     # crackInfulenceCoefficient = 1.0
    # #     # crackArea = crackData['Defect_A']
    # #     # crackAreaSum = crackArea.sum()
    # #     # crackAreaRate = crackAreaSum / totalSpanArea
    # #     #
    # #     # if crackLength == 0:
    # #     #     crackWidthGrade = 5.0
    # #     # elif 0.1 > crackLength and 0.1 > crackAreaRate:
    # #     #     crackWidthGrade = 4.0
    # #     #
    # #     # elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
    # #     #     crackWidthGrade = 3.0
    # #     # elif crackLength >= 0.5 and crackAreaRate >= 0.5:
    # #     #     crackWidthGrade = 2.0
    # #     # else:
    # #     #     crackWidthGrade = 1.0
    # #     crackWidthGrade = 4.0
    # #     crackInfulenceCoefficient = 1.0
    # #     contractionWidthGrade = 0.0
    # #     contractionInfluenceCoefficeient = 0.0
    # #     desquamationGrade = 0.0
    # #     desquamationInfulenceCoefficient = 0.0
    # #
    # # elif [csvData['Type']] == "contraction":
    # #     print('contraction')
    # #
    # #     contractiondata = csvData[csvData['Type'] == "contraction"]
    # #     crackWidthGrade = 0.0
    # #     crackInfulenceCoefficient = 0.0
    # #     contractionWidthGrade = 4.0
    # #     contractionInfluenceCoefficeient = 4.0
    # #     desquamationGrade = 0.0
    # #     desquamationInfulenceCoefficient = 0.0
    # #
    # # if [csvData['Type']] == "Desqu":
    # #     print('Desqu')
    # #     DesqudataData = csvData[csvData['Type'] == "Desqu"]
    # #
    # #     crackWidthGrade = 0.0
    # #     crackInfulenceCoefficient = 0.0
    # #     contractionWidthGrade = 0.0
    # #     contractionInfluenceCoefficeient = 0.0
    # #     desquamationGrade = 4.0
    # #     desquamationInfulenceCoefficient = 4.0


    #print(crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient)

    dataLen = len(csvData)
    print('----231024 Test Start----')

    #균열
    crackWidthGrade = ''
    crackInfulenceCoefficient = ''

    # 박리
    spallingGrade = ''
    spallingInfulenceCoefficient = ''

    # 박락
    desquamationGrade = ''
    desquamationInfulenceCoefficient = ''

    # 누수
    leakageGrade = ''
    leakageInfulenceCoefficient = ''

    # 파손 및 손상
    failGrade = ''
    failInfulenceCoefficient = ''

    #백태
    efflorescenceGrade = ''
    efflorescenceInfulenceCoefficient = ''


    # #균열
    # crackWidthGrade = 5
    # crackInfulenceCoefficient = 1.0
    #
    # # 박리
    # spallingGrade = 5
    # spallingInfulenceCoefficient = 1.0
    #
    # # 박락
    # desquamationGrade = 5
    # desquamationInfulenceCoefficient = 1.0
    #
    # # 누수
    # leakageGrade = 5
    # leakageInfulenceCoefficient = 1.0
    #
    # # 파손 및 손상
    # failGrade = 5
    # failInfulenceCoefficient = 1.0
    #
    # #백태
    # efflorescenceGrade = 5
    # efflorescenceInfulenceCoefficient = 1.0



    if dataLen != 0 :
        totalSpanArea = 10000
        crackData = csvData[csvData['Type'] == "Crack"]

        crackWidth = crackData['Defect_W'].max()
        # crackInfulenceCoefficient = 1.0
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

        desquamationArea = csvData[csvData['Type'] == "Desqu"]
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

        contractionArea = csvData[csvData['Type'] == "contraction"]
        contractionArea = contractionArea['Defect_A']
        contractionAreaSum = contractionArea.sum()
        contractionAreaRate = contractionAreaSum / totalSpanArea
        contractionAreaRate = contractionAreaRate * 100000




        print('crackAreaRate = ' + str(crackAreaRate))
        print('crackWidth = '+str(crackWidth))
        print('failAreaRate = ' + str(failAreaRate))
        print('efflorescenceAreaRate = ' + str(efflorescenceAreaRate))
        print('spallingAreaRate = ' + str(spallingAreaRate))
        print('failAreaRate = ' + str(failAreaRate))
        print('desquamationAreaRate = ' + str(desquamationAreaRate))
        print('leakageAreaRate = ' + str(leakageAreaRate))
        print('contractionAreaRate = ' + str(contractionAreaRate))


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

        else:
            leakageGrade = 0


        ## 균열 파트
        if crackAreaRate <= 5 :
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
            else:
                crackWidthGrade = 0
                crackInfulenceCoefficient = 0

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
            else:
                crackWidthGrade = 0
                crackInfulenceCoefficient = 0
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
            else:
                crackWidthGrade = 0
                crackInfulenceCoefficient = 0

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

        elif 50<= failDepth < 80  and failAreaRate < 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth < 50  and failAreaRate > 10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth > 50  and failAreaRate > 10 :
            failGrade = 1
            failInfulenceCoefficient = 3.0

        else:
            failGrade = 0
            failInfulenceCoefficient = 0

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
        else:
            efflorescenceGrade = 0
            efflorescenceInfulenceCoefficient = 0

        ## 박리 파트

        if spallingAreaRate == 0:
            spallingGrade = 5
            spallingInfulenceCoefficient = 1.0

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
        else:
            spallingGrade = 0
            spallingInfulenceCoefficient = 0


        #박락
        if desquamationDepth == 0:
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
        else:
            desquamationGrade = 0
            desquamationInfulenceCoefficient = 0

        # if efflorescenceAreaRate == 0:
        #     efflorescenceGrade = 5
        #
        # elif efflorescenceAreaRate < 5 :
        #     efflorescenceGrade = 4
        #     failInfulenceCoefficient = 1.1
        #
        # elif efflorescenceAreaRate < 10:
        #     efflorescenceGrade = 3
        #     failInfulenceCoefficient = 1.3
        #
        # elif efflorescenceAreaRate < 20:
        #     efflorescenceGrade = 2
        #     efflorescenceInfulenceCoefficient = 1.7
        # elif efflorescenceAreaRate > 20 :
        #     efflorescenceGrade = 1
        #     efflorescenceInfulenceCoefficient = 3.0

    # 균열 , 균열 영향 계수
        return crackWidthGrade, crackInfulenceCoefficient, \
               contractionWidthGrade, contractionInfluenceCoefficeient,\
               spallingGrade, spallingInfulenceCoefficient, \
               desquamationGrade, desquamationInfulenceCoefficient, \
               leakageGrade, leakageInfulenceCoefficient, \
               failGrade, failInfulenceCoefficient, \
               efflorescenceGrade, efflorescenceInfulenceCoefficient,

        # my_list = [crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient,
        #            desquamationGrade, desquamationInfulenceCoefficient]
        #
        # # 빈 문자열을 제외하고 최솟값을 찾기
        # filtered_list = [x for x in my_list if x != '']
        # print('filtered_list')
        # print(filtered_list)
        # if filtered_list:
        #     min_value = min(filtered_list)
        # else:
        #     # 빈 문자열만 있는 경우 처리
        #     min_value = None
        #
        # print('min_value')
        # print(min_value)

    contractionWidthGrade, contractionInfluenceCoefficeient = 0, 0
    # 균열 , 균열 영향 계수
    return crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient

def DownStreamOneStep(csvData)  -> object:

    # print('----DamFloor Start----')
    # print(csvData)
    #dataLen = len(csvData)
    #print('type')

    ## 231024
    # smdata = csvData['Type']
    #
    #
    # #균열
    #
    # totalSpanArea = 692.30
    #
    # crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient = 0, 0, 0, 0, 0, 0
    #
    # crackData = csvData[csvData['Type'] == "Crack"]
    # crackLength = crackData['Defect_L'].max() * 1000
    # crackInfulenceCoefficient = 1.0
    # crackArea = crackData['Defect_A']
    # crackAreaSum = crackArea.sum()
    # crackAreaRate = crackAreaSum / totalSpanArea
    #
    # if crackLength == 0:
    #     crackWidthGrade = 5.0
    # elif 0.1 > crackLength and 0.1 > crackAreaRate:
    #     crackWidthGrade = 4.0
    # elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
    #     crackWidthGrade = 3.0
    # elif crackLength >= 0.5 and crackAreaRate >= 0.5:
    #     crackWidthGrade = 2.0
    # else:
    #     crackWidthGrade = 1.0
    #
    # if (smdata.to_string().split(' ')[-1] == 'Crack'):
    #     crackWidthGrade = 5.0
    #     crackInfulenceCoefficient = 1.0
    #     contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
    #     desquamationGrade, desquamationInfulenceCoefficient = '', ''
    #
    # elif (smdata.to_string().split(' ')[-1] == 'contraction'):
    #     crackWidthGrade, crackInfulenceCoefficient = '', ''
    #     contractionWidthGrade, contractionInfluenceCoefficeient = 4, 1
    #     desquamationGrade, desquamationInfulenceCoefficient = '', ''
    #
    # elif (smdata.to_string().split(' ')[-1] == 'Desqu'):
    #     crackWidthGrade, crackInfulenceCoefficient = '', ''
    #     contractionWidthGrade, contractionInfluenceCoefficeient = '', ''
    #     desquamationGrade, desquamationInfulenceCoefficient = 3, 1
    #
    # else:
    #     crackWidthGrade, crackInfulenceCoefficient = 1, 1
    #     contractionWidthGrade, contractionInfluenceCoefficeient = 1, 1
    #     desquamationGrade, desquamationInfulenceCoefficient = 1, 1
    # #
    # # print(gradeScore(min(float(crackWidthGrade)*float(crackInfulenceCoefficient), float(contractionWidthGrade)*float(contractionInfluenceCoefficeient),
    # #                      float(desquamationGrade)*float(desquamationInfulenceCoefficient))))
    #
    #
    # # print(gradeScore(min((crackWidthGrade*crackInfulenceCoefficient), (contractionWidthGrade*contractionInfluenceCoefficeient), (desquamationGrade*desquamationInfulenceCoefficient))))
    #
    #
    # # contractionData = csvData[csvData['Type'] == "contraction"]
    # # contractionLength = contractionData['Defect_L'].max() * 1000
    # # contractionInfulenceCoefficient = 1.0
    # # contractionArea = contractionData['Defect_A']
    # # contractionAreaSum = contractionArea.sum()
    # # contractionAreaRate = contractionAreaSum / totalSpanArea
    # #
    # # if contractionLength == 0:
    # #     contractionInfulenceCoefficient = 5.0
    # # elif 0.1 > contractionLength and 0.1 > contractionAreaRate:
    # #     contractionInfulenceCoefficient = 4.0
    # # elif 0.5 > contractionLength >= 0.1 and 0.5 > contractionAreaRate >= 0.1:
    # #     contractionInfulenceCoefficient = 3.0
    # # elif contractionLength >= 0.5 and contractionAreaRate >= 0.5:
    # #     contractionInfulenceCoefficient = 2.0
    # # else:
    # #     contractionInfulenceCoefficient = 1.0
    #
    #
    # # if [csvData['Type']] == "crack":
    # #     print('crack')
    # #     # crackData = csvData[csvData['Type'] == "Crack"]
    # #     # crackDatalen = len(crackData)
    # #     # print(crackDatalen)
    # #     # # print(crackData)
    # #     # # print(type(crackData))
    # #     # # print(crackData)
    # #     # crackLength = crackData['Defect_L'].max() * 1000
    # #     # crackInfulenceCoefficient = 1.0
    # #     # crackArea = crackData['Defect_A']
    # #     # crackAreaSum = crackArea.sum()
    # #     # crackAreaRate = crackAreaSum / totalSpanArea
    # #     #
    # #     # if crackLength == 0:
    # #     #     crackWidthGrade = 5.0
    # #     # elif 0.1 > crackLength and 0.1 > crackAreaRate:
    # #     #     crackWidthGrade = 4.0
    # #     #
    # #     # elif 0.5 > crackLength >= 0.1 and 0.5 > crackAreaRate >= 0.1:
    # #     #     crackWidthGrade = 3.0
    # #     # elif crackLength >= 0.5 and crackAreaRate >= 0.5:
    # #     #     crackWidthGrade = 2.0
    # #     # else:
    # #     #     crackWidthGrade = 1.0
    # #     crackWidthGrade = 4.0
    # #     crackInfulenceCoefficient = 1.0
    # #     contractionWidthGrade = 0.0
    # #     contractionInfluenceCoefficeient = 0.0
    # #     desquamationGrade = 0.0
    # #     desquamationInfulenceCoefficient = 0.0
    # #
    # # elif [csvData['Type']] == "contraction":
    # #     print('contraction')
    # #
    # #     contractiondata = csvData[csvData['Type'] == "contraction"]
    # #     crackWidthGrade = 0.0
    # #     crackInfulenceCoefficient = 0.0
    # #     contractionWidthGrade = 4.0
    # #     contractionInfluenceCoefficeient = 4.0
    # #     desquamationGrade = 0.0
    # #     desquamationInfulenceCoefficient = 0.0
    # #
    # # if [csvData['Type']] == "Desqu":
    # #     print('Desqu')
    # #     DesqudataData = csvData[csvData['Type'] == "Desqu"]
    # #
    # #     crackWidthGrade = 0.0
    # #     crackInfulenceCoefficient = 0.0
    # #     contractionWidthGrade = 0.0
    # #     contractionInfluenceCoefficeient = 0.0
    # #     desquamationGrade = 4.0
    # #     desquamationInfulenceCoefficient = 4.0


    #print(crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient)

    dataLen = len(csvData)
    print('----231024 Test Start----')


    #균열
    crackWidthGrade = 0
    crackInfulenceCoefficient = 0

    # 박리
    spallingGrade = 0
    spallingInfulenceCoefficient = 0

    # 박락
    desquamationGrade = 0
    desquamationInfulenceCoefficient = 0

    # 누수
    leakageGrade = 0
    leakageInfulenceCoefficient = 0

    # 파손 및 손상
    failGrade =0
    failInfulenceCoefficient =0

    #백태
    efflorescenceGrade = 0
    efflorescenceInfulenceCoefficient = 0



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

        desquamationArea = csvData[csvData['Type'] == "Desqu"]
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

        contractionArea = csvData[csvData['Type'] == "contraction"]
        contractionArea = contractionArea['Defect_A']
        contractionAreaSum = contractionArea.sum()
        contractionAreaRate = contractionAreaSum / totalSpanArea
        contractionAreaRate = contractionAreaRate * 100000




        print('crackAreaRate = ' + str(crackAreaRate))
        print('crackWidth = '+str(crackWidth))
        print('failAreaRate = ' + str(failAreaRate))
        print('efflorescenceAreaRate = ' + str(efflorescenceAreaRate))
        print('spallingAreaRate = ' + str(spallingAreaRate))
        print('failAreaRate = ' + str(failAreaRate))
        print('desquamationAreaRate = ' + str(desquamationAreaRate))
        print('leakageAreaRate = ' + str(leakageAreaRate))
        print('contractionAreaRate = ' + str(contractionAreaRate))


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
        if crackAreaRate <= 5:
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

        elif crackAreaRate <= 20:
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
            else:
                crackWidthGrade = 1
                crackInfulenceCoefficient = 1.0

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
        else:
            crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient = 5,1,5,1,5,1,5,1,5,1,5,1,5,1

            return crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient


    # 균열 , 균열 영향 계수
    return crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, desquamationGrade, desquamationInfulenceCoefficient

def OverflowOneStep(csvData)  -> object:
    dataLen = len(csvData)
    print('----231206 Test Start----')

    #균열
    crackWidthGrade = None
    crackInfulenceCoefficient = None

    # 박리
    spallingGrade = None
    spallingInfulenceCoefficient = None

    # 박락
    desquamationGrade = None
    desquamationInfulenceCoefficient = None
    # 누수
    leakageGrade = None
    leakageInfulenceCoefficient = None

    # 파손 및 손상
    failGrade = 0
    failInfulenceCoefficient = None

    #백태
    efflorescenceGrade = None
    efflorescenceInfulenceCoefficient = None


    if dataLen != 0 :
        totalSpanArea = 228.58

        crackData = csvData[csvData['Type'] == "Crack"]
        crackWidth = crackData['Defect_W'].max()
        crackInfulenceCoefficient = 1.0
        crackArea = crackData['Defect_A']
        crackAreaSum = crackArea.sum()
        crackAreaRate = crackAreaSum / totalSpanArea * 100000

        failArea = csvData[csvData['Type'] == "Fail"]
        failDepth = failArea['Defect_Depth'].max()
        failInfulenceCoefficient = 1.0
        failArea = failArea['Defect_A']
        failAreaSum = failArea.sum()
        failAreaRate = failAreaSum / totalSpanArea * 100000

        efflorescenceArea = csvData[csvData['Type'] == "Efflorescence"]
        efflorescenceArea = efflorescenceArea['Defect_A']
        efflorescenceAreaSum = efflorescenceArea.sum()
        efflorescenceAreaRate = efflorescenceAreaSum / totalSpanArea * 100000

        spallingArea = csvData[csvData['Type'] == "Spalling"]
        spallingArea = spallingArea['Defect_A']
        spallingAreaSum = spallingArea.sum()
        spallingAreaRate = spallingAreaSum / totalSpanArea * 100000

        desquamationArea = csvData[csvData['Type'] == "Desqu"]
        desquamationDepth = desquamationArea['Defect_Depth'].max()
        desquamationArea = desquamationArea['Defect_A']
        desquamationAreaSum = desquamationArea.sum()
        desquamationAreaRate = desquamationAreaSum / totalSpanArea * 100000

        leakageArea = csvData[csvData['Type'] == "Leakage"]
        leakageArea = leakageArea['Defect_A']
        leakageAreaSum = leakageArea.sum()
        leakageAreaRate = leakageAreaSum / totalSpanArea
        leakageAreaRate = leakageAreaRate * 100000

        REArea = csvData[csvData['Type'] == "RE"]
        REArea = REArea['Defect_A']
        REAreaSum = REArea.sum()
        REAreaRate = REAreaSum / totalSpanArea
        REAreaRate = REAreaRate * 100000

        PeelingArea = csvData[csvData['Type'] == "Peeling"]
        PeelingArea = PeelingArea['Defect_A']
        PeelingAreaSum = PeelingArea.sum()
        PeelingAreaSumRate = PeelingAreaSum / totalSpanArea
        PeelingAreaSumRate = PeelingAreaSumRate * 100000

        PeelingData = csvData[csvData['Type'] == "Peeling"]
        peelingArea = PeelingData['Defect_A']

        if REAreaRate == 0:
            crackWidthGrade = 1
            crackInfulenceCoefficient = 1
            spallingGrade = 1
            spallingInfulenceCoefficient = 1
            desquamationGrade = 1
            desquamationInfulenceCoefficient = 1
            leakageGrade = 1
            leakageInfulenceCoefficient = 1
            failGrade = 1
            failInfulenceCoefficient = 1
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 1

        if PeelingAreaSumRate == 0:
            crackWidthGrade = 1
            crackInfulenceCoefficient = 1
            spallingGrade = 1
            spallingInfulenceCoefficient = 1
            desquamationGrade = 1
            desquamationInfulenceCoefficient = 1
            leakageGrade = 1
            leakageInfulenceCoefficient = 1
            failGrade = 1
            failInfulenceCoefficient = 1
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 1

        print('crackAreaRate = ' + str(crackAreaRate))
        print('crackWidth = '+str(crackWidth))
        print('failAreaRate = ' + str(failAreaRate))
        print('efflorescenceAreaRate = ' + str(efflorescenceAreaRate))
        print('spallingAreaRate = ' + str(spallingAreaRate))
        print('failAreaRate = ' + str(failAreaRate))
        print('desquamationAreaRate = ' + str(desquamationAreaRate))
        print('leakageAreaRate = ' + str(leakageAreaRate))

        if leakageAreaRate == 0:
            leakageGrade = None
            leakageInfulenceCoefficient = None

        elif leakageAreaRate < 1:
            leakageGrade = 5
            leakageInfulenceCoefficient = 1.0

        elif leakageAreaRate < 0.1:
            leakageGrade = 4
            leakageInfulenceCoefficient = 1.0

        elif leakageAreaRate < 0.2:
            leakageGrade = 3
            leakageInfulenceCoefficient = 1.0

        elif leakageAreaRate < 0.3:
            leakageGrade = 2
            leakageInfulenceCoefficient = 1.0

        elif leakageAreaRate >= 0.4:
            leakageGrade = 1
            leakageInfulenceCoefficient = 1.0

        ## 균열 파트
        if crackAreaRate == 0:
            crackWidthGrade = None
            crackInfulenceCoefficient = None

        elif crackAreaRate == 0 and crackWidth != 0:
            crackWidthGrade = 5
            crackInfulenceCoefficient = 1

        elif crackAreaRate <= 0.5:
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

        elif crackAreaRate  <= 0.20 :
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
            else:
                crackWidthGrade = None
                crackInfulenceCoefficient = None

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
        if failAreaRate == 0:
            failGrade = None
            failInfulenceCoefficient = None

        elif failDepth <= 0.20 :
            failGrade = 5

        elif failDepth < 0.20 and failAreaRate < 0.10 :
            failGrade = 4
            failInfulenceCoefficient = 1.1

        elif 20<= failDepth < 0.50  and failAreaRate < 0.10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif failDepth <0.20  and failAreaRate > 0.10 :
            failGrade = 3
            failInfulenceCoefficient = 1.3

        elif failDepth <0.80  and failAreaRate < 0.10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth < 0.50  and failAreaRate > 0.10 :
            failGrade = 2
            failInfulenceCoefficient = 1.7

        elif failDepth > 0.50  and failAreaRate > 0.10 :
            failGrade = 1
            failInfulenceCoefficient = 3.0


        ## 백태 파트
        if efflorescenceAreaRate == 0:
            efflorescenceGrade = None
            failInfulenceCoefficient = None

        elif efflorescenceAreaRate != 0:
            efflorescenceGrade = 5
            failInfulenceCoefficient = 1.0

        elif efflorescenceAreaRate < 0.5 :
            efflorescenceGrade = 4
            failInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 0.10:
            efflorescenceGrade = 3
            failInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 0.20:
            efflorescenceGrade = 2
            efflorescenceInfulenceCoefficient = 1.7

        elif efflorescenceAreaRate > 0.02 :
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 3.0

        ## 박리 파트

        if spallingAreaRate == 0:
            spallingGrade = None
            spallingInfulenceCoefficient = None

        elif spallingAreaRate != 0:
            spallingGrade = 5
            spallingInfulenceCoefficient = 1.0

        elif spallingAreaRate < 0.5:
            spallingGrade = 4
            spallingInfulenceCoefficient = 1.1

        elif spallingAreaRate < 0.10:
            spallingGrade = 3
            spallingInfulenceCoefficient = 1.3

        elif spallingAreaRate < 0.20:
            spallingGrade = 2
            spallingInfulenceCoefficient = 1.7

        elif spallingAreaRate > 0.20 :
            spallingGrade = 1
            spallingInfulenceCoefficient = 3.0

        #박락
        if desquamationAreaRate == 0:
            desquamationGrade = None
            desquamationInfulenceCoefficient = None

        elif desquamationDepth != 0:
            desquamationGrade = 5
            desquamationInfulenceCoefficient = 1.0

        elif desquamationDepth < 0.15 and desquamationAreaRate < 0.10 :
            desquamationGrade = 4
            desquamationInfulenceCoefficient = 1.1

        elif 15<= desquamationDepth < 0.20  and desquamationAreaRate < 0.10 :
            desquamationGrade = 3
            desquamationInfulenceCoefficient = 1.2

        elif desquamationDepth <0.15  and desquamationAreaRate > 0.10 :
            desquamationGrade = 3
            desquamationInfulenceCoefficient = 1.2

        elif 20<= desquamationDepth <0.25  and desquamationAreaRate < 0.10 :
            desquamationGrade = 2
            desquamationInfulenceCoefficient = 1.4

        elif desquamationDepth < 0.20  and desquamationAreaRate > 0.10 :
            desquamationGrade = 2
            desquamationInfulenceCoefficient = 1.4
        elif desquamationDepth > 0.20  and desquamationAreaRate < 0.10  :
            desquamationGrade = 1
            desquamationInfulenceCoefficient = 2.0

        if efflorescenceAreaRate == 0:
            efflorescenceGrade = None
            efflorescenceInfulenceCoefficient = None

        elif efflorescenceAreaRate != 0:
            efflorescenceInfulenceCoefficient = 5

        elif efflorescenceAreaRate < 5 :
            efflorescenceGrade = 4
            efflorescenceInfulenceCoefficient = 1.1

        elif efflorescenceAreaRate < 10:
            efflorescenceGrade = 3
            efflorescenceInfulenceCoefficient = 1.3

        elif efflorescenceAreaRate < 20:
            efflorescenceGrade = 2
            efflorescenceInfulenceCoefficient = 1.7
        elif efflorescenceAreaRate > 20 :
            efflorescenceGrade = 1
            efflorescenceInfulenceCoefficient = 3.0

    # 균열 , 균열 영향 계수
            return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, \
                   desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, \
                   efflorescenceGrade, efflorescenceInfulenceCoefficient
        else:
            crackWidthGrade, crackInfulenceCoefficient, contractionWidthGrade, contractionInfluenceCoefficeient, spallingGrade, \
            spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, \
            failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient = None,None,None,None,None,None,None,None,None,None,None,None,None,None

            return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, \
                   desquamationGrade, desquamationInfulenceCoefficient, leakageGrade, leakageInfulenceCoefficient, \
                   failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient


    print('step1', crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, \
           leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient)

    # 균열 , 균열 영향 계수
    return crackWidthGrade, crackInfulenceCoefficient, spallingGrade, spallingInfulenceCoefficient, desquamationGrade, desquamationInfulenceCoefficient, \
           leakageGrade, leakageInfulenceCoefficient, failGrade, failInfulenceCoefficient, efflorescenceGrade, efflorescenceInfulenceCoefficient



def step2Grade(structureName, data ,structureLen,totalDataFrame) :
    import math
    print('Stage 2')

    # for i in range(1,structureLen+1) :
    for i in range(1, structureLen):
        unitNum = i
        unitData = data[data["Unit"] == unitNum]
        print(structureName, unitNum, unitData)



        crackGrade = float(unitData["crackWidthGrade"] * unitData["crackInfulenceCoefficient"])
        desquamationGrade = float(unitData["desquamationGrade"] * unitData["desquamationInfulenceCoefficient"])
        leakageGrade = float(unitData["leakageGrade"] * unitData["leakageInfulenceCoefficient"])
        failGrade = float(unitData["failGrade"] * unitData["failInfulenceCoefficient"])
        spallingGrade = float(unitData["desquamationGrade"] * unitData["desquamationInfulenceCoefficient"])
        efflorescenceGrade = float(unitData["efflorescenceGrade"] * unitData["efflorescenceInfulenceCoefficient"])


        print('stage2 grade')
        print(crackGrade, desquamationGrade, leakageGrade, failGrade, spallingGrade, efflorescenceGrade)

        my_list = [crackGrade, desquamationGrade, leakageGrade, failGrade, spallingGrade, efflorescenceGrade]
        # 빈 문자열을 제외하고 최솟값을 찾기
        filtered_list = [x for x in my_list if not math.isnan(x)]
        print('filtered_list', filtered_list)

        #
        # # if all(math.isnan(x) for x in filtered_list):
        # #     return 5
        #
        if filtered_list:
            min_value = min(filtered_list)
        else:
            # 빈 문자열만 있는 경우 처리
            min_value = None
        #
        print('min_value', min_value)

        # # grade = gradeScore[minValue]
        grade = gradeScore(min_value)
        # # print(gradeScore(minValue))
        print(grade)
        # # print(minValue)
        #
        if min_value is None:
            dataList = [structureName, unitNum, gradeScore(min_value), 5]
        else:
            dataList = [structureName, unitNum, gradeScore(min_value), min_value]
        totalDataFrame.append(dataList)

def step2Grades(e2List):
    print(e2List)
    # 0이 아닌 값들을 걸러내기
    non_zero_values = [x for x in e2List if x != 0]

    # 최솟값 구하기
    min_non_zero = min(non_zero_values)

    print('min_non_zero', min_non_zero)
    return min_non_zero

def step3Grades(structureName, e2List, totalDataFrame) :
    print('------------------------------')
    print('Stage 3 test')

    print(e2List)

    importanceWeight = 100
    print(structureName)
    print('가중치:' + str(importanceWeight))
    denominator =0 # 분모 계산식 합
    numerator = 0 # 분자 계산식 합
    count = 1

    adjustmentCoefficient = adjustmentCoefficientScore(e2List)
    awCalc = importanceWeight * adjustmentCoefficient
    eawCalc = e2List * awCalc
    print(structureName +'-'+str(count).zfill(3)+ ' E2 : '+ str(e2List) + ' , A : ' + str(adjustmentCoefficient)+
          ' , W : '+str(importanceWeight) +' , aw : '+str(awCalc) +' , eaw : '+str(eawCalc) )
    denominator += awCalc
    numerator += eawCalc
    count +=1
    print('numerator:' + str(numerator))

    if denominator != 0:
        e3 = round(numerator / denominator, 2)
        e3grade = gradeScore(e3)
    else:
        e3 = 1

    print('1.복합부재의 상태 평가 지수(E3) 값 = %s'%(e3))
    print('2.복합부재의 상태평가 등급 = %s등급'%(gradeScore(e3)))
    dataList = [structureName, e3, gradeScore(e3)]

    totalDataFrame.append(dataList)
    # print('3.분자합  = %s' % (numerator))
    # print('4.분모합 = %s' % (denominator))

def step3Grade(structureName, e2List, structureLen, totalDataFrame) :

    dataLen = len(e2List)

    if dataLen != 0 :

        print('------------------------------')
        print('Stage 3')

        print(e2List)

        importanceWeight = 100 / structureLen
        print(structureName)
        print('가중치:' + str(importanceWeight))
        denominator =0 # 분모 계산식 합
        numerator = 0 # 분자 계산식 합
        count = 1
        for e2 in e2List :
            adjustmentCoefficient = adjustmentCoefficientScore(e2)
            awCalc = importanceWeight * adjustmentCoefficient
            eawCalc = e2 * awCalc
            print(structureName +'-'+str(count).zfill(3)+ ' E2 : '+ str(e2) + ' , A : ' + str(adjustmentCoefficient)+
                  ' , W : '+str(importanceWeight) +' , aw : '+str(awCalc) +' , eaw : '+str(eawCalc) )
            denominator += awCalc
            numerator += eawCalc
            count +=1
            print('numerator:' + str(numerator))

        if denominator != 0:
            e3 = round(numerator / denominator, 2)
            e3grade = gradeScore(e3)
        else:
            e3 = 1

        print('1.복합부재의 상태 평가 지수(E3) 값 = %s'%(e3))
        print('2.복합부재의 상태평가 등급 = %s등급'%(gradeScore(e3)))
        dataList = [structureName, e3, gradeScore(e3)]

        totalDataFrame.append(dataList)
        # print('3.분자합  = %s' % (numerator))
        # print('4.분모합 = %s' % (denominator))

# 4단계(5-52p)
def step4Grade(structureName, e3List, scaleList, e3SList, txtList, totalDataFrame):
    print('Stage 4')

    e3Max = max(e3List)
    e3Min = min(e3List)
    # e3Max = e3List
    # e3Min = e3List
    scaleSum = sum(scaleList)
    e3ScaleSum = sum(e3SList)

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

