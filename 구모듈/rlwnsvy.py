# -*- coding: utf-8 -*-

import time
import DamFloor_StateEstimation
import DefectTableUseTotal


import pandas as pd
import os
import argparse


from datetime import datetime
from datetime import date

from collections import defaultdict
from io import BytesIO

import configparser


import win32com.client as win32
import win32com.client
from PIL import Image
import glob


import shutil

     
def 현황표():
      
        # 글자속성(17,1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(f"{config['Cover']['DamName']} 현황표")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("ParagraphShapeAlignLeft")
        crateTable(15,6)
       
        
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableCellBlockExtend")
        # 표간격(130)
        
        hwp.HAction.Run("TableResizeExLeft")
     
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
       # 첫 번째 셀로 이동
       
       # 첫 번째 열로 이동
        hwp.Run("TableColPageUp")
        hwp.Run("TableColBegin")  # 추가한 행의 첫 번째 셀로 이동
            
    
        # 첫 번째 행으로 이동

        hwp.Run("Cancel")


  
        
    
        cellMergeRange(2,1)
        cellMergeRange(3,2)
        
        hwp.HAction.Run("MoveDown")
        hwp.Run("TableColBegin")
        
        
        cellMergeRange(2,1)
        cellMergeRange(3,2)
        
        hwp.HAction.Run("MoveDown")
        hwp.Run("TableColBegin")
        
        cellMergeRange(5,1)
        
        hwp.HAction.Run("MoveDown")
        hwp.Run("TableColBegin")
        
        cellMergeRange(2,1)
        cellMergeRange(3,2)
        
        hwp.HAction.Run("MoveDown")
        hwp.Run("TableColBegin")
        
        cellLowerMerge(5)
        
        hwp.HAction.Run("MoveDown")
        
        cellLowerMerge(5)
        
        for _ in range(3):
            hwp.HAction.Run("MoveRight")
        
        for _ in range(6):
            hwp.HAction.Run("MoveUp")
        
        cellLowerMerge(5)
        
        hwp.HAction.Run("MoveDown")
        
        cellLowerMerge(5)
        
        hwp.Run("TableColBegin")
        for _ in range(5):
            hwp.HAction.Run("MoveUp")
            
        textTable()
   
      

def init_hwp(visible=True):
    """
    아래아한글 시작
    """
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = visible
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp

def insert_text(text):
        """
        문서에 텍스트 삽입
        """
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def cellRightText(listText):
        for i in listText:
            insert_text(i)
            hwp.HAction.Run("MoveRight")
def cellDownText(listText):
    for i in listText:
        insert_text(i)
        hwp.HAction.Run("MoveDown")
    
def cellLowerMerge(number):
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableCellBlockExtend")
    for _ in range(number):
        hwp.HAction.Run("TableLowerCell")
    hwp.HAction.Run("TableMergeCell")

def cellMergeRange(number,number2):
            
    
        for _ in range(number2):
            hwp.HAction.Run("MoveRight")
    
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(number-number2):
            hwp.HAction.Run("TableRightCell")
        
        hwp.HAction.Run("TableMergeCell")
        
def cellLowerMerge(number):
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(number):
            hwp.HAction.Run("TableLowerCell")
        hwp.HAction.Run("TableMergeCell")
        
def 기준표():
        cellMergeRange(8,5)
        hwp.HAction.Run("MoveRight")
        cellLowerMerge(5)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveRight")
        for _ in range(31):
            hwp.HAction.Run("MoveUp")
        
        cellLowerMerge(5)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        hwp.HAction.Run("MoveDown")
        cellLowerMerge(4)
        
    
        hwp.HAction.Run("MoveLeft")
        
        for _ in range(7):
            hwp.HAction.Run("MoveUp")
        cellDownText(['상태변화','균열','박리','박락','누수','백태','철근노출','침하'])
        
        for _ in range(8):
            hwp.HAction.Run("MoveUp")
          
        hwp.HAction.Run("TableRightCell")
        
        cellDownText(['평가유형','국부결함','국부결함','국부결함','국부결함','일반손상','국부결함','중요결함'])
   
        for _ in range(8):
            hwp.HAction.Run("MoveUp")
           
        hwp.HAction.Run("TableRightCell")
   
        cellDownText([
    '영향계수','영향계수',
    '1.0','1.1','1.2','1.4','2.0',
    '1.0','1.1','1.2','1.4','2.0',
    '1.0','1.1','1.2','1.4','2.0',
    '1.0','1.1','1.2','1.4','2.0',
    '1.0','1.1','1.3','1.7','3.0',
    '1.0','1.1','1.2','1.4','2.0',
   
    ])
        cellLowerMerge(4)
        insert_text("1.0")
     
        hwp.HAction.Run("MoveUp")
        hwp.HAction.Run("MoveRight")
        for _ in range(36):
            
            hwp.HAction.Run("MoveUp")
        cellDownText([
    '평가기준','평가기준',
    'A','B','C','D','E',
    'A','B','C','D','E',
    'A','B','C','D','E',
    'A','B','C','D','E',
    'A','B','C','D','E',
    'A','B','C','D','E',
    'A','B','C','D','E',
    ])
    
        hwp.HAction.Run("MoveUp")
        hwp.HAction.Run("MoveRight")
        for _ in range(36):
        
            hwp.HAction.Run("MoveUp")
        cellDownText([
    '평가점수','평가점수',
    '5','4','3','2','1',
    '5','4','3','2','1',
    '5','4','3','2','1',
    '5','4','3','2','1',
    '5','4','3','2','1',
    '5','4','3','2','1',
    '5','4','3','2','1',
    ])
        
        hwp.HAction.Run("MoveUp")
        hwp.HAction.Run("MoveRight")
        for _ in range(36):
            hwp.HAction.Run("MoveUp")
        insert_text("평가내용")
        hwp.HAction.Run("MoveDown")
        cellDownText([
    '최대균열폭',
    '균열폭 0.1mm 미만 ',
    '균열폭 0.1mm 이상 균열폭 0.3mm 미만 ',
    '균열폭 0.3mm 이상 균열폭 0.5mm 미만 ',
    '균열폭 0.5mm 이상 균열폭 1.0mm 미만',
    '균열폭 1.0mm 이상',
    ])
        hwp.HAction.Run("MoveRight")
        for _ in range(6):
            hwp.HAction.Run("MoveUp")
        
    
        cellDownText([
        '면적율 5%이하',
        'A','A','A','B','C',
        ])
        hwp.HAction.Run("MoveRight")
        for _ in range(6):
            hwp.HAction.Run("MoveUp")
        
    
        cellDownText([
        '면적율 20%이하',
        'A','A','B','C','D'
        ])
        hwp.HAction.Run("MoveRight")
        for _ in range(6):
            hwp.HAction.Run("MoveUp")
        
    
        cellDownText([
        '면적율 20%이상',
        'A','B','C','D','E',
        ])
        
        hwp.HAction.Run("MoveLeft")
        hwp.HAction.Run("MoveLeft")
        hwp.HAction.Run("MoveLeft")
        def listMergemake(text):
            cellMergeRange(3,0)
            insert_text(text)
            hwp.HAction.Run("MovePrevParaBegin")
            hwp.HAction.Run("TableLowerCell")
            
        listMergemake("박리 면적 25mm² 미만")
        listMergemake("박리 면적 25mm² 이상 박리 면적 75mm² 미만")
        listMergemake('박리 면적 75mm² 이상 박리 면적 150mm²미만')
        listMergemake('박리 면적 150mm² 이상 박리 면적 300mm² 미만')
        listMergemake('박리 면적 300mm² 이상')
    
        listMergemake('박락 면적 25mm² 미만')
        listMergemake('박락 면적 25mm² 이상 박락 면적 75mm² 미만')
        listMergemake('박락 면적 75mm² 이상 박락 면적 150mm²미만')
        listMergemake('박락 면적 150mm² 이상 박락 면적 300mm² 미만')
        listMergemake('박락 면적 300mm² 이상') 
            
        listMergemake('면적율 1%미만')
        listMergemake('면적율 5%미만')
        listMergemake('면적율 5~10%미만')
        listMergemake('면적율 10~20%미만')
        listMergemake('면적율 20%이상')

        listMergemake('면적율 1%미만')
        listMergemake('면적율 5%미만')
        listMergemake('면적율 5~10%미만')
        listMergemake('면적율 10~20%미만')
        listMergemake('면적율 20%이상')
        
        listMergemake('면적율 0.5%미만')
        listMergemake('면적율 1%미만')
        listMergemake('면적율 1~3%미만')
        listMergemake('면적율 3~5%미만')
        listMergemake('면적율 5%이상')
        listMergemake('침하깊이 0㎝～5㎝ 미만')
        listMergemake('침하깊이 5㎝ 이상～10㎝ 미만')
        listMergemake('침하깊이 10㎝ 이상～50㎝ 미만')
        listMergemake('침하깊이 50㎝ 이상 100㎝ 미만')
        listMergemake('침하깊이 100㎝ 이상')
        
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        
        hwp.HAction.Run("TableCellBlockExtend")
        # 표간격(130)
        # hwp.HAction.Run("TableDistributeCellHeight")
        # hwp.HAction.Run("TableDistributeCellWidth")
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("Close")

def crateTable(rows,cols):
        # 테이블 생성
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)


        # HTableCreation 파라미터 설정
        hwp.HParameterSet.HTableCreation.Rows = rows
        hwp.HParameterSet.HTableCreation.Cols = cols
        hwp.HParameterSet.HTableCreation.WidthType = 0
        hwp.HParameterSet.HTableCreation.HeightType = 0
        hwp.HParameterSet.HTableCreation.WidthValue = 0.0
        hwp.HParameterSet.HTableCreation.HeightValue = 0.0

        # 테이블 폭 설정
        hwp.HParameterSet.HTableCreation.TableProperties.Width = 41954
        # 테이블 생성 액션 실행
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)    

def insert_text_right(text):
    """
    문서에 텍스트 삽입
    """
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("MoveRight")       
def textTable():
    
    insert_text_right("시설물명")
    insert_text_right(config["Cover"]['facility_name'])
    insert_text_right("시설물번호")
    insert_text_right(config["Cover"]['facility_managenumber'])

    insert_text_right("준공년월일")
    insert_text_right(config["Cover"]['project_completiondate'])
    insert_text_right("관리번호")
    insert_text_right(config["Cover"]['facility_number'])
    
    insert_text_right('위치')
    insert_text_right(config["Cover"]['project_location'])
    
    insert_text_right('관리주체')
    insert_text_right(config["Cover"]['project_manager'])
    insert_text_right('TEL')
    insert_text_right(config["Cover"]['facility_tel'])
    
    cellDownText(["댐제원","여수로"])
    
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    
    cellDownText(['하천명','댐형식','댐정상 표고',
                '댐 높이','댐 길이','댐 체적',
                '여수로 계획홍수량','여수로 계획방류량','여수로 최대방류량','여수로 게이트',
                '여수로 형식'])
    
    for _ in range(14):
        hwp.HAction.Run("MoveUp")
   
    hwp.HAction.Run("MoveRight")
    
    cellDownText([config["Cover"]['river_name'],config["Cover"]['damspecifictype'],config["Cover"]['dampeak'],
                config["Cover"]['damheight'],config["Cover"]['damlength'],config["Cover"]['damvolume'],
                config["Cover"]['여수로 계획홍수량'],config["Cover"]['여수로 계획방류량'],config["Cover"]['여수로 최대방류량'],config["Cover"]['여수로 게이트'],config["Cover"]['여수로 형식'],
                ])
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    hwp.HAction.Run("MoveUp")
    
    cellDownText(["저수지제원","보조여수로"])
    
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    
    cellDownText(['총저수량','유효저수량','유역면적',
                '계획 홍수위','상시 홍수위','저수위',
                '보조여수로 계획홍수량', '보조여수로 계획방류량','보조여수로 최대방류량','보조여수로 게이트',
                '보조여수로 형식'])

   
    for _ in range(16):
        hwp.HAction.Run("MoveUp")
   
    hwp.HAction.Run("TableRightCell")
    
    cellDownText([config["Cover"]['총저수량'],config["Cover"]['유효저수량'],config["Cover"]['유역면적'],
                config["Cover"]['계획홍수위'],config["Cover"]['상시만수위'],config["Cover"]['저수위'],
                config["Cover"]['보조여수로 계획홍수량'],config["Cover"]['보조여수로 계획방류량'],config["Cover"]['보조여수로 최대방류량'],config["Cover"]['보조여수로 게이트'],config["Cover"]['보조여수로 형식'],
                ])
               
if __name__ == '__main__':  
    config = configparser.ConfigParser()    

    
    with open(r"C:\Users\user\Desktop\statevlK2\25DGD_HRM.conf", 'r', encoding='utf-8') as configfile:
        config.read_file(configfile)
    # with open(r"C:\Users\user\Desktop\statevlK\27SYD_YSR.conf", 'r', encoding='utf-8') as configfile:
    #     config.read_file(configfile)
    hwp = init_hwp()
    print(type(hwp))
    현황표()
    # crateTable(37,9)
    # 기준표()
    