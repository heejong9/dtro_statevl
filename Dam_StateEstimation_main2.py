# -*- coding: utf-8 -*-

import time
import DamFloor_StateEstimation
import DefectTableUseTotal
import win32clipboard

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


class Win32COMCacheManager:
    def __init__(self):
        """Temp 폴더 내 gen_py 경로 설정"""
        self.temp_gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")

    def ensure_gen_py_exists(self):
        """gen_py 폴더가 없으면 생성"""
        if not os.path.exists(self.temp_gen_py_path):
            os.makedirs(self.temp_gen_py_path)
            print(f" Created {self.temp_gen_py_path} folder.")

    def clear_cache(self):
        """Temp 경로의 gen_py 폴더 삭제"""
        if os.path.exists(self.temp_gen_py_path):
            try:
                shutil.rmtree(self.temp_gen_py_path)
                print(f" Deleted {self.temp_gen_py_path} cache.")
            except PermissionError:
                print(f" Warning: Could not delete {self.temp_gen_py_path} due to permission error.")

    def regenerate_all_modules(self):
        """EnsureDispatch를 사용해 COM 모듈을 다시 생성"""
        self.ensure_gen_py_exists()

        try:
            # 기본적인 COM 모듈 호출하여 gen_py 재생성
            win32com.client.gencache.EnsureDispatch("Excel.Application")
            win32com.client.gencache.EnsureDispatch("Word.Application")
            win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
            print(" 기본 COM 객체 재생성 완료.")
        except Exception as e:
            print(f" 기본 COM 객체 생성 중 오류 발생: {e}")


def checkBujea(bujea):
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

def get_unique_filenames_without_extension(directory_path):
    # 파일 이름 (확장자 제거) 저장을 위한 set 생성
    
    unique_filenames = set()
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
   # 지정된 폴더 내 파일만 확인
    for file in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file)
        
        # 파일인지 확인 (폴더 제외)
        if os.path.isfile(file_path):
            filename_without_extension, extension = os.path.splitext(file)
            
            if extension.lower() in image_extensions:
                allPath[filename_without_extension].add(file_path)
                unique_filenames.add(filename_without_extension)


            
            
    return unique_filenames
     


bujeaDictionary = {
    "YSR": "여수로",
    "BYI": "보조여수로",
    "BYO": "보조여수로",
    "BYR": "보조여수로",
    "DMR": "댐마루",
    "HRM": "하류사면",
    "SRM": "상류사면",
    "CSA": "취수탑",
    "CSB": "취수탑",
    "CSC": "취수탑",
    "CSD": "취수탑",
}


def init_hwp(visible=True):
    """
    아래아한글 시작
    """
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = visible
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp


# if __name__ == '__main__':
#     config = configparser.ConfigParser()
#     with open('./statevl/SYD/27SYD_YSR.conf', 'r', encoding='utf-8') as configfile:
#         config.read_file(configfile)


#     # main_page()
#     # hwp.Run("BreakPage")  # 페이지 나누기 삽입
#     # 목차()
#     # hwp.Run("BreakPage")
#     # 안전진단표()
#     # 다음페이지로2()
#     # 결과요약()
#     # 다음페이지로()
   
 
#     # hwp.Run("BreakPage")  # 페이지 나누기 삽입
#     # 시설물현황()

#     # hwp.Run("BreakPage")  # 페이지 나누기 삽입
#     # 상태평가기준()
#     # hwp.Run("BreakPage")
#     report("YSR")
config = configparser.ConfigParser()    

def makeHwp(damName,bujeaName,dirPath,imgPath):
    
    with open(config1, 'r', encoding='utf-8') as configfile:
        config.read_file(configfile)
    def 들여쓰기(number):
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

        # 들여쓰기 설정 (첫 줄 들여쓰기만 적용)
        hwp.HParameterSet.HParaShape.Indentation = number

        # 문단 모양 적용
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    def 여백생성(number):
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

        # 단위 변환: pt → HWP 내부 단위 (1pt = 100 HWPUnit)
        hwp.HParameterSet.HParaShape.LeftMargin = number*100*2
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    def insert_text(text, level=0):
        """
        문서에 제목 스타일을 적용한 후 텍스트 삽입
        """
      

        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    def insert_text_right(text):
        """
        문서에 텍스트 삽입
        """
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HAction.Run("TableRightCell")


   


    def append_hwp(filename):
   
        """
        문서 끼워넣기
        """
        hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
        hwp.HParameterSet.HInsertFile.KeepSection = 0
        hwp.HParameterSet.HInsertFile.KeepCharshape = 0
        hwp.HParameterSet.HInsertFile.KeepParashape = 0
        hwp.HParameterSet.HInsertFile.KeepStyle = 0
        hwp.HParameterSet.HInsertFile.filename = filename
        hwp.HAction.Execute("InsertFile", hwp.HParameterSet.HInsertFile.HSet)


    def 클립보드로_이미지_삽입_로고(filepath):
        """
        한/글 API의 InsertPicture 메서드는
        셀의 크기를 변경하지 않는 반면(이미지가 찌그러짐)
        클립보드를 통해 이미지를 삽입하면
        이미지의 종횡비에 맞춰
        셀의 높이가 자동으로 조절됨.
        """
        #이미지 = Image.open(filepath)
        

        # 아웃풋 = BytesIO()
        # 이미지.convert('RGB').save(아웃풋, 'BMP')
        # 최종데이터 = 아웃풋.getvalue()[14:]
        # 아웃풋.close()

        # win32clipboard.OpenClipboard()
        # win32clipboard.EmptyClipboard()
        # win32clipboard.SetClipboardData(win32clipboard.CF_DIB, 최종데이터)
        # win32clipboard.CloseClipboard()
        # HWP 실행
        # hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # hwp.XHwpWindows.Item(0).Visible = True
    
        # hwp.Run("Paste")
        
         # 이미지 삽입
     
        
         
        hwp.InsertPicture(filepath, True, 1, False, False, 0, 100, 20)

        # 이미지 크기 조절
  
        
        
        # HParameterSet 설정
        hwp.Run("SelectCtrlReverse")
    
        # 개체에 대한 설정 적용
        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
        

        
        # 개체를 글자처럼 취급하도록 설정
        hwp.HParameterSet.HShapeObject.HSet.SetItem("TreatAsChar", 1)
        
        # 개체 유형을 설정 (1은 일반 이미지)
        hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 1)
        
        # 설정된 매개변수로 실행
        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

       
        hwp.Run('Cancel')
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableCellBlockExtend")
        
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
        hwp.Run('Cancel')
     
   
    def 클립보드로_이미지_삽입(filepath):
        """
        한/글 API의 InsertPicture 메서드는
        셀의 크기를 변경하지 않는 반면(이미지가 찌그러짐)
        클립보드를 통해 이미지를 삽입하면
        이미지의 종횡비에 맞춰
        셀의 높이가 자동으로 조절됨.
        """
        #이미지 = Image.open(filepath)
        

        # 아웃풋 = BytesIO()
        # 이미지.convert('RGB').save(아웃풋, 'BMP')
        # 최종데이터 = 아웃풋.getvalue()[14:]
        # 아웃풋.close()

        # win32clipboard.OpenClipboard()
        # win32clipboard.EmptyClipboard()
        # win32clipboard.SetClipboardData(win32clipboard.CF_DIB, 최종데이터)
        # win32clipboard.CloseClipboard()
        # HWP 실행
        # hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # hwp.XHwpWindows.Item(0).Visible = True
    
        # hwp.Run("Paste")
        
         # 이미지 삽입
     
        
         
        hwp.InsertPicture(filepath, True, 1, False, False, 0, 130, 130)

        # 이미지 크기 조절
    
        
        
        # HParameterSet 설정
        hwp.Run("SelectCtrlReverse")
    
        # 개체에 대한 설정 적용
        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
        

        
        # 개체를 글자처럼 취급하도록 설정
        hwp.HParameterSet.HShapeObject.HSet.SetItem("TreatAsChar", 1)
        
        # 개체 유형을 설정 (1은 일반 이미지)
        hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 1)
        
        # 설정된 매개변수로 실행
        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

       
        hwp.Run('Cancel')
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableCellBlockExtend")
        
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
        hwp.Run('Cancel')
     

   
    def 클립보드로_이미지_삽입2(filepath):
        """
        한/글 API의 InsertPicture 메서드는
        셀의 크기를 변경하지 않는 반면(이미지가 찌그러짐)
        클립보드를 통해 이미지를 삽입하면
        이미지의 종횡비에 맞춰
        셀의 높이가 자동으로 조절됨.
        """
        # 이미지 = Image.open(filepath)
        

        # 아웃풋 = BytesIO()
        # 이미지.convert('RGB').save(아웃풋, 'BMP')
        # 최종데이터 = 아웃풋.getvalue()[14:]
        # 아웃풋.close()

        # win32clipboard.OpenClipboard()
        # win32clipboard.EmptyClipboard()
        # win32clipboard.SetClipboardData(win32clipboard.CF_DIB, 최종데이터)
        # win32clipboard.CloseClipboard()
        # # HWP 실행
        # time.sleep(0.3)
        # hwp.Run("Paste")
        
        
        
        
        def insert_resized_image(hwp, filepath, max_width_mm=150, max_height_mm=100):
            img = Image.open(filepath)
            width_px, height_px = img.size  # 픽셀 단위 크기

            # mm 단위 기준 비율 계산 (원본 비율 유지)
            aspect_ratio = width_px / height_px

            # 기준에 따라 크기 조정
            if aspect_ratio >= max_width_mm / max_height_mm:
                # 너비 기준으로 조정 (이미지가 더 가로로 큼)
                target_width_mm = max_width_mm
                target_height_mm = target_width_mm / aspect_ratio
            else:
                # 높이 기준으로 조정 (이미지가 더 세로로 큼)
                target_height_mm = max_height_mm
                target_width_mm = target_height_mm * aspect_ratio

            # 이미지 삽입
            hwp.InsertPicture(filepath, True, 1, False, False, 0, target_width_mm, target_height_mm)
        
        
        insert_resized_image(hwp, filepath, max_width_mm=150, max_height_mm=200)
         # 이미지 삽입
     
        # img = Image.open(filepath)
        # width, height = img.size  # 픽셀 단위

        # # 고정할 높이(mm)
        # target_height_mm = 100

        # # 비율 계산: 원본 비율을 유지하며 너비(mm) 계산
        # aspect_ratio = width / height
        # target_width_mm = target_height_mm * aspect_ratio
         
        # hwp.InsertPicture(filepath, True, 1, False, False, 0,target_width_mm, target_height_mm)

        # 이미지 크기 조절
 
        
        
        # HParameterSet 설정
        hwp.Run("SelectCtrlReverse")
    
        # 개체에 대한 설정 적용
        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)
        

        
        # 개체를 글자처럼 취급하도록 설정
        hwp.HParameterSet.HShapeObject.HSet.SetItem("TreatAsChar", 1)
        
        # 개체 유형을 설정 (1은 일반 이미지)
        hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 1)
        
        # 설정된 매개변수로 실행
        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

       
        
        # hwp.HAction.Run("TableCellBlock")
        # hwp.HAction.Run("TableCellBlockExtend")
        # hwp.HAction.Run("TableCellBlockExtend")
        
        # hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
        hwp.Run('Cancel')
        hwp.Run("MoveDown")


    


    

    
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
  
        
    def cellMerge(text1,text2,text3,text4,number):
            insert_text(text1)   
            hwp.HAction.Run("MoveRight")
            insert_text(config["Cover"][text3])
            hwp.HAction.Run("MoveRight")
            insert_text(text2)
            hwp.HAction.Run("MoveRight")
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            for _ in range(number):
                hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableMergeCell")
            insert_text(config["Cover"][text4])
            hwp.HAction.Run("TableColBegin")  # 첫 번째 열로 이동
            
            hwp.HAction.Run("TableLowerCell")  # 아래 행으로 이동
            

    def cellMergeRange(number,number2):
            
        
            for _ in range(number2):
                hwp.HAction.Run("MoveRight")
        
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            for _ in range(number-number2):
                hwp.HAction.Run("TableRightCell")
            
            hwp.HAction.Run("TableMergeCell")
        
            
        
    def cellNoMerge(text1,text2,text3,text4,text5,text6):
        insert_text(text1)   
        hwp.HAction.Run("MoveRight")
        insert_text(config["Cover"][text4])
        hwp.HAction.Run("MoveRight")
        insert_text(text2)
        hwp.HAction.Run("MoveRight")
        insert_text(config["Cover"][text5])
        hwp.HAction.Run("MoveRight")
        insert_text(text3)
        hwp.HAction.Run("MoveRight")
        insert_text(config["Cover"][text6])
        hwp.HAction.Run("TableColBegin")  # 첫 번째 열로 이동
        hwp.HAction.Run("TableLowerCell")  # 아래 행으로 이동
    def cellAllMerge(text1):
        insert_text(text1)   
        hwp.HAction.Run("MoveRight")
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(4):
            hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("TableMergeCell")
        hwp.HAction.Run("TableColBegin")  # 첫 번째 열로 이동
        hwp.HAction.Run("TableLowerCell")  # 아래 행으로 이동
        
    def cellAllMerge2(text1,n):
        insert_text(text1)   
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(n):
            hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("TableMergeCell")
        hwp.HAction.Run("TableColBegin")  # 첫 번째 열로 이동
        hwp.HAction.Run("TableLowerCell")  # 아래 행으로 이동

    def 안전진단표():
        # 텍스트 삽입
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(17,1)
        insert_text(f"{config['Cover']['damname']} 정밀안전진단 결과표")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("ParagraphShapeAlignJustify")

        # 텍스트 삽입
        글자속성(15,1)
        insert_text('1. 기본현황')
        
        # 테이블 생성
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

        # 열 너비와 행 높이를 위한 배열을 수동으로 생성
        col_widths = [21.1] * 6  # 6개의 열 너비 값을 설정
        row_heights = [0.0] * 11  # 11개의 행 높이 값을 설정

        # HTableCreation 파라미터 설정
        hwp.HParameterSet.HTableCreation.Rows = 11
        hwp.HParameterSet.HTableCreation.Cols = 6
        hwp.HParameterSet.HTableCreation.WidthType = 0
        hwp.HParameterSet.HTableCreation.HeightType = 0
        hwp.HParameterSet.HTableCreation.WidthValue = 0.0
        hwp.HParameterSet.HTableCreation.HeightValue = 0.0

        # ColWidth와 RowHeight 설정
        for i in range(6):
            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_widths[i])

        for i in range(11):
            hwp.HParameterSet.HTableCreation.RowHeight.SetItem(i, row_heights[i])

        # 테이블 폭 설정
        hwp.HParameterSet.HTableCreation.TableProperties.Width = 41954

        # 테이블 생성 액션 실행
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

        # 테이블 셀 작업
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(5):
            hwp.HAction.Run("TableRightCell")
    
        
        hwp.HAction.Run("TableMergeCell")
        insert_text("가. 일반현황")
        hwp.HAction.Run("MoveDown")
        
        
       
        cellMerge("용역명","진단기간",'project_name','project_period',2)
        cellMerge("관리주체명","대표자",'project_manager','project_head',2)
        cellMerge("공동수급","계약방법",'coapply','project_contract',2)
        cellNoMerge("시설물","시설물구분","종별",'project_facility','project_facilitytype','project_facilityclass')
        cellNoMerge("준공일","진단금액","안전등급",'project_completiondate','inspectionprice','safetygrade')
        cellMerge("시설물위치","시설물규모",'project_location','project_scale',2)
        
        
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(5):
            hwp.HAction.Run("TableRightCell")
    
        
        hwp.HAction.Run("TableMergeCell")
        insert_text("나. 진단 실시결과 현황")
        
        hwp.HAction.Run("TableColBegin")
        hwp.HAction.Run("MoveDown")
        
        cellAllMerge("중대결함")
        cellAllMerge("진단 주요결과")

        cellAllMerge("주요 보수보강")



    def get_xlsx_files_in_current_directory(dirPath):
        # 현재 디렉토리에서 .xlsx 파일 찾기
        table_path = os.path.join(dirPath, "table")
        xlsx_files = glob.glob(os.path.join(table_path, '*.csv'))
    
    
    
        return xlsx_files

    def main_page():
        ## 한글 API 문제로 인한 절대경로 수정
        # rel_path = r".\hwp\test1.hwp"
        # abs_path = os.path.abspath(rel_path)
        # hwp.Open(abs_path)
        # print(os.getcwd())
        # os.chdir(hwp.Path.rsplit("\\", maxsplit=1)[0])
        # print(os.getcwd())
        # 한글 문서에서 쪽 번호 위치 설정
        hwp.HAction.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)

        # DrawPos 설정: 하단 중앙(3은 하단 중앙을 나타냄)
        hwp.HParameterSet.HPageNumPos.DrawPos = 5  # 3 = 하단 중앙, 1 = 상단 좌측, 2 = 상단 중앙, 4 = 하단 좌측, 등

        # 설정 적용
        hwp.HAction.Execute("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        
        글자속성(28,True)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = f"{config['Cover']['damname']} 상태평가 보고서"
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    
        hwp.HAction.Run("MoveDown")
    

        # 글자 크기를 10으로 설정
        글자속성(15,True)
        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
        for i in range(0,10):
            hwp.HAction.Run("BreakPara")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        today_date = date.today().strftime("%Y년 %m월 %d일")
        hwp.HParameterSet.HInsertText.Text = f"\n\n작성일: {today_date}"
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        for i in range(0,8):
            hwp.HAction.Run("BreakPara")
        글자속성(14,True)
        클립보드로_이미지_삽입_로고(mainPath+'/logo.png')
        hwp.HAction.Run("MoveDocEnd")
        hwp.HAction.Run("BreakPara")
        
      

 
    def report(bujea,dirPath,imgpath):  
        
    
        xlsx_files_list = get_xlsx_files_in_current_directory(dirPath)
      


        for idx, data in enumerate(xlsx_files_list):
            file_name = os.path.splitext(os.path.basename(data))[0]
        
            df = pd.read_csv(os.path.join(os.getcwd(), data),encoding='cp949')
            
            append_hwp(os.path.join(os.getcwd(), "템플릿3.template"))
            time.sleep(0.1)
            #get_nst_table(idx)  
            hwp.FindCtrl()      
            print("Selected Control ID:", hwp.HeadCtrl.CtrlID)
            print("Current Control Position:", hwp.HeadCtrl.GetAnchorPos(0))

            hwp.Run("ShapeObjTableSelCell")

            hwp.Run("TableLowerCell")
            #여기서 부재 이름이랑 부재 코드 쓸것
            insert_text(file_name)
            
            hwp.Run("TableRightCell")
            insert_text(bujeaDictionary.get(bujea))
            hwp.Run("TableColBegin")
            hwp.Run("TableLowerCell")
            #이미지 들어가야함
            data_name = file_name
            img_name = data_name+'.jpg'
            imgPahtR=os.path.dirname(imgPath)
          
         
            
           
            action = hwp.CreateAction("TableSplitCell")
            hwp.Run('TableCellBlockExtend')
            if action is None:
                raise ValueError("액션 객체 생성 실패")

            # 파라미터셋 생성 및 기본값 설정
            param_set = action.CreateSet()
            action.GetDefault(param_set)

            # 필요한 파라미터 설정 (예: 분할할 셀의 범위 지정)
            # param_set.SetItem("ItemName", "Value") 형식으로 설정
            # 예시: param_set.SetItem("Row", 1)
            #       param_set.SetItem("Column", 1)

            param_set.SetItem("Rows", len(allPath[data_name]))  # 분할할 행 수
            param_set.SetItem("Cols", 0)
            # 액션 실행
            action.Execute(param_set)
      
            for i in allPath[data_name]:
                i = i.replace("stage05", "stage05_thumb")
                print(i)
                클립보드로_이미지_삽입2(i)
              
           
            hwp.Run("TableColBegin")
            hwp.Run("TableLowerCell")   
            hwp.Run("TableLowerCell")
            hwp.Run("TableLowerCell")

            for row in range(len(df)):  # 모든 행을 순회하면서
                for col in range(6):  # 한 행씩 입력
                    if not pd.isnull(df.iloc[row, col]):
                    
                        insert_text(df.iloc[row,col])  # 첫 번째 셀부터 차례대로 입력
                
                    if col != 5:  # 마지막 열애 도착하기 전까지는
                        hwp.Run("TableRightCell")  # 입력후 우측셀로
                if len(df) - row != 1:  # 마지막 열에 도착하면
                    hwp.Run("TableAppendRow")  # 우측셀로 가지 말고 행 추가
                    hwp.Run("TableColBegin")  # 추가한 행의 첫 번째 셀로 이동
            
            
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableCellBlockExtend")
            
            hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
            hwp.Run('Cancel')
            hwp.Run("MoveDown")
            # hwp.Run("MoveDown")
       
            hwp.HAction.Run("MoveDocEnd") 
            hwp.Run("BreakPage")
            time.sleep(0.5)



    def 글자속성(font_size=11,bold=False):
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        hwp.HParameterSet.HCharShape.HSet.SetItem("Bold", 1 if bold else 0)  # 진하게 설정 (1: 진하게, 0: 일반)
        hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(font_size)
    
        hwp.HParameterSet.HCharShape.FaceNameUser = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameSymbol = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameOther = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameJapanese = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameHanja = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameLatin = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameHangul = "한컴바탕"  # 글자모양 - 글꼴종류

        hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType("TTF")  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType("TTF")  # 글자모양 - 폰트타입

        hwp.HParameterSet.HCharShape.SizeUser = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeSymbol = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeOther = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeJapanese = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeHanja = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeLatin = 100  # 글자모양 - 상대크기%
        hwp.HParameterSet.HCharShape.SizeHangul = 100  # 글자모양 - 상대크기%

        hwp.HParameterSet.HCharShape.RatioUser = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioSymbol = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioOther = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioJapanese = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioHanja = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioLatin = 100  # 글자모양 - 장평%
        hwp.HParameterSet.HCharShape.RatioHangul = 100  # 글자모양 - 장평%

        hwp.HParameterSet.HCharShape.SpacingUser = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingSymbol = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingOther = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingJapanese = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingHanja = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingLatin = 0  # 글자모양 - 자간%
        hwp.HParameterSet.HCharShape.SpacingHangul = 0  # 글자모양 - 자간%

        hwp.HParameterSet.HCharShape.OffsetUser = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetSymbol = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetOther = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetJapanese = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetHanja = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetLatin = 0  # 글자모양 - 글자위치%
        hwp.HParameterSet.HCharShape.OffsetHangul = 0  # 글자모양 - 글자위치%




        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
        
    def 표간격(value):
        
    #     # PageSetup 액션 실행
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
        hparashape = hwp.HParameterSet.HParaShape
        hparashape.LineSpacingType = 0
        hparashape.LineSpacing = value  # 줄 간격 설정
        
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    def 줄간격(line_spacing):
        # paragraph_shape = hwp.XHwpDocuments.Item(0).XHwpParagraphShape
        # paragraph_shape.LineSpacing = line_spacing
        act = hwp.CreateAction("ParagraphShape")  # 액션 생성
        pset = act.CreateSet()  # 파라미터셋 생성
        act.GetDefault(pset)  # 파라미터셋에 현재 상태값 채워넣기

      

        pset.SetItem("LineSpacing", line_spacing)  # 줄간격을 300%로 설정
        act.Execute(pset)  # 설정한 파라미터셋으로 액션 실행
        


      

        pset.SetItem("LineSpacing", line_spacing)  # 줄간격을 300%로 설정
        act.Execute(pset)  # 설정한 파라미터셋으로 액션 실행
    
    def 목차():
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        
        글자속성(23,1)
        insert_text("목       차")
        
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")

        Set = hwp.HParameterSet.HParaShape
        hwp.HAction.GetDefault("ParagraphShape", Set.HSet)
        tab_def = Set.TabDef
        tab_def.CreateItemArray("TabItem", 12)
        tab_def.TabItem.SetItem(0, 92000)
        tab_def.TabItem.SetItem(1, 3)
        tab_def.TabItem.SetItem(2, 0)
        hwp.HAction.Execute("ParagraphShape", Set.HSet)

        글자속성(15,1)

        # 목차를 수동으로 삽입 (예시)
        hwp.HAction.Run("ParagraphShapeAlignJustify") 
        insert_text("1. 시설물 개요\t3")
        hwp.HAction.Run("BreakPara")
        글자속성(15,0)
        insert_text("  1.1 시설물 현황\t3")
        hwp.HAction.Run("BreakPara")
        insert_text("  1.2 관련 사진\t4")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        글자속성(15,1)
        insert_text("2. 상태평가 개요\t5")
        글자속성(15,0)
        hwp.HAction.Run("BreakPara")
        insert_text("  2.1 상태평가 항목 및 기준\t5")
        hwp.HAction.Run("BreakPara")
        insert_text("    2.1.1 평가유형 영향계수 및 기준산정 방법\t5")
        hwp.HAction.Run("BreakPara")
        insert_text("    2.1.2 상태평가 항목 및 기준\t6")
        hwp.HAction.Run("BreakPara")
        insert_text("  2.2 상태평가 결과 산정 방법\t8")
        hwp.HAction.Run("BreakPara")
        insert_text("    2.2.1 댐 시설물 평가 단계별 절차\t8")
        hwp.HAction.Run("BreakPara")
        insert_text("    2.2.2 상태평가 단계별 구분\t8")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        글자속성(15,1)
        insert_text("3. 상태평가 결과\t9")
        글자속성(15,0)
        hwp.HAction.Run("BreakPara")
        insert_text("   3.1 1단계 상태평가 : 부재(部材)별 손상상태 평가표 작성 \t9")
        hwp.HAction.Run("BreakPage")
                
    def 결과요약():
        글자속성(15,1)
        insert_text("2. 결과요약")
        hwp.HAction.Run("BreakPara")
        crateTable(2,1)
        hwp.HAction.Run("MoveUp")
        hwp.HAction.Run("MoveLineEnd")
        
        select_nearest_table()
        
        hwp.HAction.Run("MoveDown")
    
        insert_text("책임 기술자 종합 의견")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        글자속성(13,1)
        insert_text("가. 정밀안전진단 외관조사 결과 기본사항")
        crateTable(8,5)
        cellAllMerge2("종합평가 결과",4)
        cellMergeRange(1,0)
        listText = ['결함발생 부재',	'상태평가 결과',	'결함 종류',	'보수 보강']
        cellRightText(listText)
        listText = ['필댐체',	'우안 비여수로',	'좌안 비여수로',	'여수로']
        cellDownText(listText)
        hwp.HAction.Run("MoveUp")
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(2):
            hwp.HAction.Run("TableLowerCell")
        hwp.HAction.Run("TableMergeCell")
        hwp.HAction.Run("MoveDown")
        
    def cellRightText(listText):
        for i in listText:
            insert_text(i)
            hwp.HAction.Run("TableRightCell")
    def cellDownText(listText):
        for i in listText:
            insert_text(i)
            
            hwp.HAction.Run("TableLowerCell")
        
    def cellLowerMerge(number):
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        for _ in range(number):
            hwp.HAction.Run("TableLowerCell")
        hwp.HAction.Run("TableMergeCell")
        
    def select_nearest_table(): 
        # 현재 커서 위치 가져오기
    
        try:
            hwp.HAction.Run("Table")  # 커서를 다음 표로 이동
        except Exception as e:
            print("표로 이동할 수 없습니다.", e)
    
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
        
        hwp.HAction.Run("TableUpperCell")
      
        
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        cellDownText(['하천명','댐형식','댐정상 표고',
                    '댐 높이','댐 길이','댐 체적',
                    '여수로 계획홍수량','여수로 계획방류량','여수로 최대방류량','여수로 게이트',
                    '여수로 형식'])
        
        for _ in range(13):
            hwp.HAction.Run("MoveUp")
       
        hwp.HAction.Run("MoveRight")
      
        cellDownText([config["Cover"]['river_name'],config["Cover"]['damspecifictype'],config["Cover"]['dampeak'],
                    config["Cover"]['damheight'],config["Cover"]['damlength'],config["Cover"]['damvolume'],
                    config["Cover"]['여수로 계획홍수량'],config["Cover"]['여수로 계획방류량'],config["Cover"]['여수로 최대방류량'],config["Cover"]['여수로 게이트'],config["Cover"]['여수로 형식'],
                    ])
        
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableRightCell")
       
        hwp.HAction.Run("TableUpperCell")
       
        cellDownText(["저수지제원","보조여수로"])
       
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableUpperCell")
       
        cellDownText(['총저수량','유효저수량','유역면적',
                    '계획 홍수위','상시 홍수위','저수위',
                    '보조여수로 계획홍수량', '보조여수로 계획방류량','보조여수로 최대방류량','보조여수로 게이트',
                    '보조여수로 형식'])
        
        

    
        for _ in range(10):
            hwp.HAction.Run("TableUpperCell")
   
        hwp.HAction.Run("TableRightCell")
      
        cellDownText([config["Cover"]['총저수량'],config["Cover"]['유효저수량'],config["Cover"]['유역면적'],
                    config["Cover"]['계획홍수위'],config["Cover"]['상시만수위'],config["Cover"]['저수위'],
                    config["Cover"]['보조여수로 계획홍수량'],config["Cover"]['보조여수로 계획방류량'],config["Cover"]['보조여수로 최대방류량'],config["Cover"]['보조여수로 게이트'],config["Cover"]['보조여수로 형식'],
                    ])
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableCellBlockExtend")
            # 표간격(130)
  
            
        hwp.HAction.Run("TableResizeExLeft")   
        hwp.HAction.Run("TableResizeExLeft") 
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExDown")
        hwp.HAction.Run("TableResizeExDown")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        hwp.HAction.Run("Close")
        hwp.HAction.Run("MoveDocEnd") 
    
        
    def 현황표():
        
            # 글자속성(17,1)
           
            hwp.HAction.Run("ParagraphShapeAlignJustify")
            crateTable(15,6)
        
            
           
           
            
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
            
         

    def 시설물현황():
        글자속성(13,1)
       
        insert_text("1. 시설물 개요",1)

        글자속성(11,0)
        hwp.HAction.Run("BreakPara")
        여백생성(13.3)
        들여쓰기(1000)
        insert_text(f"{config['Cover']['damname']}은 {config['Cover']['project_location']}에 위치하고, 높이 {config['Cover']['damheight']}, 길이 {config['Cover']['damlength']}의 {config['Cover']['damspecifictype']}으로 {config['Cover']['Project_completiondate']}에 준공되었다. 댐의 제원 및 시설물의 현황은 다음과 같다.")
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        여백생성(8.3)
        글자속성(13,1)
        insert_text("1.1 시설물 현황",2)

        hwp.HAction.Run("BreakPara")
        현황표()
        줄간격(180)
        hwp.HAction.Run("BreakPara")
        
        hwp.Run("BreakPage")
        글자속성(13,1)
        hwp.HAction.Run("ParagraphShapeAlignJustify")
        insert_text(" 1.2 관련 사진",2)

        hwp.HAction.Run("BreakPara")
        ##사진 표
        crateTable(1,1)
        
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), os.path.join(estimatorPath, "그림2_2.png")))
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("ParagraphShapeAlignJustify")
        hwp.Run("BreakPage")
        ##사진 표
        # crateTable(1,1)   
        # 클립보드로_이미지_삽입(os.path.join(os.getcwd(), os.path.join(".", "StateEstimator", "그림", "그림2_3.png")))
        # hwp.HAction.Run("MoveDown")
        # hwp.HAction.Run("ParagraphShapeAlignCenter")
        # insert_text(f"[그림2.3]{config['Cover']['facility_name']} 표준 단면도")
        # hwp.HAction.Run("BreakPara")
        # hwp.HAction.Run("ParagraphShapeAlignJustify")
        # hwp.Run("BreakPage")
        글자속성(13,1)
        insert_text("2. 상태평가 개요 ",1)
     
        글자속성()
        hwp.HAction.Run("BreakPara")
        여백생성(13.3)
        들여쓰기(1000)
        insert_text(f"시설물의 상태평가는 「시설물의 안전 및 유지관리 실시 지침 국토교통부, 국토안전관리원)」에 따라 실시하며, 상태평가에 대한 세부적인 사항은 「시설물의 안전 및 유지관리 실시 세부지침(안전점검·진단편의 댐편, 2021.12)」을 준용하였다. 따라서, 「시설물의 안전 및 유지관리 실시 세부지침(안전점검·진단편의 댐편)」에 의거하여 외관조사 및 내구성 조사의 항목 및 수량에 따라 과업을 실시한 후, 상태평가를 위하여 중요 손상 및 결함을 세부 기준에 의해 분류하고 평가기법 및 절차에 따라 각 개별시설물에 대한 결함의 등급과 점수 및 지수를 산정하여 상태등급을 최종적으로 결정하였다.")
        
        
        # hwp.Run("BreakPa")  # 페이지 나누기 삽입
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
       
        글자속성(13,1)
        여백생성(8.3)
        insert_text("2.1 상태평가 항목 및 기준 ",2)
    
        hwp.HAction.Run("BreakPara")
        글자속성(13,1)
        여백생성(18.3)
        insert_text("2.1.1 평가유형·영향계수 및 기준산정 방법",3)
        
        글자속성()
        hwp.HAction.Run("BreakPara")
       
        들여쓰기(1000)
        text="시설물의 상태평가는 결함 및 손상에 따른 각각의 상태평가 기준을 적용하며, 상태변화가 전체 구조물에 미치는 안전성의 영향정도, 구조적인 중요도가 적절히 고려되어 평가될 수 있도록 결함 및 손상을 평가유형(評價類型)별로 구분하여 영향계수를 적용한다. \n\n\
1) 평가유형의 구분 \n\
결함 및 손상에 대한 평가유형은 다음과 같이 구분한다.\n\n\
① 중요결함\n\
침하, 경사\전도 및 활동 등과 같이 전체 구조물의 구조적인 안전에 직접 영향을 미치는 결함\n\n\
② 국부결함\n\
수평이음부 불량 등과 같이 구조물의 안전성에 직접적인 영향을 미치지는 않지만 손상이 진전될 경우 전체 구조물의 안전에 상당한 영향을 끼칠 수 있는 결함\n\n\
③ 일반손상\n\
파손, 마모, 콘크리트 재료분리 등과 같이 구조물의 안전에 크게 영향을 주지 않는 일반적인 손상\n\n\
2) 영향계수의 적용\n\
각 부재에서 발생하는 각종 손상 및 결함에 대한 상태평가 시 손상이 전체 구조물에 미치는 안전성의 영향정도, 구조적인 중요도가 적절히 고려되어 평가될 수 있도록 영향계수를 적용한다.영향계수는 안전성에 직접적인 영향을 미치는 중요 결함의 상태등급을 기준으로 하여 국부적인 결함의 등급을 상향 조정함으로써 이들이 전체 구조물에 미치는 영향을 평가 절하하는 계수이며, 영향계수는 상태평가를 위한 표준기준이며, 책임기술자의 판단으로 다소 조정할 수 있다.\n"
        # 텍스트를 줄바꿈 문자('\n')로 분리
        lines = text.split('\n')
        margins = [33.3,0,43.3,63.3,0,56.3,78.3,0,56.3,78.3,0,56.3,78.3,0,43.3,63.3]
        indents = [800,0,800,0,0,1400,2000,0,1400,2000,0,1400,2000,0,800,2000]
        # 한글 문서에 텍스트 삽입 및 줄바꿈 처리
        for line,margin,indent in zip(lines,margins,indents):
            여백생성(margin)
            들여쓰기(indent)
            insert_text(line)
            hwp.HAction.Run("BreakPara")  # 줄바꿈 삽입
        hwp.HAction.Run("BreakPara")
        여백생성(0)
        들여쓰기(0)

    def 기준표():
        cellMergeRange(8,5)
        hwp.HAction.Run("TableRightCell")
        cellLowerMerge(5)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableRightCell")
        for _ in range(31):
            hwp.HAction.Run("MoveUp")
        
        cellLowerMerge(5)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        hwp.HAction.Run("TableLowerCell")
        cellLowerMerge(4)
        
    
        hwp.HAction.Run("TableLeftCell")

        for _ in range(7):
            hwp.HAction.Run("TableUpperCell")
        cellDownText(['상태변화','균열','박리','박락','누수','백태','철근노출','침하'])
      
        for _ in range(8):
            hwp.HAction.Run("TableUpperCell")
          
        hwp.HAction.Run("TableRightCell")
        
        cellDownText(['평가유형','국부결함','국부결함','국부결함','국부결함','일반손상','국부결함','중요결함'])
        
        for _ in range(8):
            hwp.HAction.Run("TableUpperCell")
           
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
     
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableRightCell")
           
        for _ in range(31):
            hwp.HAction.Run("TableUpperCell")
        
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
    
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableRightCell")
       
        for _ in range(36):
           
            hwp.HAction.Run("TableUpperCell")
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
        
        hwp.HAction.Run("TableUpperCell")
        hwp.HAction.Run("TableRightCell")
        for _ in range(36):
            hwp.HAction.Run("TableUpperCell")
       
        insert_text("평가내용")
        hwp.HAction.Run("TableLowerCell")
        cellDownText([
    '최대균열폭',
    '균열폭 0.1mm 미만 ',
    '균열폭 0.1mm 이상 균열폭 0.3mm 미만 ',
    '균열폭 0.3mm 이상 균열폭 0.5mm 미만 ',
    '균열폭 0.5mm 이상 균열폭 1.0mm 미만',
    '균열폭 1.0mm 이상',
    ])
        hwp.HAction.Run("TableRightCell")
        for _ in range(6):
            hwp.HAction.Run("TableUpperCell")
        
    
        cellDownText([
        '면적율 5%이하',
        'A','A','A','B','C',
        ])
        hwp.HAction.Run("TableRightCell")
        for _ in range(6):
            hwp.HAction.Run("TableUpperCell")
        
    
        cellDownText([
        '면적율 20%이하',
        'A','A','B','C','D'
        ])
        hwp.HAction.Run("TableRightCell")
        for _ in range(6):
            hwp.HAction.Run("TableUpperCell")
        
    
        cellDownText([
        '면적율 20%이상',
        'A','B','C','D','E',
        ])
       
        hwp.HAction.Run("TableLeftCell")
        hwp.HAction.Run("TableLeftCell")
        hwp.HAction.Run("TableLeftCell")
     
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
        표간격(130)
        ctrl = hwp.CurSelectedCtrl or hwp.ParentCtrl
        pset = hwp.CreateSet("Table")
        pset.SetItem("TreatAsChar", False)  # 글자처럼 취급
        ctrl.Properties = pset
      
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("TableResizeExLeft")
        hwp.HAction.Run("Close")
  

    def 상태평가기준():
        글자속성(13,1)
        여백생성(18.3)
        insert_text("2.1.2 상태평가 항목 및 기준 ")
        글자속성()
        hwp.HAction.Run("BreakPara")
        여백생성(33.3)
        들여쓰기(1000)
        insert_text(f"본 과업대상 시설물인 {config['Cover']['facility_name']}은(는) {config['Cover']['DamSpecificType']} 형식으로 {config['Cover']['Project_FacilityType']}에 준하는 기준을 적용하였다.")
        hwp.HAction.Run("BreakLine")
        insert_text("「세부지침」에 준하여 정량적이고 객관적인 상태평가를 위하여 부재별, 개별부재별, 복합부재별, 개별시설별 각 부재별 상태평가 항목은 다음과 같다.")
        
        hwp.HAction.Run("BreakPara")
        ## 표 작성
      
        crateTable(37,9)
       
    
        기준표()
       
        hwp.Run("BreakPara")
        hwp.Run("BreakPage")  # 페이지 나누기 삽입
       
        글자속성(13,1)
        여백생성(8.3)
        insert_text("2.2 상태평가 결과 산정 방법 ",2)
      
        글자속성()
        hwp.HAction.Run("BreakPara")
        여백생성(18.3)
        글자속성(13,1)
        insert_text("2.2.1 댐 시설물 평가 단계별 절차 ",3)
        hwp.HAction.Run("BreakPara")
        글자속성()
        여백생성(33.3)
        들여쓰기(1000)
        insert_text("댐 시설물에 대한 상태평가는 댐 시설물의 상태평가 단계별 구분표와 같이 단계별로 구분할 때 댐 시설물은 통합시설물 (6단계) 에 해당하는 시설물로서 간주하고, 하위단계인 복합시설, 개별시설, 복합부재, 개별부재로 구분한다.외관조사망도는 개별부재에 대하여 작성하는 것을 원칙으로 하고 필요시 개별부재의 크기, 면적에 따라 부위별로 분할하여 작성한다.")
        hwp.HAction.Run("BreakPara")
        insert_text("AI기반 결함검출 결과 정보를 기반으로 자동생성하는 상태평가보고서는 1단계 부재별 결함정보을 표기하는 내용으로 한정한다.")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        여백생성(18.3)
        글자속성(13,1)
        insert_text("2.2.2 상태평가 단계별 구분",3)
        글자속성()
        hwp.HAction.Run("BreakPara")
        여백생성(33.3)
        들여쓰기(1000)
        insert_text("시설물의 상태를 평가하기 위하여 시설물을 단계별로 구분하여 다음 표와 같이 평가단계별 구분표를 작성하고 본 보고서에 수록한다.")
        들여쓰기(0)
        여백생성(0)
        ##사진필요
        crateTable(1,1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), os.path.join(estimatorPath, "표3_11.png")))
        hwp.HAction.Run("BreakPara")
       
      
    
        hwp.HAction.Run("MoveDown")
     
    
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("댐 시설물의 상태평가 단계별 구분표")
        hwp.Run("BreakPage")
        글자속성(13,1)
        hwp.HAction.Run("ParagraphShapeAlignJustify")
        여백생성(0)
        들여쓰기(0)
        insert_text("3. 상태평가 결과 ",1)
       
        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        insert_text("3.1 1단계 상태평가 : 부재(部材)별 손상상태 평가표 작성 ")
        hwp.HAction.Run("BreakPara")
        글자속성()
        여백생성(18.3)
        들여쓰기(1000)
        insert_text("시설물의 상태평가 단계별 구분표에 따라 개별부재를 1개 외관조사망도 또는 필요에 따라 부위별로 다수의 외관조사망도로 구분하여 개략도에 손상 및 결함상태를 도시하고, 조사결과표에 개별부재에 대한 손상내용을 상세히 기록한 후, 그 손상 정도에 대하여 5단계(a～e) 상태평가 결과 및 평가점수를 부여한다.")
        hwp.HAction.Run("BreakPara")
        여백생성(28.3)
        insert_text("○  손상상태 평가표에는 평가항목에 없는 상태변화라 할지라도 모두 기록하는 것을 원칙으로 한다.")
        hwp.HAction.Run("BreakPara")
        insert_text("○ 각 상태변화에 대한 상태평가 결과가 c, d, e 등급일 경우 보수ㆍ보강 우선순위에 따라 보수ㆍ보강을 한다.")
        hwp.HAction.Run("BreakPara")
        insert_text("○ 해당 상태평가 보고서는 AI 분석을 통한 1단계 상태평가를 기준으로 작성되었다.")
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        
    

    try:
        hwp = init_hwp()
        hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)

        # HSecDef 설정 변경
        hsecdef = hwp.HParameterSet.HSecDef
        hsecdef.PageDef.LeftMargin = 20.0*283.465  # 왼쪽 여백
        hsecdef.PageDef.RightMargin =20.0*283.465  # 오른쪽 여백
        hsecdef.PageDef.TopMargin = 10*283.465
        hsecdef.HSet.SetItem("ApplyClass", 24)  # 적용 클래스 설정 (24: 현재 섹션)
        hsecdef.HSet.SetItem("ApplyTo", 3)     # 적용 대상 설정 (3: 전체 문서)

        # PageSetup 액션 실행
        hwp.HAction.Execute("PageSetup", hsecdef.HSet)
        
        줄간격(180)
        main_page()
        hwp.Run("BreakPage")  # 페이지 나누기 삽입
        목차()
        # 안전진단표() 삭제
        # 다음페이지로2()
        # 결과요약() 삭제
        # 다음페이지로()
    
    
        
        시설물현황()
      
        # hwp.Run("BreakPage")  # 페이지 나누기 삽입
        상태평가기준()
        
        hwp.Run("BreakPage")
        report(bujeaName,dirPath,imgPath)
       
      

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        # set save filename
        hwp.HParameterSet.HFileOpenSave.filename = f'{dirPath}\\{damName}_{bujeaName}_H.hwp'
        # set save format to "pdf"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        # save
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        # set save filename
        hwp.HParameterSet.HFileOpenSave.filename = f'{dirPath}\\{damName}_{bujeaName}_P.pdf'
        # set save format to "pdf"
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        # save
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        
        # hwp.Quit()  # 한글 종료
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        time.sleep(3) 
        if hwp:
            hwp.Clear(option=1) #오류발생시 한글 버림
            hwp.Quit()
           

   

def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

 



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

if __name__ == '__main__':
    manager = Win32COMCacheManager()

    print(" Temp gen_py 폴더 초기화 및 COM 모듈 재생성 중...")
    manager.clear_cache()
    manager.regenerate_all_modules()
    #231228 1100
    parser = argparse.ArgumentParser(description="Read text from an image and filter by specific conditions")

    # parser.add_argument('--measure-csv', default='../measure/27SYD-5YSR.csv', help="검출 결과 measure 파일")
    parser.add_argument('--measure-csv', default='./SYD/YSR.csv', help="검출 결과 measure 파일")

    # # parser.add_argument('--measure-csv', help="검출 결과 measure 파일")
    parser.add_argument('--output-path', default='241007', help="상태 등급 output 경로")
    #
    parser.add_argument('--config', default='./SYD/27SYD-5YSR-231205.conf', help="시설물 Conf 파일")
    #
    parser.add_argument('--estimatorPath', default='./StateEstimator', help="상태평가보고서 작성용 파일들")

    parser.add_argument('--dam', default='27SYD', help="상태평가보고서 작성용 파일들")
    
    parser.add_argument('--bujea', default='YSR', help="부재 이름")
    parser.add_argument('--attribute', default='', help="부재 속성(필 = F 콘크리트 = C 대청댐 = DCD)")
    parser.add_argument('--shotingday', default='', help="촬영시간")
    parser.add_argument('--mainpath', default='', help="메인 경로")
    args = parser.parse_args()

    file = args.measure_csv
    outputpath = args.output_path
    estimatorPath = args.estimatorPath
    damName = args.dam
    bujeaName =args.bujea
    config1 = args.config
   
    attribute = args.attribute
    mainPath = args.mainpath
    shotingDay = args.shotingday

   


    stage4 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage04\\04"
    stage5 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage05\\05"
    stage6 = mainPath+"\\"+damName+"\\"+shotingDay+"\\"+"stage06\\06"
    bujeaFolderNames = checkBujea(bujeaName)
    allPath = defaultdict(set)
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
  
    

    totalDataFrame2 = []
    file2 = os.path.join(outputpath, '1단계결과.csv')
    csvData2 = pd.read_csv(file2)

    file1_5 = os.path.join(outputpath, '1.5단계결과.csv')
    csvData1_5 = pd.read_csv(file1_5)
   



    Data2 = csvData2[csvData2["StructureType"] == structurename]

  

    tablepath = os.path.join(outputpath, 'table')
    createDirectory(tablepath)
    DefectTableUseTotal.defectTable(structurename, file, tablepath,csvData1_5,outputpath,attribute, fileName)
  
    DamFloor_StateEstimation.step2Grade(structurename, Data2, structureLen, totalDataFrame2,attribute,list(fileName))



    totalFrame2 = pd.DataFrame(totalDataFrame2, columns=('StructureType', 'Unit', 'Grade', 'Evalue'))
    stage2_result = os.path.join(outputpath, '2단계결과.csv')
    totalFrame2.to_csv(stage2_result, index=False)


    totalDataFrame3 = []
    file3 = stage2_result
    stage3_result = os.path.join(outputpath, '3단계결과.csv')

 
 
  
   
    csvData3 = pd.read_csv(file3)
    Data3 = csvData3[csvData3["StructureType"] == structurename].Evalue.tolist()
   

    DamFloorstructureLen3 = len(Data3)


    DamFloor_StateEstimation.step3Grade(structurename, Data3, DamFloorstructureLen3, totalDataFrame3)



    totalFrame3 = pd.DataFrame(totalDataFrame3, columns=('StructureType', 'Evalue', 'Grade'))
    totalFrame3.to_csv(stage3_result, index=False)


    makeHwp(damName,bujeaName,outputpath,file)


