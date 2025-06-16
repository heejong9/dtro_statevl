from datetime import date
import os
from collections import defaultdict
from io import BytesIO
import time
import configparser
import pandas as pd
import win32clipboard
import win32com.client as win32
from PIL import Image
import glob
bujeaDictionary = {
    "YSR": "여수로",
    "BYR": "보조여수로",
    "DMR": "댐마루",
    "HRM": "하류사면",
    "SRM": "상류사면"
}
abs_path = os.path.abspath(r'./')
def init_hwp(visible=True):
    """
    아래아한글 시작
    """
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = False
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp
hwp = init_hwp()

def get_nst_table(n):
    """
    n번째 표 앞으로 이동
    n은 0부터 시작함.
    """
    ctrl = hwp.HeadCtrl
    count = 0
    found_table = False
    
    while ctrl:
        print(f"컨트롤 ID: {ctrl.CtrlID}, 타입: {type(ctrl.CtrlID)}")
        if ctrl.CtrlID == "tbl":
            if count == n:
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                found_table = True
                break
            else:
                count += 1
        ctrl = ctrl.Next
    
    if not found_table:
        print(f"Warning: There are only {count} tables in the document.")
    current_pos = hwp.GetPos()
    print(f"get_nst_table 후 현재 위치: {current_pos}")

def insert_text(text):
    """
    문서에 텍스트 삽입
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
    hwp.HAction.Run("MoveRight")


def get_text():
    """
    문서 선택범위의 문자열을 추출
    """
    hwp.InitScan(Range=0xff)
    total_text = ""
    state = 2
    while state not in [0, 1]:
        state, text = hwp.GetText()
        total_text += text
    hwp.ReleaseScan()
    return total_text


def append_hwp(filename):
    print(filename)
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


def 사진리스트추출():
    """
    하위의 "그림폴더" 안에서
    모든 그림파일 경로를 추출하여
    리스트로 반환
    """
    pic_list = []
    os.chdir("그림폴더")
    for sub_1 in os.listdir():
        os.chdir(sub_1)
        for sub_2 in os.listdir():
            os.chdir(sub_2)
            for sub_3 in os.listdir():
                os.chdir(sub_3)
                pic_list.append([os.path.join(os.getcwd(), i) for i in os.listdir()])
                os.chdir("..")
            os.chdir("..")
        os.chdir("..")
    os.chdir("..")
    return pic_list

def move_to_page_end(hwp):
    # 현재 페이지에서 다음 페이지로 이동하는 명령어 실행
    while True:
        current_page = hwp.XHwpDocuments.GetCurrentPage()
        hwp.Run("MoveLineDown")  # 한 줄씩 아래로 이동
        next_page = hwp.XHwpDocuments.GetCurrentPage()
        
        if next_page > current_page:  # 페이지가 넘어갔으면 다시 한 줄 위로 올라감
            hwp.Run("MoveLineUp")
            break
        
def 사진제목추출():
    """
    사진파일 경로에서
    학교명, 장소명, 항목, 타입을 추출한 후
    "항목"을 제외한 세 개의 요소를 이용해
    "학교명_장소명_타입" 형태의
    사진표 제목을 list로 반환

    이와 별개로
    제목을 key로,
    제목에 포함되는 사진파일들의 경로를 value로 갖는
    dict도 하나 더 반환
    (이미지 삽입시 사용)
    """
    제목사전 = defaultdict(list)
    제목리스트 = []
    for 사진폴더 in 사진리스트:
        for 사진파일 in 사진폴더:
            제목 = 사진파일.rsplit("\\")[-1]
            학교명, 장소명, 항목, 타입 = 제목[:-4].split("_")
            제목명 = f"{학교명}_{장소명}_{타입}"
            if 제목명 not in 제목리스트:
                제목리스트.append(제목명)
            제목사전[제목명].append(사진파일)
    return 제목리스트, 제목사전

def resize_image_to_box(image_path, box_width_mm, box_height_mm):
    """
    이미지를 비율을 유지한 채 주어진 크기의 상자에 맞게 리사이즈
    box_width_mm: 상자의 가로(mm)
    box_height_mm: 상자의 세로(mm)
    """
    # mm를 픽셀로 변환 (1mm ≈ 3.7795275591px, 96 DPI 기준)
    mm_to_px = 3.7795275591
    box_width_px = int(box_width_mm * mm_to_px)
    box_height_px = int(box_height_mm * mm_to_px)
   
    with Image.open(image_path) as img:
        img_ratio = img.width / img.height
        box_ratio = box_width_px / box_height_px

        if img_ratio > box_ratio:  # 이미지가 더 넓을 경우
            new_width = box_width_px
            new_height = int(box_width_px / img_ratio)
        else:  # 이미지가 더 높을 경우
            new_width = int(box_height_px * img_ratio)
            new_height = box_height_px

        resized_img = img.resize((new_width, new_height), Image.LANCZOS)
        # resized_path = f"resized_{os.path.basename(image_path)}"
        # resized_img.save(resized_path)
        return resized_img  # mm 단위로 반환

def center_image_in_page(hwp):
    """이미지를 페이지 가로축 가운데에 정렬"""
    hwp.HAction.GetDefault("ShapeObjTableCell", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.HorzAlign = 1  # 가로 정렬: 가운데
    hwp.HAction.Execute("ShapeObjTableCell", hwp.HParameterSet.HShapeObject.HSet)
    
def  resize_image_by_ratio(input_path, scale_factor):
    # 이미지 열기
    original_image = Image.open(input_path)

    # 이미지 크기 가져오기
    original_width, original_height = original_image.size

    # 비율에 따라 새로운 크기 계산
    new_width = int(original_width * scale_factor)
    new_height = int(original_height * scale_factor)

    # 이미지 리사이징
    resized_image = original_image.resize((new_width, new_height), Image.LANCZOS)
    
    return resized_image

def 클립보드로_이미지_삽입(filepath):
    """
    한/글 API의 InsertPicture 메서드는
    셀의 크기를 변경하지 않는 반면(이미지가 찌그러짐)
    클립보드를 통해 이미지를 삽입하면
    이미지의 종횡비에 맞춰
    셀의 높이가 자동으로 조절됨.
    """
    #이미지 = Image.open(filepath)
    

    이미지 = resize_image_to_box(filepath, 100,100)
    아웃풋 = BytesIO()
    이미지.convert('RGB').save(아웃풋, 'BMP')
    최종데이터 = 아웃풋.getvalue()[14:]
    아웃풋.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, 최종데이터)
    win32clipboard.CloseClipboard()
    # HWP 실행
    # hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    # hwp.XHwpWindows.Item(0).Visible = True
   
    hwp.Run("Paste")
    
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

    hwp.Run("Cancel")
    
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableCellBlockExtend")
    hwp.HAction.Run("TableCellBlockExtend")
    
    hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행
    hwp.Run("Cancel")

   


 

 
def insert_txt(txt):
    act = hwp.CreateAction("InsertText")  # 액션 변수 생성
    param = act.CreateSet()  # 파라미터셋 변수 생성(구조는 만들어지지만 값이 비어있음)
    act.GetDefault(param)  # 파라미터셋 초기화(현재 상태값으로 채움)
    param.SetItem("Text", txt)  # 파라미터셋 중 원하는 값 수정
    act.Execute(param)  # 파라미터셋 넣고 액션 실행


def 첫번째행으로_이동():
    """
    표의 A2 셀로 이동
    """
    hwp.Run("TableColBegin")
    hwp.Run("TableColPageUp")
    hwp.Run("TableLowerCell")
    hwp.Run("TableLowerCell")
    hwp.Run("TableCellBlock")


def 위셀과병합():
    """
    제목 그대로임
    """
    hwp.Run("TableDeleteCell")  # 아래 셀 내용을 지운 후
    hwp.Run("TableCellBlockExtend")  # 셀 다중선택모드
    hwp.Run("TableUpperCell")  # 위 셀까지 선택
    hwp.Run("TableMergeCell")  # 선택된 셀 병합
    hwp.Run("TableCellBlock")  # 다시 셀 선택 모드

def 옆셀과병합():
    """
    제목 그대로임
    """
    hwp.Run("TableDeleteCell")  # 오른쪽 셀 내용을 지운 후
    hwp.Run("TableCellBlockExtend")  # 셀 다중선택모드
    hwp.Run("TableLeftCell")  # 위 셀까지 선택
    hwp.Run("TableMergeCell")  # 선택된 셀 병합
    hwp.Run("TableCellBlock")  # 다시 셀 선택 모드


def 다음페이지로():
    """
    표에서 나와서
    문서 끝으로 이동 후
    Ctrl-Enter를 통해
    다음페이지로 넘어감
    """
    hwp.Run("Cancel")
    hwp.HAction.Run("MoveDown")
def 다음페이지로2():
    """
    표에서 나와서
    문서 끝으로 이동 후
    Ctrl-Enter를 통해
    다음페이지로 넘어감
    """
    
    hwp.Run("Cancel")
   
    hwp.HAction.Run("MoveDown")    
    hwp.HAction.Run("MoveDown")   
   

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
    
    
    config['Cover']['project_period'] = '2024-10-04'
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



def 완료메시지():
    msgbox = hwp.XHwpMessageBox  # 메시지박스 생성
    msgbox.string = "문서작성을 완료하였습니다."
    msgbox.Flag = 0  # [확인] 버튼만 나타나게 설정
    msgbox.DoModal()  # 메시지박스 보이기

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
    

    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "사연댐 상태평가 보고서"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
  
    hwp.HAction.Run("MoveSelPrevWord")
    hwp.HAction.Run("MoveSelPrevWord")
    hwp.HAction.Run("MoveSelPrevWord")
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(15)
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HAction.Run("CharShapeBold")

    hwp.HAction.Run("MoveDown")
   
 
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HAction.Run("CharShapeNormal")  # 볼드체 취소
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    # 글자 크기를 10으로 설정
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(10)  # 10포인트로 설정
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    for i in range(0,36):
        hwp.HAction.Run("BreakPara")
   
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    today_date = date.today().strftime("%Y년 %m월 %d일")
    hwp.HParameterSet.HInsertText.Text = f"\n\n작성일: {today_date}"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignRight")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = f"작성자 : (주)딥인스펙션"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("BreakPara")

    # 날짜 삽입
    
    
def report(bujea,dirPath):  
    
 
    xlsx_files_list = get_xlsx_files_in_current_directory(dirPath)
    print(xlsx_files_list)


    for idx, data in enumerate(xlsx_files_list):
        file_name = os.path.splitext(os.path.basename(data))[0]
       
        df = pd.read_csv(os.path.join(os.getcwd(), data),encoding='cp949')
        append_hwp(os.path.join(os.getcwd(), "템플릿3.template"))

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
     
        modified_path = dirPath.replace("stage07", "stage06").replace("075", "065")
        클립보드로_이미지_삽입(os.path.join(modified_path, img_name))
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
        
        hwp.Run("MoveDown")
        hwp.Run("BreakPage")



def 글자속성(font_size=10,bold=False):
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.HSet.SetItem("Bold", 1 if bold else 0)  # 진하게 설정 (1: 진하게, 0: 일반)
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(font_size)
    # hwp.HParameterSet.HCharShape.SizeHangul = 100
    # hwp.HParameterSet.HCharShape.RatioHangul = 100
    # hwp.HParameterSet.HCharShape.SpacingHangul = 0
    # hwp.HParameterSet.HCharShape.OffsetHangul = 100
    # hwp.HParameterSet.HCharShape.SizeLatin = 100
    # hwp.HParameterSet.HCharShape.RatioLatin = 100
    # hwp.HParameterSet.HCharShape.SpacingLatin = 0
    # hwp.HParameterSet.HCharShape.OffsetLatin = 100
    # hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameUser = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameSymbol = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameOther = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameJapanese = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameHanja = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameLatin = "함초롬바탕"  # 글자모양 - 글꼴종류
    hwp.HParameterSet.HCharShape.FaceNameHangul = "함초롬바탕"  # 글자모양 - 글꼴종류

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
def 목차():
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    
    글자속성(17,1)
    insert_text("목 차")
    
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("BreakPara")

    # 목차 항목 추가 함수
    def add_toc_item(text, level=0, font_size=10, bold=False):
          # 글자 속성 설정
        글자속성(font_size,bold)
 

        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = "  " * level + text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
       
       
        hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
    # 목차 항목 추가
    add_toc_item("1. 시설물 개요",0,15,True)
    add_toc_item("1.1 시설물 현황", 1,13)
    add_toc_item("1.2 관련 도면", 1,13)
    add_toc_item("2. 상태평가 개요",0,15,True)
    add_toc_item("3. 상태평가 기준 및 방법",0,15,True)
    add_toc_item("3.1 상태평가 항목 및 기준", 1,13)
    add_toc_item("3.1.1 평가유형 영향계수 및 기준산정 방법", 2)
    add_toc_item("3.1.2 상태평가 항목 및 기준", 2)
    add_toc_item("3.2 상태평가 결과 산정 방법", 1,13)
    add_toc_item("3.2.1 댐 시설물 평가 단계별 절차", 2)
    add_toc_item("3.2.2 상태평가 단계별 구분", 2)
    add_toc_item("3.2.3 기계 및 전기 설비", 2)
    add_toc_item("4. 상태평가 결과",0,15,True)

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
    글자속성(15,1)
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
    
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    
    cellDownText(['하천명','댐형식','댐정상 표고',
                '댐 높이','댐 길이','댐 체적',
                '여수로계획홍수량/계획방류량','여수로게이트',
                '여수로형식'])
    
    for _ in range(10):
        hwp.HAction.Run("MoveUp")
    
    hwp.HAction.Run("MoveRight")
    
    cellDownText([config["Cover"]['river_name'],config["Cover"]['damspecifictype'],config["Cover"]['dampeak'],
                config["Cover"]['damheight'],config["Cover"]['damlength'],config["Cover"]['damvolume'],
                config["Cover"]['여수로계획홍수량/계획방류량'],config["Cover"]['여수로게이트'],config["Cover"]['여수로형식'],
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
                '보조여수로계획홍수량/계획방류량','보조여수로게이트',
                '보조여수로형식'])


    for _ in range(12):
        hwp.HAction.Run("MoveUp")
    
    hwp.HAction.Run("MoveRight")
    
    cellDownText([config["Cover"]['총저수량'],config["Cover"]['유효저수량'],config["Cover"]['유역면적'],
                config["Cover"]['계획홍수위'],config["Cover"]['상시만수위'],config["Cover"]['저수위'],
                config["Cover"]['보조여수로계획홍수량/계획방류량'],config["Cover"]['보조여수로게이트'],config["Cover"]['보조여수로형식'],
                ])
    
def 현황표():
    
    글자속성(17,1)
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    insert_text(f"{config['Cover']['DamName']} 현황표")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
    crateTable(13,6)
    
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
    
    cellLowerMerge(3)
    
    for _ in range(3):
        hwp.HAction.Run("MoveRight")
    for _ in range(6):
        hwp.HAction.Run("MoveUp")
        
    cellLowerMerge(6)
    
    hwp.HAction.Run("MoveDown")
    
    cellLowerMerge(3)
    
    hwp.Run("TableColBegin")
    for _ in range(5):
        hwp.HAction.Run("MoveUp")
        
    textTable()

def 시설물현황():
    글자속성(17,1)
    insert_text("1. 시설물 개요")
    글자속성()
    hwp.HAction.Run("BreakPara")
    insert_text(f"{config['Cover']['damname']}은 {config['Cover']['project_location']}에 위치하고, 높이 {config['Cover']['damheight']}, 길이 {config['Cover']['damlength']}의 {config['Cover']['damspecifictype']}으로\
{config['Cover']['Project_completiondate']}년에 준공되었다. 댐의 제원 및 시설물의 현황은 다음과 같다.")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text("1.1 시설물 현황")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text("1.1.1 시설물 현황")
    글자속성()
    hwp.HAction.Run("BreakPara")
    insert_text(f"{config['Cover']['damname']}의 기본 현황은 [표 2.1]과 같다.")
    hwp.HAction.Run("BreakPara")

    hwp.HAction.Run("BreakPara")
    현황표()
    
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    insert_text("[표 2.1] 시설물 현황.")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
    hwp.Run("BreakPage")
    글자속성(15,1)
    insert_text(" 1.1.2. 관련도면")
    hwp.HAction.Run("BreakPara")
    ##사진 표
    crateTable(1,1)
    
    클립보드로_이미지_삽입(f"StateEstimator\\그림\\그림2_2.png")
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    insert_text(f" [그림2.1]{config['Cover']['facility_name']} 종평면도")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
    hwp.Run("BreakPage")
    ##사진 표
    crateTable(1,1)   
    클립보드로_이미지_삽입(("StateEstimator\\그림\\그림2_3.png"))
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    insert_text(f"[그림2.3]{config['Cover']['facility_name']} 표준 단면도")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
    hwp.Run("BreakPage")
    글자속성(17,1)
    insert_text("2. 상태평가 개요 ")
    글자속성()
    hwp.HAction.Run("BreakPara")

    insert_text(f"시설물의 상태평가는 「시설물의 안전 및 유지관리 실시 지침\
국토교통부, 국토안전관리원)」에 따라 실시하며, 상태평가에 대한 세부적인 사항은\
「시설물의 안전 및 유지관리 실시 세부지침(안전점검·진단편의 댐편, 2021.12)」을 준용하였다.\
따라서, 「시설물의 안전 및 유지관리 실시 세부지침(안전점검·진단편의 댐편)」에 의거하여 외관조사\
및 내구성 조사의 항목 및 수량에 따라 과업을 실시한 후, 상태평가를 위하여 중요 손상 및 결함을\
세부 기준에 의해 분류하고 평가기법 및 절차에 따라 각 개별시설물에 대한 결함의 등급과 점수 및 지수를 산정하여 상태등급을 최종적으로 결정하였다.")
    hwp.Run("BreakPage")  # 페이지 나누기 삽입
    글자속성(17,1)
    insert_text("3. 상태평가 기준 및 방법 ")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text("3.1 상태평가 항목 및 기준 ")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text("3.1.1 평가유형·영향계수 및 기준산정 방법")
    글자속성()
    hwp.HAction.Run("BreakPara")
    text="시설물의 상태평가는 결함 및 손상에 따른 각각의 상태평가 기준을 적용하며,\n\
상태변화가 전체 구조물에 미치는 안전성의 영향정도, 구조적인 중요도가 적절히 고려되어 평가될 수 있도록 결함 및 손상을 평가유형(評價類型)별로 구분하여 영향계수를 적용한다. \n \
1) 평가유형의 구분 \n \
결함 및 손상에 대한 평가유형은 다음과 같이 구분한다.\n\
① 중요결함\n\
침하, 경사\전도 및 활동 등과 같이 전체 구조물의 구조적인 안전에 직접 영향을 미치는 결함\n\
② 국부결함\n \
수평이음부 불량 등과 같이 구조물의 안전성에 직접적인 영향을 미치지는 않지만 손상이 진전될 경우 전체 구조물의 안전에 상당한 영향을 끼칠 수 있는 결함\n\
③ 일반손상\n\
파손, 마모, 콘크리트 재료분리 등과 같이 구조물의 안전에 크게 영향을 주지 않는 일반적인 손상\n\
2) 영향계수의 적용\n\
각 부재에서 발생하는 각종 손상 및 결함에 대한 상태평가 시 손상이 전체 구조물에 미치는 안전성의 영향정도, 구조적인 중요도가 적절히 고려되어 평가될 수 있도록 영향계수를 적용한다.영향계수는 안전성에 직접적인 영향을 미치는 중요 결함의 상태등급을 기준으로 하여 국부적인 결함의 등급을 상향 조정함으로써 이들이 전체 구조물에 미치는 영향을 평가 절하하는 계수이며, 영향계수는 상태평가를 위한 표준기준이며, 책임기술자의 판단으로 다소 조정할 수 있다."
    # 텍스트를 줄바꿈 문자('\n')로 분리
    lines = text.split('\n')

    # 한글 문서에 텍스트 삽입 및 줄바꿈 처리
    for line in lines:
        insert_text(line)
        hwp.HAction.Run("BreakLine")  # 줄바꿈 삽입
    hwp.HAction.Run("BreakPara")
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
    hwp.HAction.Run("MoveRight")
    for _ in range(26):
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
   
  
    hwp.HAction.Run("MoveLeft")
    for _ in range(6):
        hwp.HAction.Run("MoveUp")
    cellDownText(['상태변화','균열','박리','박락','누수','백태','철근노출'])
    
    for _ in range(7):
        hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    cellDownText(['평가유형','국부결함','국부결함','국부결함','국부결함','일반손상','국부결함'])
    
    for _ in range(7):
        hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    cellDownText([
'영향계수','영향계수',
'1.0','1.1','1.2','1.4','2.0',
'1.0','1.1','1.2','1.4','2.0',
'1.0','1.1','1.2','1.4','2.0',
'1.0','1.1','1.2','1.4','2.0',
'1.0','1.1','1.3','1.7','3.0',
'1.0','1.1','1.2','1.4','2.0',
])

    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    for _ in range(31):
        
        hwp.HAction.Run("MoveUp")
    cellDownText([
'평가기준','평가기준',
'A','B','C','D','E',
'A','B','C','D','E',
'A','B','C','D','E',
'A','B','C','D','E',
'A','B','C','D','E',
'A','B','C','D','E',
])
   
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    for _ in range(31):
       
        hwp.HAction.Run("MoveUp")
    cellDownText([
'평가점수','평가점수',
'5','4','3','2','1',
'5','4','3','2','1',
'5','4','3','2','1',
'5','4','3','2','1',
'5','4','3','2','1',
'5','4','3','2','1',
])
    
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveRight")
    for _ in range(31):
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
    
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("TableCellBlockExtend")
    
    hwp.HAction.Run("TableCellBlockExtend")

    # hwp.HAction.Run("TableDistributeCellHeight")
    # hwp.HAction.Run("TableDistributeCellWidth")
    hwp.HAction.Run("TableResizeExLeft")
    hwp.HAction.Run("TableResizeExLeft")
    hwp.HAction.Run("TableResizeExLeft")
    hwp.HAction.Run("Close")
def 상태평가기준():
    글자속성(15,1)
    insert_text("3.1.2 상태평가 항목 및 기준 ")
    글자속성()
    hwp.HAction.Run("BreakPara")
    insert_text(f"본 과업대상 시설물인{config['Cover']['facility_name']}은{config['Cover']['Project_FacilityType']}형식으로{config['Cover']['Project_FacilityClass']}에 준하는 기준을 적용하였다.")
    hwp.HAction.Run("BreakPara")
    insert_text("「세부지침」에 준하여 정량적이고 객관적인 상태평가를 위하여 부재별, 개별부재별, 복합부재별, 개별시설별 각 부재별 상태평가 항목은 다음과 같다.")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text(f"{config['Cover']['facility_name']}")
    hwp.HAction.Run("BreakPara")
    ## 표 작성
    crateTable(32,9)
    기준표()
    hwp.Run("BreakPage")  # 페이지 나누기 삽입
    글자속성(15,1)
    insert_text("3.2 상태평가 결과 산정 방법 ")
    글자속성()
    hwp.HAction.Run("BreakPara")
    insert_text("3.2.1 댐 시설물 평가 단계별 절차 ")
    hwp.HAction.Run("BreakPara")
    insert_text("댐 시설물에 대한 상태평가는 [그림 3.1]과 같이 단계별로 구분할 때 댐 시설물은 통합시설물 (6단계) 에 해당하는 시설물로서 간주하고, 하위단계인 복합시설, 개별시설, 복합부재, 개별부재로 구분한다.외관조사망도는 개별부재에 대하여 작성하는 것을 원칙으로 하고 필요시 개별부재의 크기, 면적에 따라 부위별로 분할하여 작성한다.")
    hwp.HAction.Run("BreakPara")
    insert_text("3.2.2 상태평가 단계별 구분")
    hwp.HAction.Run("BreakPara")
    insert_text("시설물의 상태를 평가하기 위하여 시설물을 단계별로 구분하여 다음 표와 같이 평가단계별 구분표를 작성하고 본 보고서에 수록한다.")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    insert_text("[표3.11] 댐 시설물의 상태평가 단계별 구분표(예시)")
    hwp.HAction.Run("BreakPara")
    hwp.HAction.Run("ParagraphShapeAlignLeft")
   
    ##사진필요
    crateTable(1,1)
    클립보드로_이미지_삽입("StateEstimator\\그림\\표3_11.png")
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("BreakPara")
    글자속성(15,1)
    insert_text("1) 1단계 상태평가 : 부재(部材)별 손상상태 평가표 작성 ")
    글자속성()
    insert_text("시설물의 상태평가 단계별 구분표에 따라 개별부재를 1개 외관조사망도 또는 필요에 따라 부위별로 다수의 외관조사망도로 구분하여 개략도에 손상 및 결함상태를 도시하고, 조사결과표에 개별부재에 대한 손상내용을 상세히 기록한 후, 그 손상 정도에 대하여 5단계(a～e) 상태평가 결과 및 평가점수를 부여한다.")
    hwp.HAction.Run("BreakPara")
    insert_text("○  손상상태 평가표에는 평가항목에 없는 상태변화라 할지라도 모두 기록하는 것을 원칙으로 한다.")
    hwp.HAction.Run("BreakPara")
    insert_text("○ 각 상태변화에 대한 상태평가 결과가 c, d, e 등급일 경우 보수ㆍ보강 우선순위에 따라 보수ㆍ보강을 한다.")


if __name__ == '__main__':
    config = configparser.ConfigParser()
    with open('./statevl/SYD/27SYD_YSR.conf', 'r', encoding='utf-8') as configfile:
        config.read_file(configfile)


    # main_page()
    # hwp.Run("BreakPage")  # 페이지 나누기 삽입
    # 목차()
    # hwp.Run("BreakPage")
    # 안전진단표()
    # 다음페이지로2()
    # 결과요약()
    # 다음페이지로()
   
 
    # hwp.Run("BreakPage")  # 페이지 나누기 삽입
    # 시설물현황()

    # hwp.Run("BreakPage")  # 페이지 나누기 삽입
    # 상태평가기준()
    # hwp.Run("BreakPage")
    report("YSR")
config = configparser.ConfigParser()    

def makeHwp(damName,bujeaName,dirPath):
    with open(f'./{damName}_{bujeaName}.conf', 'r', encoding='utf-8') as configfile:
        config.read_file(configfile)


  
    main_page()
    hwp.Run("BreakPage")  # 페이지 나누기 삽입
    목차()
    hwp.Run("BreakPage")
    안전진단표()
    다음페이지로2()
    결과요약()
    다음페이지로()
   
 
    hwp.Run("BreakPage")  # 페이지 나누기 삽입
    시설물현황()

    hwp.Run("BreakPage")  # 페이지 나누기 삽입
    상태평가기준()
    hwp.Run("BreakPage")
    report(bujeaName,dirPath)
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

    
    hwp.Quit()  # 한글 종료

   