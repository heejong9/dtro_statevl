import os
from collections import defaultdict
from io import BytesIO
import time

import pandas as pd
import win32clipboard
import win32com.client as win32
from PIL import Image
import glob

def init_hwp(visible=True):
    """
    아래아한글 시작
    """
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = visible
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp


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

def resize_image_by_ratio(input_path, scale_factor):
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
    # 이미지 = Image.open(filepath)
    이미지 = resize_image_by_ratio(filepath, 0.1)
    아웃풋 = BytesIO()
    이미지.convert('RGB').save(아웃풋, 'BMP')
    최종데이터 = 아웃풋.getvalue()[14:]
    아웃풋.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, 최종데이터)
    win32clipboard.CloseClipboard()

    hwp.Run("Paste")

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
    hwp.Run("MoveTopLevelEnd")
    hwp.Run("BreakPage")


def 완료메시지():
    msgbox = hwp.XHwpMessageBox  # 메시지박스 생성
    msgbox.string = "문서작성을 완료하였습니다."
    msgbox.Flag = 0  # [확인] 버튼만 나타나게 설정
    msgbox.DoModal()  # 메시지박스 보이기

def get_xlsx_files_in_current_directory():
    # 현재 디렉토리에서 .xlsx 파일 찾기
    xlsx_files = glob.glob(os.path.join(os.getcwd(), '*.xlsx'))

    return xlsx_files

if __name__ == '__main__':
    hwp = init_hwp()
    # hwp.Open(os.path.join(os.getcwd(), "예시파일.hwp"))
    # hwp.Run("FileOpen")
    hwp.Open(r"C:/Users/user/Desktop/statevl/hwp/test1.hwp")
    os.chdir(hwp.Path.rsplit("\\", maxsplit=1)[0])  # 현재 열린 문서가 위치한 경로로 이동
    xlsx_files_list = get_xlsx_files_in_current_directory()

    print('xlsx_files_list', xlsx_files_list)

    for idx, data in enumerate(xlsx_files_list):
        df = pd.read_excel(os.path.join(os.getcwd(), data))
        append_hwp(os.path.join(os.getcwd(), "템플릿1.template"))

        get_nst_table(idx)      
        hwp.FindCtrl() 
        print("Selected Control ID:", hwp.HeadCtrl.CtrlID)
        print("Current Control Position:", hwp.HeadCtrl.GetAnchorPos(0))
        hwp.Run("ShapeObjTableSelCell")
    
       
        hwp.Run("TableLowerCell")
        
        hwp.Run("TableLowerCell")
        
        
        for row in range(len(df)):  # 모든 행을 순회하면서
            for col in range(5):  # 한 행씩 입력
                print("Current Row:", row, "Current Column:", col)
                insert_text(df.iloc[row,col])  # 첫 번째 셀부터 차례대로 입력
       
                if col != 4:  # 마지막 열애 도착하기 전까지는
                    hwp.Run("TableRightCell")  # 입력후 우측셀로
            if len(df) - row != 1:  # 마지막 열에 도착하면
                hwp.Run("TableAppendRow")  # 우측셀로 가지 말고 행 추가
                hwp.Run("TableColBegin")  # 추가한 행의 첫 번째 셀로 이동

        첫번째행으로_이동()
        for i in range(2):
            for i in range(len(df)):
                val = get_text()
                hwp.Run("TableLowerCell")
                lower_val = get_text()
                if val == lower_val and i <= len(df)-3:
                    위셀과병합()
            첫번째행으로_이동()
            hwp.Run("TableRightCell")

        hwp.Run("TableLowerCell")
        hwp.Run("TableRightCell")
        for i in range(6):
            옆셀과병합()
            hwp.Run("TableLowerCell")
            hwp.Run("TableRightCell")

        get_nst_table(idx)
        hwp.FindCtrl()
        hwp.Run("ShapeObjTableSelCell")
        hwp.Run("TableLowerCell")
        hwp.Run("TableLowerCell")
        hwp.Run("TableLowerCell")

        data_name = data.split('\\')[-1]
        img_name = data_name[:2]+data_name[3:7]+data_name[8:10]+'.jpg'
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), img_name))
        insert_txt(data_name.split('_')[0])
        hwp.Run("TableCellBlockExtend")  # 셀 다중선택모드
        hwp.Run("TableUpperCell")  # 위 셀까지 선택
        hwp.Run("TableMergeCell")  # 선택된 셀 병합
        hwp.Run("TableCellBlock")  # 다시 셀 선택 모드
        hwp.Run("TableCellBlockExtend")  # 셀 다중선택모드
        hwp.Run("TableLowerCell")  # 위 셀까지 선택
        hwp.Run("TableMergeCell")  # 선택된 셀 병합

        hwp.Run("Cancel")
        hwp.Run("MoveTopLevelEnd")
        hwp.Run("BreakPage")


    완료메시지()