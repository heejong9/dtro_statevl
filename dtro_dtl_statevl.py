import json
import os
import time
from collections import defaultdict
from datetime import date, datetime
import pandas as pd
import win32com.client as win32
from PIL import Image
import argparse
import db_config


# DB에서 데이터 추출
def fetch_data_from_db(project_id, sub_project_id):
    connection = db_config.connect_db()
    cursor = connection.cursor(dictionary=True)

    # 1) 프로젝트 설정값 가져오기
    cursor.execute(
        """
        SELECT SETTING_VALUE
        FROM PROJECT
        WHERE PROJECT_ID = %s;
    """,
        (project_id,),
    )
    row = cursor.fetchone()
    setting_value = (
        json.loads(row["SETTING_VALUE"]) if row and row.get("SETTING_VALUE") else {}
    )

    # 2) 결함 정보 가져오기 (sub_project_id 기준 필터: 상/하행 모두)
    # sub_project_id 예: "ST3_MPY_PSS" (방향 없음)
    sub_id = (sub_project_id or "").strip().upper() or None

    sub_id_up = f"{sub_id}_UP"
    sub_id_down = f"{sub_id}_DOWN"

    query_defects = """
        SELECT
            DDI.DEFECT_DETECT_ID,
            DDI.IMAGE_NUM,
            DDID.TYPE AS DEFECT_TYPE,
            DDI.DEFECT_IMAGE_EXIST_LEFT_UP    AS LEFT_UP_EXIST,
            DDI.DEFECT_IMAGE_EXIST_RIGHT_UP   AS RIGHT_UP_EXIST,
            DDI.DEFECT_IMAGE_EXIST_LEFT_DOWN  AS LEFT_DOWN_EXIST,
            DDI.DEFECT_IMAGE_EXIST_RIGHT_DOWN AS RIGHT_DOWN_EXIST
        FROM DEFECT_DETECT_IMAGE AS DDI
        LEFT JOIN DEFECT_DETECT_IMAGE_DETAIL AS DDID
        ON DDID.DEFECT_DETECT_IMAGE_ID = DDI.DEFECT_DETECT_IMAGE_ID
        JOIN DEFECT_DETECT AS DD
        ON DDI.DEFECT_DETECT_ID = DD.DEFECT_DETECT_ID
        WHERE DD.PROJECT_ID = %s
        AND DD.SUB_ID IN (%s, %s)
        ORDER BY DDI.IMAGE_NUM ASC;
    """
    cursor.execute(query_defects, (project_id, sub_id_up, sub_id_down))
    defect_rows = cursor.fetchall()

    # detect_id별 애자 개수 집계
    insulator_count = """
        SELECT DDI.DEFECT_DETECT_ID, COUNT(*) AS INS_CNT
        FROM DEFECT_DETECT_IMAGE AS DDI
        JOIN DEFECT_DETECT AS DD
        ON DDI.DEFECT_DETECT_ID = DD.DEFECT_DETECT_ID
        WHERE DD.PROJECT_ID = %s
        AND DD.SUB_ID IN (%s, %s)
        GROUP BY DDI.DEFECT_DETECT_ID
    """
    cursor.execute(insulator_count, (project_id, sub_id_up, sub_id_down))
    cnt_rows = cursor.fetchall()

    CNT_BY_DETECT = {r["DEFECT_DETECT_ID"]: int(r["INS_CNT"]) for r in cnt_rows}
    # 3) 역 정보 가져오기
    cursor.execute("SELECT INITIAL, NAME, LINE, STATION_ORDER FROM SUBWAY_STATIONS;")
    station_rows = cursor.fetchall()
    STATION_MAP = {r["INITIAL"]: r["NAME"] for r in station_rows}
    STATION_ORDER = {
        (row["LINE"], row["INITIAL"]): int(row["STATION_ORDER"]) for row in station_rows
    }

    cursor.close()
    connection.close()

    # === 한글명 매핑 변환 ===
    DEFECT_TYPE_MAP = {
        "Crack": "균열",
        "Damage": "파손",
        "ArchornPeeling": "아크혼_박리",
        "ArchornScorching": "아크혼_그을음",
        "Stain": "얼룩",
        "crack": "균열",
        "damaged": "파손",
        "damage": "파손",
        "archorn_peeling": "아크혼_박리",
        "archorn_soot": "아크혼_그을음",
        "stain": "얼룩",
        "etc": "기타결함",
    }
    LINE_MAP = {"ST1": "1호선", "ST2": "2호선", "ST3": "3호선"}

    for r in defect_rows:
        r["DEFECT_TYPE_KR"] = DEFECT_TYPE_MAP.get(r["DEFECT_TYPE"], r["DEFECT_TYPE"])
        parts = (r["IMAGE_NUM"] or "").split("_")
        if len(parts) >= 5:
            line, from_init, to_init, direction, defect, ins_no = parts
            r["LINE_KR"] = LINE_MAP.get(line, line)
            r["FROM_STATION"] = STATION_MAP.get(from_init, from_init)
            r["TO_STATION"] = STATION_MAP.get(to_init, to_init)
            r["FROM_ORDER"] = STATION_ORDER.get((line, from_init), 10**9)
            r["TO_ORDER"] = STATION_ORDER.get((line, to_init), 10**9)
            r["DIRECTION"] = (
                "상행"
                if direction.upper() == "UP"
                else ("하행" if direction.upper() == "DOWN" else direction)
            )
            r["INSULATOR_NO"] = str(int(ins_no))  # "0001" -> "1"
            r["SUB_PROJECT_ID"] = "_".join(parts[:-2])  # ST3_MPY_PSS_UP
            r["SUBPROJECT_KR"] = (
                f"{r['FROM_STATION']}~{r['TO_STATION']}({r['DIRECTION']})"
            )

    # 애자 개수 맵 
    subproject_insulator_counts = {}

    insulator_map = {f"{sub_id}_UP", f"{sub_id}_DOWN"}

    # SUB_PROJECT_ID → 해당 detect_id들의 후보 집합 만들기
    from collections import defaultdict
    sp_to_detects = defaultdict(set)
    for r in defect_rows:
        sp = r.get("SUB_PROJECT_ID")
        did = r.get("DEFECT_DETECT_ID")
        if sp and did in CNT_BY_DETECT:
            sp_to_detects[sp].add(did)

    # 각 SUB_PROJECT_ID별 최종 애자 개수 선택(여러 detect가 있으면 최댓값 사용)
    for sp in insulator_map:
        candidates = sp_to_detects.get(sp, set())
        if not candidates:
            continue
        subproject_insulator_counts[sp] = max(CNT_BY_DETECT[did] for did in candidates)


    # ====== 표지 메타 구성 ======
    # PROECT_ID_DESC_KR  1호선(월배~상인) 에서 가져오기
    project_id_desc_eng = sub_project_id

    # 토큰 파싱: ST1_SOH_AJR → ["ST1","SOH","AJR"]
    parts = project_id_desc_eng.split("_")
    if len(parts) >= 3:
        line_code  = parts[0]       # ST1
        from_init  = parts[1]       # SOH
        to_init    = parts[2]       # AJR

        line_name     = LINE_MAP.get(line_code, line_code)  # 1호선
        from_station  = STATION_MAP.get(from_init, from_init)  # 송현
        to_station    = STATION_MAP.get(to_init, to_init)      # 안지랑

        section_core  = f"{from_station}~{to_station}"
    else:
        line_name = project_id_desc_eng
        from_station = to_station = ""
        section_core = ""

    # 3) 표지 필드들
    facility_name = setting_value.get("facilityName", "")
    managed_number = setting_value.get("managedNumber", "")
    inspector_raw = setting_value.get("inspector", "")
    inspector = ", ".join([x.strip() for x in inspector_raw.split(",") if x.strip()])
    approver = setting_value.get("approver", "")
    camera_mode = setting_value.get("cameraMode", "")
    writer = "Deep Inspector(AI 안전점검 프로그램)"
    written_date = date.today().strftime("%Y년 %m월 %d일")

    # 4) 제목 라인 (대괄호 표기)
    # 예: 대구 지하철 [2]호선 [수성알파시티] ~ [정평] 상태평가 보고서
    try:
        left, right = [s.strip() for s in section_core.split("~", 1)]
    except Exception:
        left, right = section_core, ""

    title_line = f"대구 지하철 {line_name.replace('호선','')}호선 {left} ~ {right} (상·하행) 상세 상태평가 보고서"

    cover_meta = {
        "title_line": title_line,
        "facility_name": facility_name,
        "managed_number": managed_number,
        "written_date": written_date,
        "place": f"{line_name} {section_core}",  # 검사장소
        "inspector": inspector,
        "camera_mode": camera_mode,
        "writer": writer,
        "approver": approver,
        "line_name": line_name,
        "section_core": section_core,
        "from_station": from_station,
        "to_station": to_station,
    }
    print(subproject_insulator_counts)
    return cover_meta, subproject_insulator_counts, defect_rows


def init_hwp(visible=True):
    """
    아래아한글 시작
    """
    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        print("한글 실행 성공")
    except Exception as e:
        print("실행 실패:", e)
    hwp.XHwpWindows.Item(0).Visible = visible
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp


def makeHwp(root_dir, project_id, sub_project_id):
    """
    ================================
             기본 동작 정의
    ================================
    """

    def 글자속성(font_size=11, bold=False, color=None):
        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
        hwp.HParameterSet.HCharShape.HSet.SetItem(
            "Bold", 1 if bold else 0
        )  # 진하게 설정 (1: 진하게, 0: 일반)
        hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(font_size)

        hwp.HParameterSet.HCharShape.FaceNameUser = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameSymbol = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameOther = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameJapanese = (
            "한컴바탕"  # 글자모양 - 글꼴종류
        )
        hwp.HParameterSet.HCharShape.FaceNameHanja = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameLatin = "한컴바탕"  # 글자모양 - 글꼴종류
        hwp.HParameterSet.HCharShape.FaceNameHangul = "한컴바탕"  # 글자모양 - 글꼴종류

        hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입
        hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType(
            "TTF"
        )  # 글자모양 - 폰트타입

        if color:
            if color.lower() == "red":
                hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(255, 0, 0)
            elif color.lower() == "black":
                hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(0, 0, 0)

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

    def insert_text(text, level=0):
        """
        문서에 제목 스타일을 적용한 후 텍스트 삽입
        """
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    def 들여쓰기(number):
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

        # 들여쓰기 설정 (첫 줄 들여쓰기만 적용)
        hwp.HParameterSet.HParaShape.Indentation = number

        # 문단 모양 적용
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

    def 여백생성(number):
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

        # 단위 변환: pt → HWP 내부 단위 (1pt = 100 HWPUnit)
        hwp.HParameterSet.HParaShape.LeftMargin = number * 100 * 2
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

    def 클립보드로_이미지_삽입(filepath, width, height):

        # 이미지 삽입
        hwp.InsertPicture(filepath, True, 1, False, False, 0, width, height)

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
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행

    def 원본비율유지_이미지_삽입(filepath):
        """비율을 유지하며 이미지 삽압"""

        # 이미지의 가로세로 크기 받아오기
        try:
            with Image.open(filepath) as im:
                w_px, h_px = im.size
        except Exception as e:
            # 파일을 못 열면 텍스트로 대체
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_text("이미지를 열 수 없습니다.")
            return

        # 비율 계산
        width = 75  # 가로 길이 고정
        height = width * (h_px / w_px)

        # 이미지 삽입
        hwp.InsertPicture(filepath, True, 1, False, False, 0, width, height)

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
        hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬 실행

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

    def createTable(rows, cols):
        # 테이블 생성
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

        max_width = 16  # 최대 테이블 폭 (단위: HWP 내부 단위)

        # HTableCreation 파라미터 설정
        hwp.HParameterSet.HTableCreation.Rows = rows
        hwp.HParameterSet.HTableCreation.Cols = cols
        hwp.HParameterSet.HTableCreation.WidthType = (
            max_width / cols
        )  # 테이블 폭을 열 수에 맞게 균등 분할
        hwp.HParameterSet.HTableCreation.HeightType = 0
        hwp.HParameterSet.HTableCreation.WidthValue = 0.0
        hwp.HParameterSet.HTableCreation.HeightValue = 0.0
        # 테이블 폭 설정
        hwp.HParameterSet.HTableCreation.TableProperties.Width = max_width
        # 테이블 생성 액션 실행
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        hwp.HParameterSet.HShapeObject.HSet.SetItem("TreatAsChar", 1)
        hwp.HAction.Run("TableColWidthEven")

    def resizeTable(hwp, left_count=0, down_count=0):
        """
        표 셀 블록 확장 후 크기 조정
        - extend는 반드시 2번 실행
        - left_count: TableResizeExLeft 실행 횟수
        - down_count: TableResizeExDown 실행 횟수
        """
        # 블록 확장 (2번 고정)
        hwp.HAction.Run("TableCellBlockExtendAbs")
        hwp.HAction.Run("TableCellBlockExtend")

        # 왼쪽으로 줄이기
        for _ in range(left_count):
            hwp.HAction.Run("TableResizeExLeft")

        # 아래로 줄이기
        for _ in range(down_count):
            hwp.HAction.Run("TableResizeExDown")

    # 표 2,3,4 전용 표
    def resizeTable2(hwp, left_count=0, up_count=0):
        """
        표 셀 블록 확장 후 크기 조정
        - extend는 반드시 2번 실행
        - left_count: TableResizeExLeft 실행 횟수
        - down_count: TableResizeExDown 실행 횟수
        """
        # 블록 확장 (2번 고정)
        hwp.HAction.Run("TableCellBlockExtendAbs")
        hwp.HAction.Run("TableCellBlockExtend")

        # 왼쪽으로 줄이기
        for _ in range(left_count):
            hwp.HAction.Run("TableResizeExLeft")

        # 위로 줄이기
        for _ in range(up_count):
            hwp.HAction.Run("TableResizeExUp")

    """
    ========================================
            대구교통공사 전용 동작 정의
    ========================================
    """

    # 상태평가 요약 내용 제작
    def get_insulator_defect_summary(project_id, sub_project_id):
        # DB에서 데이터 가져오기
        cover_meta, insulator_count_map, defect_rows = fetch_data_from_db(
            project_id, sub_project_id
        )
        camera_mode = str(cover_meta.get("camera_mode", "2CH")).upper()
        # DataFrame 변환
        df = pd.DataFrame(defect_rows)

        df["START_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].min(axis=1)
        df["END_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].max(axis=1)

        
        # ▼ 방향 코드 우선순위 추가: UP=0, DOWN=1, 기타=2
        df["DIR_CODE"]  = df["SUB_PROJECT_ID"].str.split("_").str[-1].str.upper()
        df["DIR_ORDER"] = df["DIR_CODE"].map({"UP": 0, "DOWN": 1}).fillna(2).astype(int)

        # 구간별( SUB_PROJECT_ID ) 대표 정렬키 뽑기
        order_df = (
            df[["SUB_PROJECT_ID", "LINE_KR", "START_ORDER", "END_ORDER", "DIR_ORDER"]]
            .drop_duplicates()
            .sort_values(["LINE_KR", "START_ORDER", "END_ORDER", "DIR_ORDER"])
        )
        # 정렬 인덱스 맵: SUB_PROJECT_ID 순서
        order_idx = {sp: i for i, sp in enumerate(order_df["SUB_PROJECT_ID"].tolist())}

        # 주요/주의 결함 분류
        major_defects = {"균열", "파손"}  # 교체 필요
        caution_defects = {
            "얼룩",
            "아크혼_박리",
            "아크혼_그을음",
            "기타결함",
        }  # 주의 필요

        # 결과 저장용 리스트
        report_rows = []

        NOT_FIND_LABELS = []
        if camera_mode == "2CH":
            NOT_FIND_LABELS = [
                ("LEFT_DOWN_EXIST", "좌하단 촬영불가지역"),
                ("RIGHT_DOWN_EXIST", "우하단 촬영불가지역"),
            ]

            MAJOR_ORDER = ["균열", "파손", "좌하단 촬영불가지역", "우하단 촬영불가지역"]
            CAUTION_ORDER = ["아크혼_박리", "아크혼_그을음", "얼룩", "기타결함"]
        else:
            NOT_FIND_LABELS = [
                ("LEFT_UP_EXIST", "좌상단 촬영불가지역"),
                ("RIGHT_UP_EXIST", "우상단 촬영불가지역"),
                ("LEFT_DOWN_EXIST", "좌하단 촬영불가지역"),
                ("RIGHT_DOWN_EXIST", "우하단 촬영불가지역"),
            ]

            MAJOR_ORDER = [
                "균열",
                "파손",
                "좌상단 촬영불가지역",
                "우상단 촬영불가지역",
                "좌하단 촬영불가지역",
                "우하단 촬영불가지역",
            ]
            CAUTION_ORDER = ["아크혼_박리", "아크혼_그을음", "얼룩", "기타결함"]

        # 전체 서브프로젝트이름, 애자 갯수로 반복문
        for sub_project, total_count in sorted(
            insulator_count_map.items(), key=lambda kv: order_idx.get(kv[0], 10**9)
        ):

            # 현재 루프에 해당하는 구간만 필터링
            df_sub = df[df["SUB_PROJECT_ID"] == sub_project]

            if not df_sub.empty:
                # 한글 역명은 집합으로 모아서 가장 앞/뒤 선택
                from_station = df_sub["FROM_STATION"].iloc[0]
                to_station = df_sub["TO_STATION"].iloc[0]

                # 방향은 SUB_PROJECT_ID 마지막 토큰으로 확정
                direction_code = sub_project.split("_")[-1].upper()
                direction_map = {"UP": "상행", "DOWN": "하행"}
                direction_kr = direction_map.get(direction_code, direction_code)

                subproject_kr = f"{from_station}~{to_station}({direction_kr})"
            else:
                # fallback
                parts = sub_project.split("_")
                if len(parts) >= 4:
                    direction = parts[3] if len(parts) > 3 else ""
                    direction_map = {"UP": "상행", "DOWN": "하행"}
                    from_station = cover_meta["from_station"]
                    to_station = cover_meta["to_station"]
                    direction_kr = direction_map.get(direction.upper(), direction)
                    subproject_kr = f"{from_station}~{to_station}({direction_kr})"
                else:
                    subproject_kr = sub_project

            # 결함 집계 (결함종류별로 애자번호 묶기)
            majors_map = defaultdict(set)  # 예: {"균열": {2,4,22}, "파손": {5}}
            cautions_map = defaultdict(set)  # 예: {"아크혼_박리": {2,3}, "얼룩": {7,8}}

            # INSULATOR_NO 정렬(문자 '0001' 대응)
            def _to_int_no(x):
                return int(str(x).lstrip("0") or "0")

            for ins_no in sorted(df_sub["INSULATOR_NO"].unique(), key=_to_int_no):
                mask = df_sub["INSULATOR_NO"] == ins_no
                dset = set(
                    df_sub.loc[
                        df_sub["INSULATOR_NO"] == ins_no, "DEFECT_TYPE_KR"
                    ].tolist()
                )
                ins_num = _to_int_no(ins_no)

                # 한 애자의 결함들 모으기
                for d in dset:
                    if d in major_defects:
                        majors_map[d].add(ins_num)
                    if d in caution_defects:
                        cautions_map[d].add(ins_num)

                cols = [
                    "LEFT_UP_EXIST",
                    "RIGHT_UP_EXIST",
                    "LEFT_DOWN_EXIST",
                    "RIGHT_DOWN_EXIST",
                ]
                sub_flags = (
                    df_sub.loc[mask, cols]
                    .astype(str)
                    .apply(lambda s: s.str.upper().fillna("N"))
                )

                for col, label in NOT_FIND_LABELS:
                    # 이 애자(ins_no)에 대해, 해당 위치 컬럼에 'N'이 하나라도 있으면 촬영불가지역로 간주
                    if (sub_flags[col] == "N").any():
                        majors_map[label].add(ins_num)

            def _format_grouped(by_defect: dict, order_list: list) -> str:
                if not by_defect:
                    return "정상상태"
                parts = []
                ordered = [k for k in order_list if k in by_defect] + [
                    k for k in sorted(by_defect.keys()) if k not in order_list
                ]
                for k in ordered:
                    nums = ", ".join(f"#{n}" for n in sorted(by_defect[k]))
                    parts.append(f"@{k}:@{nums}")
                return "@".join(parts)

            # 점검결과 / 비고 텍스트
            major_text = _format_grouped(majors_map, MAJOR_ORDER)
            caution_text = _format_grouped(cautions_map, CAUTION_ORDER)

            # 요약 행 추가
            report_rows.append(
                [
                    "지지물",
                    f"{subproject_kr} #1~#{total_count}",
                    "지지애자@/손상여부",
                    "특이상태 점검@/AI점검",
                    major_text,
                    caution_text,
                ]
            )

        return report_rows

    # 결과표에 필요한 내용 제작
    def get_insulator_defect_details(project_id, sub_project_id):
        cover_meta, insulator_count_map, defect_rows = fetch_data_from_db(
            project_id, sub_project_id
        )
        camera_mode = str(cover_meta.get("camera_mode", "2CH")).upper()

        # 촬영불가지역 라벨 (카메라 모드별)
        if camera_mode == "2CH":
            notfind_cols = [
                ("LEFT_DOWN_EXIST", "좌하단 촬영불가지역"),
                ("RIGHT_DOWN_EXIST", "우하단 촬영불가지역"),
            ]
        else:
            notfind_cols = [
                ("LEFT_UP_EXIST", "좌상단 촬영불가지역"),
                ("RIGHT_UP_EXIST", "우상단 촬영불가지역"),
                ("LEFT_DOWN_EXIST", "좌하단 촬영불가지역"),
                ("RIGHT_DOWN_EXIST", "우하단 촬영불가지역"),
            ]

        # 분류 집합 & 출력 순서
        majors_order = ["균열", "파손"]
        minors_order = ["얼룩", "아크혼_박리", "아크혼_그을음"]

        df = pd.DataFrame(defect_rows)

        # 정렬 키
        df["START_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].min(axis=1)
        df["END_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].max(axis=1)
        df["DIR_CODE"] = df["SUB_PROJECT_ID"].str.split("_").str[-1].str.upper()
        df["DIR_ORDER"] = df["DIR_CODE"].map({"UP": 0, "DOWN": 1}).fillna(2).astype(int)

        order = (
            df[
                [
                    "SUB_PROJECT_ID",
                    "SUBPROJECT_KR",
                    "LINE_KR",
                    "START_ORDER",
                    "END_ORDER",
                    "DIR_ORDER",
                ]
            ]
            .drop_duplicates()
            .sort_values(["LINE_KR", "START_ORDER", "END_ORDER", "DIR_ORDER"])
        )

        rows = []

        for sub_project, subproject_kr in zip(
            order["SUB_PROJECT_ID"], order["SUBPROJECT_KR"]
        ):
            total = insulator_count_map.get(sub_project)
            if not total:
                continue

            df_sub = df[df["SUB_PROJECT_ID"] == sub_project]

            for n in range(1, int(total) + 1):
                df_one = df_sub[df_sub["INSULATOR_NO"] == str(n)]
                defects = set(df_one["DEFECT_TYPE_KR"].dropna().tolist())

                majors_found = [t for t in majors_order if t in defects]

                # 촬영불가지역
                notfind_found = []
                if not df_one.empty:
                    cols_all = [
                        "LEFT_UP_EXIST",
                        "RIGHT_UP_EXIST",
                        "LEFT_DOWN_EXIST",
                        "RIGHT_DOWN_EXIST",
                    ]
                    sub_flags = (
                        df_one.reindex(columns=cols_all, fill_value=None)
                        .astype(str)
                        .apply(lambda s: s.str.upper().fillna("N"))
                    )
                    for col, label in notfind_cols:
                        if col in sub_flags.columns and (sub_flags[col] == "N").any():
                            notfind_found.append(label)

                minors_found = [t for t in minors_order if t in defects]

                # --- 점검결과 텍스트 조합 ---
                if majors_found and notfind_found:
                    # 주요결함 + 촬영불가지역 같이 있을 때
                    if set(notfind_found) == {"좌하단 촬영불가지역", "우하단 촬영불가지역"}:
                        result_text = (
                            f"{'/'.join(majors_found)} 검출, 좌·우하단 촬영불가지역"
                        )
                    else:
                        result_text = (
                            f"{'/'.join(majors_found)} 검출, {'/'.join(notfind_found)}"
                        )

                elif majors_found:
                    # 주요결함만 있을 때
                    result_text = f"{'/'.join(majors_found)} 검출"

                elif notfind_found:
                    # 촬영불가지역만 있을 때
                    if set(notfind_found) == {"좌하단 촬영불가지역", "우하단 촬영불가지역"}:
                        result_text = "좌·우하단 촬영불가지역"
                    else:
                        result_text = " / ".join(notfind_found)

                else:
                    # 아무것도 없으면 정상
                    result_text = "정상상태"

                # --- 비고 텍스트 조합 ---
                remark_text = "/".join(minors_found) + " 검출" if minors_found else ""

                loc = f"{subproject_kr}@#{n}"

                rows.append(
                    [
                        "지지물",
                        loc,
                        "지지애자 손상여부",
                        "특이상태 점검@/AI점검",
                        result_text,
                        remark_text,
                    ]
                )

        return rows

    def collect_image_pairs(root_dir, project_id, sub_project_id, direction):
        """
        direction: "UP" 또는 "DOWN"
        return: [{"key": ST1_AJR_HCL_DOWN_0001, "ins_no": 1, "merge": "...jpg", "detect": "...png", "dir": "UP|DOWN"}, ...]
        """
        base_merge = os.path.join(
            root_dir,
            str(project_id),
            "01_PRE",
            "MERGE",
            f"{sub_project_id}_{direction}",
        )
        base_detect = os.path.join(
            root_dir,
            str(project_id),
            "02_AI",
            "DETECT",
            f"{sub_project_id}_{direction}",
        )

        results = []

        if not os.path.isdir(base_merge):
            return results  # 폴더 없으면 빈 리스트

        for f in sorted(os.listdir(base_merge)):
            if not f.lower().endswith(".jpg"):
                continue
            parts = f.split("_")
            ins_no = int(parts[-1].split(".")[0])  # 0001 → 1
            key = "_".join(
                parts[:-2] + [parts[-1].split(".")[0]]
            )  # ST1_AJR_HCL_DOWN_0001

            merge_path = os.path.join(base_merge, f)

            # 대응되는 결과 파일 이름
            detect_name = f"{'_'.join(parts[:-2])}_DETECT_{parts[-1].split('.')[0]}.png"
            detect_path = os.path.join(base_detect, detect_name)

            results.append(
                {
                    "key": key,
                    "ins_no": ins_no,
                    "merge": merge_path,
                    "detect": detect_path if os.path.exists(detect_path) else None,
                    "dir": direction,
                }
            )

        return results

    def extract_ins_no_from_row(row):
        if not row or len(row) < 2:
            return None

        loc_text = str(row[1])  # 위치 정보
        if "#" in loc_text:
            try:
                return int(loc_text.split("#")[-1].strip())
            except ValueError:
                return None
        return None

    def extract_dir_from_row(row):
        """
        row[1] 위치 컬럼에서 '(상행)' or '(하행)'을 파싱해 'UP'/'DOWN'으로 리턴
        예: '수성알파시티~정평(상행)\n#1' -> 'UP'
        """
        loc_text = str(row[1]) if row and len(row) >= 2 else ""
        # 괄호 안의 '상행'/'하행' 추출
        start = loc_text.find("(")
        end = loc_text.find(")", start + 1)
        if start != -1 and end != -1:
            dir_kr = loc_text[start + 1 : end].strip()
            if dir_kr == "상행":
                return "UP"
            if dir_kr == "하행":
                return "DOWN"
        return None

    def insert_value_with_splits(value: object):
        """'@'가 있으면 나눠서 사이에 개행 삽입"""
        text = str(value)
        parts = [p.strip() for p in text.split("@")]
        if len(parts) <= 1:
            insert_text(text)
            return
        # 첫 조각
        insert_text(parts[0])
        # 이후 조각들 앞에 개행 삽입
        for p in parts[1:]:
            hwp.HAction.Run("BreakPara")
            insert_text(p)

    def insert_colored_text(text):
        # 정상상태가 포함돼있으면, 그 글자만 검은색 나머지는 빨간색
        hwp.HAction.Run("ParagraphShapeAlignLeft")
        target = "정상상태"
        idx = text.find(target)

        if idx != -1:
            # 정상상태 앞에 오는 글자들 : 빨강
            before = text[:idx]
            if before:
                글자속성(10, 0, "red")
                insert_value_with_splits(before)
            
            #정상상태 : 검정
            글자속성(10, 0, "black")
            insert_value_with_splits(target)

            #정상상태 뒤의 글자 들 : 빨강
            after = text[idx + len(target):]
            if after:
                글자속성(10, 0, "red")
                insert_value_with_splits(after)
        
        else:
            #정상상태가 없으면 모두 빨강
            글자속성(10, 0, "red")
            insert_value_with_splits(text)
        
        #다음 셀에 영향 안가게 초기화
        글자속성(10, 0, "black")

    """
    ========================================
            페이지 별 작성 함수
    ========================================
    """

    def 표지(meta):

        # 한글 문서에서 쪽 번호 위치 설정
        hwp.HAction.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)

        # DrawPos 설정: 하단 중앙(3은 하단 중앙을 나타냄)
        hwp.HParameterSet.HPageNumPos.DrawPos = (
            5  # 3 = 하단 중앙, 1 = 상단 좌측, 2 = 상단 중앙, 4 = 하단 좌측, 등
        )

        # 설정 적용
        hwp.HAction.Execute("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
        hwp.HAction.Run("ParagraphShapeAlignCenter")

        글자속성(28, True)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        insert_text(meta["title_line"])

        hwp.HAction.Run("MoveDown")

        # 글자 크기를 15으로 설정
        글자속성(15, True)
        hwp.HAction.Run("ParagraphShapeAlignLeft")  # 왼쪽 정렬

        for i in range(0, 13):
            hwp.HAction.Run("BreakPara")

        # 설비명
        insert_text(f"{'설비명':<6} : {meta['facility_name']}")
        hwp.HAction.Run("BreakPara")

        # 관리번호
        insert_text(f"{'관리번호':<5} : {meta['managed_number']}")
        hwp.HAction.Run("BreakPara")

        # 검사일시
        today_date = date.today().strftime("%Y년 %m월 %d일")
        insert_text(f"{'작성일':<6} : {today_date}")
        hwp.HAction.Run("BreakPara")

        # 검사장소
        insert_text(f"{'검사장소':<5} : {meta['place']}")
        hwp.HAction.Run("BreakPara")

        # 점검자
        insert_text(f"{'점검자':<6} : {meta['inspector']}")
        hwp.HAction.Run("BreakPara")

        # 작성자
        insert_text(f"{'작성자':<6} : {meta['writer']}")
        hwp.HAction.Run("BreakPara")

        # 확인자
        insert_text(f"{'확인자':<6} : {meta['approver']}")
        hwp.HAction.Run("BreakPara")

        글자속성(14, True)
        hwp.HAction.Run("MoveDocEnd")
        hwp.Run("BreakPage")

    def 목차():
        hwp.HAction.Run("ParagraphShapeAlignCenter")

        글자속성(23, 1)
        insert_text("목       차")

        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("ParagraphShapeAlignLeft")  # 왼쪽 정렬
        Set = hwp.HParameterSet.HParaShape
        hwp.HAction.GetDefault("ParagraphShape", Set.HSet)
        tab_def = Set.TabDef
        tab_def.CreateItemArray("TabItem", 12)
        tab_def.TabItem.SetItem(0, 92000)
        tab_def.TabItem.SetItem(1, 3)
        tab_def.TabItem.SetItem(2, 0)
        hwp.HAction.Execute("ParagraphShape", Set.HSet)

        글자속성(15, 1)

        # 목차를 수동으로 삽입 (예시)
        hwp.HAction.Run("ParagraphShapeAlignJustify")
        insert_text("1. 상태평가 개요\t3")
        hwp.HAction.Run("BreakPara")

        글자속성(15, 0)
        insert_text("  1.1 상태평가 기준\t3")
        hwp.HAction.Run("BreakPara")

        insert_text("    가. 점검방법\t3")
        hwp.HAction.Run("BreakPara")

        insert_text("  1.2 결함의 정의\t4")
        hwp.HAction.Run("BreakPara")
        insert_text("    가. 주요 결함\t4")
        hwp.HAction.Run("BreakPara")
        insert_text("    나. 주의 결함\t5")
        hwp.HAction.Run("BreakPara")
        insert_text("    다. 촬영 불가 지역\t5")
        hwp.HAction.Run("BreakPara")

        hwp.HAction.Run("BreakPara")
        글자속성(15, 1)
        insert_text("2. 상태평가 결과\t6")
        hwp.HAction.Run("BreakPara")

        글자속성(15, 0)
        insert_text("  2.1 상태평가 요약\t6")
        hwp.HAction.Run("BreakPara")
        insert_text("  2.1 상태평가 상세내용\t7")
        hwp.HAction.Run("MoveDocEnd")
        hwp.Run("BreakPage")

    def 상태평가기준(cover_meta):
        hwp.HAction.Run("ParagraphShapeAlignLeft")  # 왼쪽 정렬
        글자속성(17, 1)
        insert_text("1. 상태평가 개요")

        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        여백생성(8.3)
        글자속성(15, 1)
        insert_text("1.1 상태평가 기준")
        hwp.HAction.Run("BreakPara")
        글자속성(12, 0)
        들여쓰기(0)
        들여쓰기(3000)
        insert_text(
            f"본 상태평가는 대구 지하철 {cover_meta['place']} 구간(상·하행 전 구간)을 대상으로 한다. 평가 대상은 전차선 지지물 중 지지애자로 한정하여 실시한다."
        )
        hwp.HAction.Run("BreakPara")
        글자속성(5, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        insert_text("가. 점검방법")

        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(12, 0)
        들여쓰기(4000)
        if cover_meta["camera_mode"] == "2CH":
            insert_text(
                "본 점검은 대상 구간의 전차선 지지애자를 비접촉 촬영·분석 방식으로 수행한다. 좌/우 애자 하단 촬영용 카메라 2대와 애자번호 식별 보조 카메라 1대를 병행 사용한다."
            )
        elif cover_meta["camera_mode"] == "4CH":
            insert_text(
                "본 점검은 대상 구간의 전차선 지지애자를 비접촉 촬영·분석 방식으로 수행한다. 좌/우 애자 상/하단 촬영용 카메라 4대와 애자번호 식별 보조 카메라 1대를 병행 사용한다."
            )
        else:
            insert_text(
                "본 점검은 대상 구간의 전차선 지지애자를 비접촉 촬영·분석 방식으로 수행한다. 좌/우 애자 촬영용 카메라 2대(필요시 4대)와 애자번호 식별 보조 카메라 1대를 병행 사용한다."
            )

        hwp.HAction.Run("BreakPara")

        hwp.HAction.Run("BreakPara")
        줄간격(180)
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [표1] 카메라 설치")

        hwp.HAction.Run("BreakPara")

        """
        ==============
             표1   
        ==============
        """
        # time.sleep(1)
        createTable(2, 3)

        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림1.png"), 50.68, 50.68)
        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림2.png"), 50.68, 50.68)
        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림3.png"), 50.68, 50.68)

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(12, 0)
        insert_text("애자 촬영용(좌) 카메라")

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(12, 0)
        insert_text("애자 촬영용(우) 카메라")

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(12, 0)
        insert_text("애자번호 식별용")

        hwp.HAction.Run("Close")
        hwp.HAction.Run("MoveDocEnd")

        """
        ==============
             그림1   
        ==============
        """
        createTable(1, 1)
        hwp.HAction.Run("BreakPara")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "그림1.png"), 104.40, 58.74)
        hwp.HAction.Run("MoveRight")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("Close")

        hwp.HAction.Run("BreakPara")
        글자속성(2, 0)
        hwp.HAction.Run("BreakPara")

        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [그림1] 촬영현장")
        글자속성(10, 0)
        hwp.HAction.Run("MoveDocEnd")
        hwp.Run("BreakPage")

    def 결함의정의():
        hwp.HAction.Run("ParagraphShapeAlignLeft")  # 왼쪽 정렬

        # --------------------------
        # 1.2 손상의 정의
        # --------------------------
        들여쓰기(0)
        여백생성(8.3)
        글자속성(15, 1)
        insert_text("1.2 결함의 정의")
        hwp.HAction.Run("BreakPara")
        글자속성(12, 0)
        들여쓰기(0)
        들여쓰기(3000)
        insert_text(
            "지지애자의 결함은 다음과 같이 주요결함과 주의결함으로 나누어 평가한다."
        )
        들여쓰기(0)
        들여쓰기(2000)
        hwp.HAction.Run("BreakPara")
        글자속성(5, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        insert_text("가. 주요 결함")

        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(12, 0)
        들여쓰기(2000)
        insert_text("지지애자를 교체해야하는 수준의 결함이다.")
        hwp.HAction.Run("BreakPara")

        # 표2
        hwp.HAction.Run("BreakPara")
        줄간격(180)
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [표2] 주요결함")
        글자속성(10, 0)
        createTable(3, 3)

        # 1열
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("결함 종류")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("내용")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("예시 이미지")

        # 2열
        hwp.HAction.Run("TableRightCell")
        insert_text("균열")
        hwp.HAction.Run("BreakPara")
        insert_text("(Crack)")
        hwp.HAction.Run("TableRightCell")
        insert_text("애자 표면에 생긴 미세한 금.")
        hwp.HAction.Run("BreakPara")
        insert_text("열팽창, 기계적 충격, 전기적 스트레스 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표2_그림1.png"), 35, 29.45)

        # 3열
        hwp.HAction.Run("TableRightCell")
        insert_text("파손")
        hwp.HAction.Run("BreakPara")
        insert_text("(Damage)")

        hwp.HAction.Run("TableRightCell")
        insert_text("애자 본체가 깨지거나 일부가 떨어져 나간 상태.")
        hwp.HAction.Run("BreakPara")
        insert_text("과전압, 외부충격, 균열의 심화 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표2_그림2.png"), 35, 29.45)
        resizeTable2(hwp, left_count=2, up_count=10)

        # 표 종료
        hwp.HAction.Run("Close")

        # 페이지 넘어감
        hwp.HAction.Run("MoveDocEnd")
        hwp.Run("BreakPage")

        들여쓰기(0)
        들여쓰기(1000)
        hwp.HAction.Run("ParagraphShapeAlignLeft")
        글자속성(13, 1)
        insert_text("나. 주의 결함")
        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(12, 0)
        들여쓰기(2000)
        insert_text("애자 교체가 불필요하나, 관찰이 필요한 결함이다.")
        hwp.HAction.Run("BreakPara")

        # 표3
        글자속성(2, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(1, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        insert_text(" [표3] 주의결함")
        글자속성(10, 0)
        createTable(4, 3)

        # 1열
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("결함 종류")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("내용")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("예시 이미지")

        # 2열
        hwp.HAction.Run("TableRightCell")
        insert_text("얼룩")
        hwp.HAction.Run("BreakPara")
        insert_text("(Stain)")

        hwp.HAction.Run("TableRightCell")
        insert_text("애자 표면에 먼지, 매연 등 오염 물질이 묻어 생기는 자국.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표3_그림1.png"), 35, 29.45)

        # 3열
        hwp.HAction.Run("TableRightCell")
        insert_text("아크혼(박리)")
        hwp.HAction.Run("BreakPara")
        insert_text("(ArchornPeeling)")

        hwp.HAction.Run("TableRightCell")
        insert_text(
            "아크 방전이 발생하면서 애자의 표피가 분리되어 떨어지기 직전이거나 떨어진 상태."
        )
        hwp.HAction.Run("BreakPara")
        insert_text("아크 방전 시 발생하는 충격 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표3_그림2.png"), 35, 29.45)

        # 4열
        hwp.HAction.Run("TableRightCell")
        insert_text("아크혼(그을음)")
        hwp.HAction.Run("BreakPara")
        insert_text("(ArchornScorching)")

        hwp.HAction.Run("TableRightCell")
        insert_text("아크 방전이 발생하면서 생긴 그을림 자국.")
        hwp.HAction.Run("BreakPara")
        insert_text(
            "아크 방전 시 발생하는 고온과 연기로 아크혼 주변에 검게 탄 흔적이 남음."
        )

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표3_그림3.png"), 35, 29.45)
        resizeTable2(hwp, left_count=2, up_count=10)
        # 표 종료
        hwp.HAction.Run("Close")
        hwp.HAction.Run("BreakPara")

        hwp.HAction.Run("ParagraphShapeAlignLeft")
        글자속성(13, 1)
        insert_text("다. 촬영 불가 지역")

        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(12, 0)
        들여쓰기(2000)
        insert_text(
            "촬영 불가 지역은 촬영 작업은 수행되었으나 구조물·설비 및 안전시설 등으로 애자 본체가 가려져 영상에서 식별되지 않은 경우를 의미한다. 본 항목은 결함 유형이 아닌 판독 상태이며, 재촬영 또는 보완조치의 대상이다."
        )
        hwp.HAction.Run("BreakPara")
        글자속성(2, 0)
        # 표4
        hwp.HAction.Run("BreakPara")
        줄간격(180)
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [표4] 촬영 불가 사유")
        글자속성(10, 0)
        createTable(3, 3)

        # 1열
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("사유")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("설명")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("예시 이미지")

        # 2열
        hwp.HAction.Run("TableRightCell")
        insert_text("가려짐")
        hwp.HAction.Run("TableRightCell")
        insert_text(
            "타 구조물, 부속장치 또는 작업 여건 등으로 대상 애자가 시야에서 가려져 검출이 곤란한 경우."
        )

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표4_그림1.png"), 35, 29.45)

        # 3열
        hwp.HAction.Run("TableRightCell")
        insert_text("완전 파손")

        hwp.HAction.Run("TableRightCell")
        insert_text(
            "대상 애자가 전면적으로 파손되어 형상이 식별되지 않아 검출이 곤란한 경우."
        )

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표4_그림2.png"), 35, 29.45)
        resizeTable2(hwp, left_count=2, up_count=10)

        # 표 종료
        hwp.HAction.Run("Close")

        # 페이지 넘어감
        hwp.Run("BreakPage")

    def 상태평가요약(project_id, sub_project_id):

        hwp.HAction.Run("ParagraphShapeAlignLeft")
        글자속성(17, 1)
        insert_text("2. 상태평가 결과")
        줄간격(180)
        hwp.HAction.Run("BreakPara")
        들여쓰기(3000)
        여백생성(8.3)
        글자속성(15, 1)
        insert_text("2.1 상태평가 요약")
        줄간격(180)
        글자속성(13, 1)
        글자속성(10, 0)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(13, 1)
        insert_text(" [표5] 상태평가 요약표")

        # 표5 (요약표)
        data_list = get_insulator_defect_summary(project_id, sub_project_id)

        # 헤더 정의
        headers = ["구분", "위치", "점검항목", "점검기준@/점검방법", "점검결과", "비고"]

        # 표 생성 (행 = 헤더 1줄 + 데이터 줄 수, 열 = 헤더 길이)
        createTable(1, len(headers))
        # --- 헤더 입력 ---
        for col, head in enumerate(headers):
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_value_with_splits(head)
            if col < len(headers) - 1:
                hwp.HAction.Run("TableRightCell")

        hwp.HAction.Run("TableAppendRow")
        hwp.Run("TableColBegin")

        # --- 데이터 입력 ---
        for row in data_list:
            for col, value in enumerate(row):
                text_value = str(value)
                if col == 4:
                    insert_colored_text(text_value)
                elif col ==5:
                    hwp.HAction.Run("ParagraphShapeAlignLeft")
                    insert_value_with_splits(text_value)
                else:
                    insert_value_with_splits(text_value)
                글자속성(10, 0, "black")
                if col < len(headers) - 1:
                    hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableAppendRow")
            hwp.Run("TableColBegin")

        hwp.HAction.Run("TableDeleteRow")
        resizeTable(hwp, left_count=4, down_count=2)
        # 마지막 빈 줄 삭제 (마지막 AppendRow 때문에)
        hwp.HAction.Run("Close")
        hwp.HAction.Run("MoveDocEnd")
        hwp.Run("BreakPage")

    def 상태평가결과(project_id, sub_project_id):

        들여쓰기(0)
        글자속성(15, 1)
        hwp.HAction.Run("ParagraphShapeAlignLeft")
        여백생성(8.3)
        글자속성(15, 1)
        들여쓰기(3000)
        insert_text("2.2 상태평가 결과")
        줄간격(180)
        글자속성(10, 0)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(13, 1)
        insert_text(" [표6] 상태평가 결과표")
        글자속성(10, 0)

        # 표6 (결과표)
        data_list = get_insulator_defect_details(project_id, sub_project_id)

        #  이미지 매칭 (상/하행 분리해서 키를 (dir, ins_no)로 구성)
        pairs_up = collect_image_pairs(root_dir, project_id, sub_project_id, "UP")
        pairs_down = collect_image_pairs(root_dir, project_id, sub_project_id, "DOWN")
        image_pairs = {(p["dir"], p["ins_no"]): p for p in (pairs_up + pairs_down)}

        # data_list에서 (dir, ins_no) 기준으로 유니크한 키 목록 만들기
        unique_keys = []
        seen = set()

        for row in data_list:
            ins_no = extract_ins_no_from_row(row)
            dir_code = extract_dir_from_row(row)
            if ins_no is None or dir_code is None:
                continue
            key = (dir_code, ins_no)
            if key not in seen:
                seen.add(key)
                unique_keys.append(key)

        # 각 (dir, ins_no) 별로 표 생성
        for i, (dir_code, ins_no) in enumerate(unique_keys, start=1):
            # 이 키에 해당하는 데이터 행 찾기 (첫 행 사용)
            target_row = None
            for r in data_list:
                if (
                    extract_ins_no_from_row(r) == ins_no
                    and extract_dir_from_row(r) == dir_code
                ):
                    target_row = r
                    break
            if target_row is None:
                continue

            createTable(4, 6)
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableMergeCell")
            hwp.HAction.Run("TableColBegin")
            pair = image_pairs.get((dir_code, ins_no))
            if pair and pair.get("merge"):
                원본비율유지_이미지_삽입(pair["merge"])
            else:
                hwp.HAction.Run("ParagraphShapeAlignCenter")
                원본비율유지_이미지_삽입(os.path.join(os.getcwd(), "noImage.png"))

            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableMergeCell")
            if pair and pair.get("detect"):
                원본비율유지_이미지_삽입(pair["detect"])
            else:
                hwp.HAction.Run("ParagraphShapeAlignCenter")
                원본비율유지_이미지_삽입(os.path.join(os.getcwd(), "noImage.png"))

            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableMergeCell")
            hwp.HAction.Run("TableColBegin")
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_text("원본사진")

            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableMergeCell")
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_text("AI검출결과")

            hwp.HAction.Run("TableRightCell")

            # 헤더 정의
            headers = [
                "구분",
                "위치",
                "점검항목",
                "점검기준@/점검방법",
                "점검결과",
                "비고",
            ]

            # --- 헤더 입력 ---
            for col, head in enumerate(headers):
                insert_value_with_splits(str(head))
                if col < len(headers):
                    hwp.HAction.Run("TableRightCell")

            # --- 데이터 입력: target_row 1건만 기입 ---
            for col, value in enumerate(target_row):
                text_value = str(value)

                if col == 4:  # 점검결과 컬럼
                    insert_colored_text(text_value)
                else:
                    insert_value_with_splits(text_value)
                if col < len(headers):
                    hwp.HAction.Run("TableRightCell")

            resizeTable(hwp, left_count=4, down_count=2)
            hwp.HAction.Run("Close")
            hwp.HAction.Run("MoveDocEnd")
            # 3번째마다 page break 실행 (마지막 루프 제외)
            if i % 3 == 0 and i != len(unique_keys):
                hwp.HAction.Run("BreakPage")

    """ 작성 시작 """
    try:
        hwp = init_hwp()
        hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)

        # HSecDef 설정 변경
        hsecdef = hwp.HParameterSet.HSecDef
        hsecdef.PageDef.LeftMargin = 20.0 * 283.465  # 왼쪽 여백
        hsecdef.PageDef.RightMargin = 20.0 * 283.465  # 오른쪽 여백
        hsecdef.PageDef.TopMargin = 10 * 283.465
        hsecdef.HSet.SetItem("ApplyClass", 24)  # 적용 클래스 설정 (24: 현재 섹션)
        hsecdef.HSet.SetItem("ApplyTo", 3)  # 적용 대상 설정 (3: 전체 문서)

        # PageSetup 액션 실행
        hwp.HAction.Execute("PageSetup", hsecdef.HSet)

        cover_meta, ins_cnt_map, defect_rows = fetch_data_from_db(
            project_id, sub_project_id
        )

        표지(cover_meta)
        목차()
        상태평가기준(cover_meta)
        결함의정의()
        상태평가요약(project_id, sub_project_id)
        상태평가결과(project_id, sub_project_id)

        # ===== 파일명 생성 =====
        line_name = cover_meta["line_name"].replace(" ", "")  # "3호선"
        raw_section_core = cover_meta.get("section_core") or ""  # ~ 대신 _ 로 통일
        section_core_clean = raw_section_core.replace("~", "_").replace(" ", "")  #

        filename_base = (
            f"대구지하철_{line_name}_{section_core_clean}_상세_상태평가보고서"
        )

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        # set save filename
        # 최종 출력 폴더 구성
        base_dir_han = os.path.join(
            root_dir, str(project_id), "04_REPORT", "REPORT_HAN"
        )
        base_dir_pdf = os.path.join(
            root_dir, str(project_id), "04_REPORT", "REPORT_PDF"
        )
        os.makedirs(base_dir_han, exist_ok=True)
        os.makedirs(base_dir_pdf, exist_ok=True)

        # HWP 저장
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(
            base_dir_han, f"{sub_project_id}_DETAIL.hwp"
        )
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        # PDF 저장
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(
            base_dir_pdf, f"{sub_project_id}_DETAIL.pdf"
        )
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        # 저장 파일 경로
        hwp_path = os.path.join(base_dir_han, f"{sub_project_id}_DETAIL.hwp")
        pdf_path = os.path.join(base_dir_pdf, f"{sub_project_id}_DETAIL.pdf")

        # 콘솔 출력
        print(f"[INFO] HWP 저장 경로: {hwp_path}")
        print(f"[INFO] PDF 저장 경로: {pdf_path}")

        # hwp.Quit()  # 한글 종료
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if hwp:
            hwp.Clear(option=1)  # 오류발생시 한글 버림
            hwp.Quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="대구교통공사 상세 상태평가 보고서 생성기"
    )
    parser.add_argument("--root-dir", required=True, help="루트 디렉토리 경로")
    parser.add_argument("--project-id", required=True, type=int, help="프로젝트 ID")
    parser.add_argument("--sub-project-id", required=True, help="서브 프로젝트 ID")
    args = parser.parse_args()

    root_dir = args.root_dir
    project_id = args.project_id
    sub_project_id = args.sub_project_id

    makeHwp(root_dir, project_id, sub_project_id)

    print("프로그램 종료")
