import json
import os
import time
from collections import defaultdict
from datetime import date, datetime
import pandas as pd
import win32com.client as win32
from PIL import Image
import mysql.connector
import argparse


# DB에서 데이터 추출
def fetch_data_from_db(project_id):
    connection = mysql.connector.connect(
        host="localhost",
        port=10645,
        user="deepinspector",
        password="xoaud17!@",
        database="db_deepinspector"
    )
    cursor = connection.cursor(dictionary=True)

    # 1) 프로젝트 설정값 가져오기
    cursor.execute("""
        SELECT SETTING_VALUE, PROJECT_ID_DESC_EN
        FROM PROJECT
        WHERE PROJECT_ID = %s;
    """, (project_id,))
    row = cursor.fetchone()
    project_id_desc_en = row["PROJECT_ID_DESC_EN"]
    setting_value = json.loads(row["SETTING_VALUE"]) if row and row.get("SETTING_VALUE") else {}

    # 2) 결함 정보 가져오기
    query_defects = """
        SELECT 
            DDI.IMAGE_NUM,
            DDID.TYPE AS DEFECT_TYPE
        FROM DEFECT_DETECT_IMAGE_DETAIL AS DDID
        JOIN DEFECT_DETECT_IMAGE AS DDI
            ON DDID.DEFECT_DETECT_IMAGE_ID = DDI.DEFECT_DETECT_IMAGE_ID
        JOIN DEFECT_DETECT AS DD
            ON DDI.DEFECT_DETECT_ID = DD.DEFECT_DETECT_ID
        WHERE DD.PROJECT_ID = %s
        ORDER BY DDI.IMAGE_NUM ASC;
    """
    cursor.execute(query_defects, (project_id,))
    defect_rows = cursor.fetchall()

    # 3) 역 정보 가져오기
    cursor.execute("SELECT INITIAL, NAME, LINE, STATION_ORDER FROM SUBWAY_STATIONS;")
    station_rows = cursor.fetchall()
    STATION_MAP = {r["INITIAL"]: r["NAME"] for r in station_rows}
    STATION_ORDER = {(row["LINE"], row["INITIAL"]): int(row["STATION_ORDER"]) for row in station_rows}

    cursor.close()
    connection.close()

    # === 한글명 매핑 변환 ===
    DEFECT_TYPE_MAP = {
        "crack": "균열",
        "damaged": "파손",
        "archorn_peeling": "아크혼_박리",
        "archorn_soot": "아크혼_그을음",
        "stain": "얼룩",
    }
    LINE_MAP = {"ST1": "1호선", "ST2": "2호선", "ST3": "3호선"}

    for r in defect_rows:
        # 결함명 번역
        r["DEFECT_TYPE_KR"] = DEFECT_TYPE_MAP.get(r["DEFECT_TYPE"], r["DEFECT_TYPE"])
        # IMAGE_NUM 파싱
        parts = (r["IMAGE_NUM"] or "").split("_")
        if len(parts) >= 5:
            line, from_init, to_init, direction, ins_no = parts
            r["LINE_KR"] = LINE_MAP.get(line, line)
            r["FROM_STATION"] = STATION_MAP.get(from_init, from_init)
            r["TO_STATION"] = STATION_MAP.get(to_init, to_init)
            r["FROM_ORDER"] = STATION_ORDER.get((line, from_init), 10**9)
            r["TO_ORDER"]   = STATION_ORDER.get((line, to_init),   10**9)
            r["DIRECTION"] = "상행" if direction.upper() == "UP" else ("하행" if direction.upper() == "DOWN" else direction)
            r["INSULATOR_NO"] = str(int(ins_no))  # "0001" -> "1"
            r["SUB_PROJECT_ID"] = "_".join(parts[:-1])
            r["SUBPROJECT_KR"] = f"{r['FROM_STATION']}~{r['TO_STATION']}({r['DIRECTION']})"

    # 애자 개수 맵
    subproject_insulator_counts = {
        k: v["INSULATOR_COUNT"]
        for k, v in setting_value.items()
        if isinstance(v, dict) and "INSULATOR_COUNT" in v
    }

    # ====== 표지 메타 구성 ======
    # 1) 호선명(line_name)
    line_name = None
    if defect_rows and "LINE_KR" in defect_rows[0]:
        # 최빈값
        lines = [r.get("LINE_KR") for r in defect_rows if r.get("LINE_KR")]
        if lines:
            # 최빈값 추출
            line_counts = {}
            for ln in lines:
                line_counts[ln] = line_counts.get(ln, 0) + 1
            line_name = max(line_counts.items(), key=lambda x: x[1])[0]
    if not line_name:
        # setting key에서 STx 하나 골라 매핑
        codes = [k.split("_")[0] for k in setting_value.keys() if k.startswith("ST")]
        code = codes[0] if codes else "ST2"
        line_name = LINE_MAP.get(code, code)

    # 2) 전체 구간(section_core): settings.subwaySection 우선
    if setting_value.get("subwaySection"):
        section_core = setting_value["subwaySection"].strip()
    else:
        # defect_rows에서 SUBPROJECT_KR 정렬 후 [첫 시작역 ~ 마지막 도착역]
        if defect_rows:
            df = pd.DataFrame(defect_rows)
            if not df.empty and "SUBPROJECT_KR" in df.columns:
                df["START_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].min(axis=1)
                df["END_ORDER"]   = df[["FROM_ORDER", "TO_ORDER"]].max(axis=1)
                order = (
                    df[["SUB_PROJECT_ID", "SUBPROJECT_KR", "LINE_KR", "START_ORDER", "END_ORDER"]]
                    .drop_duplicates()
                    .sort_values(["LINE_KR", "START_ORDER", "END_ORDER"])
                )
                if not order.empty:
                    first = order.iloc[0]["SUBPROJECT_KR"]  # "만평~팔달시장(상행)"
                    last  = order.iloc[-1]["SUBPROJECT_KR"]
                    start = first.split("~")[0].strip()
                    end   = last.split("~")[1].split("(")[0].strip()
                    section_core = f"{start} ~ {end}"
                else:
                    section_core = "구간미상"
            else:
                section_core = "구간미상"
        else:
            section_core = "구간미상"

    # 3) 표지 필드들
    facility_name  = setting_value.get("facilityName", "")
    managed_number = setting_value.get("managedNumber", "")
    inspector_raw  = setting_value.get("inspector", "")
    inspector      = ", ".join([x.strip() for x in inspector_raw.split(",") if x.strip()])
    approver       = setting_value.get("approver", "")
    writer         = "Deep Inspector(AI 안전점검 프로그램)"
    written_date   = date.today().strftime("%Y년 %m월 %d일")

    # 4) 제목 라인 (대괄호 표기)
    # 예: 대구 지하철 [2]호선 [수성알파시티] ~ [정평] 상태평가 보고서
    try:
        # section_core: "수성알파시티 ~ 정평"
        left, right = [s.strip() for s in section_core.split("~", 1)]
    except Exception:
        left, right = section_core, ""
    title_line = f"대구 지하철 {line_name.replace('호선','')}호선 {left} ~ {right} 상태평가 보고서"

    cover_meta = {
        "title_line": title_line,
        "facility_name": facility_name,
        "managed_number": managed_number,
        "written_date": written_date,
        "place": f"{line_name} {section_core}",   # 검사장소
        "inspector": inspector,
        "writer": writer,
        "approver": approver,
        "line_name": line_name,
        "section_core": section_core,
        "project_id_desc_en": project_id_desc_en,
    }

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

def makeHwp(root_dir, project_id):
    
    """
    ================================
             기본 동작 정의
    ================================
    """
    def 글자속성(font_size=11, bold=False):
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

    def 클립보드로_이미지_삽입(filepath, width, height ):
        """
        한/글 API의 InsertPicture 메서드는
        셀의 크기를 변경하지 않는 반면(이미지가 찌그러짐)
        클립보드를 통해 이미지를 삽입하면
        이미지의 종횡비에 맞춰
        셀의 높이가 자동으로 조절됨.
        """
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
        hwp.HParameterSet.HTableCreation.WidthType = max_width/cols  # 테이블 폭을 열 수에 맞게 균등 분할
        hwp.HParameterSet.HTableCreation.HeightType = 0
        hwp.HParameterSet.HTableCreation.WidthValue = 0.0
        hwp.HParameterSet.HTableCreation.HeightValue = 0.0
        # 테이블 폭 설정
        hwp.HParameterSet.HTableCreation.TableProperties.Width = max_width
        # 테이블 생성 액션 실행
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        hwp.HAction.Run("TableColWidthEven")

    def resizeTable(hwp, left_count=0, down_count=0):
        """
        표 셀 블록 확장 후 크기 조정
        - extend는 반드시 2번 실행
        - left_count: TableResizeExLeft 실행 횟수
        - down_count: TableResizeExDown 실행 횟수
        """
        # 블록 확장 (2번 고정)
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableCellBlockExtend")

        # 왼쪽으로 줄이기
        for _ in range(left_count):
            hwp.HAction.Run("TableResizeExLeft")

        # 아래로 줄이기
        for _ in range(down_count):
            hwp.HAction.Run("TableResizeExDown")
    
    """
    ========================================
            대구교통공사 전용 동작 정의
    ========================================
    """
    def get_insulator_defect_summary(project_id):
        # DB에서 데이터 가져오기
        _, insulator_count_map, defect_rows = fetch_data_from_db(project_id)

        # DataFrame 변환
        df = pd.DataFrame(defect_rows)

        df["START_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].min(axis=1)
        df["END_ORDER"]   = df[["FROM_ORDER", "TO_ORDER"]].max(axis=1)

        # 구간별( SUB_PROJECT_ID ) 대표 정렬키 뽑기
        order_df = (
            df[["SUB_PROJECT_ID", "LINE_KR", "START_ORDER", "END_ORDER"]]
            .drop_duplicates()
            .sort_values(["LINE_KR", "START_ORDER", "END_ORDER"])
        )
        # 정렬 인덱스 맵: SUB_PROJECT_ID 순서
        order_idx = {sp: i for i, sp in enumerate(order_df["SUB_PROJECT_ID"].tolist())}

        # 주요/주의 결함 분류
        major_defects = {"균열", "파손"}  # 교체 필요
        caution_defects = {"얼룩", "아크혼_박리", "아크혼_그을음"}  # 주의 필요

        # 결과 저장용 리스트
        report_rows = []

        # 전체 서브프로젝트이름, 애자 갯수로 반복문
        for sub_project, total_count in sorted(
        insulator_count_map.items(),
        key=lambda kv: order_idx.get(kv[0], 10**9)):

            # 현재 루프에 해당하는 구간만 필터링
            df_sub = df[df["SUB_PROJECT_ID"] == sub_project]

            # 구간 이름 한글로 가져오기
            if not df_sub.empty:
                subproject_kr = df_sub.iloc[0]["SUBPROJECT_KR"]
            else:
                # 결함이 하나도 없는 경우: IMAGE_NUM 기반으로 기본 텍스트 생성
                parts = sub_project.split("_")
                if len(parts) >= 4:
                    from_init, to_init, direction = parts[1:4]
                    direction_map = {"UP": "상행", "DOWN": "하행"}
                    from_station = from_init 
                    to_station = to_init
                    direction_kr = direction_map.get(direction.upper(), direction)
                    subproject_kr = f"{from_station}~{to_station}({direction_kr})"
                else:
                    subproject_kr = sub_project

            replace_list = []
            caution_list = []

            # 결함 집계
            for ins_no in df_sub["INSULATOR_NO"].unique():
                defects = df_sub[df_sub["INSULATOR_NO"] == ins_no]["DEFECT_TYPE_KR"].unique()
                if len(defects) == 0:
                    continue

                defects_str = ", ".join(defects)
                ins_str = f"#{ins_no}({defects_str})"

                if any(d in major_defects for d in defects):
                    replace_list.append(ins_str)
                elif any(d in caution_defects for d in defects):
                    caution_list.append(ins_str)

            # 점검결과 / 비고 텍스트
            replace_text = "교체 필요 애자: " + ", ".join(replace_list) if replace_list else "정상상태"
            caution_text = "주의 필요 애자: " + ", ".join(caution_list) if caution_list else "정상상태"

            # 요약 행 추가
            report_rows.append([
                "지지물",
                f"{subproject_kr} #1~#{total_count}",
                "지지애자 손상여부",
                "특이상태 점검/AI점검",
                replace_text,
                caution_text
            ])

        # DataFrame으로 변환
        # "구분", "위치", "점검항목", "점검기준/점검방법", "점검결과", "비고"
        return report_rows

    def get_insulator_defect_details(project_id):
        _, insulator_count_map, defect_rows = fetch_data_from_db(project_id)
        df = pd.DataFrame(defect_rows)

        # STATION_ORDER 기반(라인/시작/끝) + 방향(UP→DOWN)까지 포함해 정렬키 생성
        df["START_ORDER"] = df[["FROM_ORDER", "TO_ORDER"]].min(axis=1)
        df["END_ORDER"]   = df[["FROM_ORDER", "TO_ORDER"]].max(axis=1)

        # SUB_PROJECT_ID의 마지막 토큰으로 방향 추출 (예: ST3_MPY_PSS_UP -> UP)
        df["DIR_CODE"]  = df["SUB_PROJECT_ID"].str.split("_").str[-1].str.upper()
        df["DIR_ORDER"] = df["DIR_CODE"].map({"UP": 0, "DOWN": 1}).fillna(2).astype(int)

        order = (
            df[["SUB_PROJECT_ID", "SUBPROJECT_KR", "LINE_KR", "START_ORDER", "END_ORDER", "DIR_ORDER"]]
            .drop_duplicates()
            .sort_values(["LINE_KR", "START_ORDER", "END_ORDER", "DIR_ORDER"])
        )

        major   = {"균열", "파손"}
        archorns = {"아크혼(박리)", "아크혼(그을음)"}
        stain   = "얼룩"

        rows = []

        for sub_project, subproject_kr in zip(order["SUB_PROJECT_ID"], order["SUBPROJECT_KR"]):
            total = insulator_count_map.get(sub_project)
            if not total:
                continue

            df_sub = df[df["SUB_PROJECT_ID"] == sub_project]

            for n in range(1, int(total) + 1):
                df_one = df_sub[df_sub["INSULATOR_NO"] == str(n)]
                defects = set(df_one["DEFECT_TYPE_KR"].tolist())

                major_found = sorted(list(defects & major))
                if len(major_found) >= 2:
                    result_text = "균열/파손 검출"
                elif len(major_found) == 1:
                    result_text = f"{major_found[0]} 검출"
                else:
                    result_text = "정상상태"

                remarks = []
                if defects & archorns:
                    remarks.append("아크혼 검출")
                if stain in defects:
                    remarks.append("얼룩 검출")
                remark_text = ", ".join(remarks)

                loc = f"{subproject_kr}\n#{n}"

                rows.append([
                    "지지물",
                    loc,
                    "지지애자 손상여부",
                    "특이상태 점검/AI점검",
                    result_text,
                    remark_text
                ])

        return rows

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
        hwp.HAction.Run("ParagraphShapeAlignLeft") # 왼쪽 정렬

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
        hwp.HAction.Run("ParagraphShapeAlignLeft") # 왼쪽 정렬
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
        insert_text("    가. 주의 결함\t4")
        hwp.HAction.Run("BreakPara")

        hwp.HAction.Run("BreakPara")
        글자속성(15, 1)
        insert_text("2. 상태평가 결과\t5")
        hwp.HAction.Run("BreakPara")

        글자속성(15, 0)
        insert_text("  2.1 상태평가 요약\t5")
        hwp.HAction.Run("BreakPara")
        insert_text("  2.1 상태평가 상세내용\t6")
        hwp.HAction.Run("BreakPara")
        
        hwp.Run("BreakPage")

    def 상태평가범위():
        hwp.HAction.Run("ParagraphShapeAlignLeft") # 왼쪽 정렬
        글자속성(17, 1)
        insert_text("1. 상태평가 개요")

        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        여백생성(8.3)
        글자속성(15, 1)
        insert_text("1.1 상태평가 기준")
    
        # hwp.Run("BreakPa")  # 페이지 나누기 삽입
        hwp.HAction.Run("BreakPara")
        들여쓰기(0)
        들여쓰기(1000)
        text = "본 상태평가는 대구 지하철 [{{]2]호선 [수성알파시티역] ~ [정평역] 구간(상·하행 전 구간)을 대상으로 한다. 평가 대상은 전차선 지지물 중 지지애자로 한정하여 실시한다."
        
        hwp.HAction.Run("BreakPara")
        insert_text("가. 점검방법")

        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(13, 0)
        들여쓰기(2000)
        insert_text("점검 방법 및 촬영 방법은 다음과 같다.")
        hwp.HAction.Run("BreakPara")

        hwp.HAction.Run("BreakPara")
        줄간격(180)
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [표1] 카메라 설치")
        글자속성(10, 0)

        hwp.HAction.Run("BreakPara")

        """
        ==============
             표1   
        ==============
        """
        #time.sleep(1)
        createTable(2, 3)

        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림1.png"),50.68,50.68)
        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림2.png"),50.68,50.68)
        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입(os.path.join(os.getcwd(), "표1_그림3.png"),50.68,50.68)

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(10, 0)
        insert_text("애자 촬영용(좌) 카메라")

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(10, 0)
        insert_text("애자 촬영용(우) 카메라")

        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(10, 0)
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
        클립보드로_이미지_삽입(  os.path.join(os.getcwd(),"그림1.png"),104.40,58.74)
        hwp.HAction.Run("MoveRight")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("Close")
        hwp.HAction.Run("MoveDocEnd")

        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text(" [그림1] 촬영현장")
        글자속성(10, 0)
        hwp.Run("BreakPage")
   
    def 결함의정의():
        hwp.HAction.Run("ParagraphShapeAlignLeft") # 왼쪽 정렬

        # --------------------------
        # 1.2 손상의 정의
        # --------------------------
        들여쓰기(0)
        여백생성(8.3)
        글자속성(15, 1)
        insert_text("1.2 결함의 정의")
        hwp.HAction.Run("BreakPara")
        글자속성(10, 0)
        들여쓰기(3000)
        insert_text("지지애자의 결함은 다음과 같이 주요결함과 주의결함으로 나누어 평가한다.")
        hwp.HAction.Run("BreakPara")

        줄간격(180)
        들여쓰기(0)
        들여쓰기(1000)
        글자속성(12, 1)
        insert_text("가. 주요 결함")
        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(10, 0)
        들여쓰기(2000)
        insert_text("지지애자를 교체해야하는 수준의 결함이다.")
        hwp.HAction.Run("BreakPara")
        
        #표2
        글자속성(2, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        insert_text(" [표2] 주요결함")
        글자속성(10, 0)
        createTable(3, 3)

        #1열
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("결함 종류")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("내용")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("예시 이미지")
        
        #2열
        hwp.HAction.Run("TableRightCell")
        insert_text("균열")
        hwp.HAction.Run("BreakPara")
        insert_text("(Crack)")
        hwp.HAction.Run("TableRightCell")
        insert_text("애자 표면에 생긴 미세한 금.")
        hwp.HAction.Run("BreakPara")
        insert_text("열팽창, 기계적 충격, 전기적 스트레스 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입( os.path.join(os.getcwd(),"표2_그림1.png"), 35 , 29.45 )

        #3열
        hwp.HAction.Run("TableRightCell")
        insert_text("파손")
        hwp.HAction.Run("BreakPara")
        insert_text("(Damaged)")

        hwp.HAction.Run("TableRightCell")
        insert_text("애자 본체가 깨지거나 일부가 떨어져 나간 상태.")
        hwp.HAction.Run("BreakPara")
        insert_text("과전압, 외부충격, 균열의 심화 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입( os.path.join(os.getcwd(),"표2_그림2.png"), 35 , 29.45 )

        # 표 종료
        resizeTable(hwp, left_count=15, down_count=5)
        hwp.HAction.Run("Close")
        글자속성(4, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(10, 0)
        들여쓰기(0)
        들여쓰기(1000)
        hwp.HAction.Run("ParagraphShapeAlignLeft")
        글자속성(12, 1)
        insert_text("나. 주의 결함")
        hwp.HAction.Run("BreakPara")
        여백생성(8.3)
        글자속성(10, 0)
        들여쓰기(2000)
        insert_text("애자 교체가 불필요하나, 관찰이 필요한 결함이다.")
        hwp.HAction.Run("BreakPara")
        
        #표3
        글자속성(13, 1)
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        글자속성(1, 0)
        hwp.HAction.Run("BreakPara")
        글자속성(13, 1)
        insert_text(" [표3] 주의결함")
        글자속성(10, 0)
        createTable(4, 3)
        
        #1열
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("결함 종류")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("내용")
        hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("ParagraphShapeAlignCenter")
        insert_text("예시 이미지")
        
        #2열
        hwp.HAction.Run("TableRightCell")
        insert_text("얼룩")
        hwp.HAction.Run("BreakPara")
        insert_text("(stain)")

        hwp.HAction.Run("TableRightCell")
        insert_text("애자 표면에 먼지, 매연 등 오염 물질이 묻어 생기는 자국.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입( os.path.join(os.getcwd(),"표3_그림1.png"), 35 , 29.45 )
        
        #3열
        hwp.HAction.Run("TableRightCell")
        insert_text("아크혼(박리)")
        hwp.HAction.Run("BreakPara")
        insert_text("(archorn_peeling)")

        hwp.HAction.Run("TableRightCell")
        insert_text("아크 방전이 발생하면서 애자의 표피가 분리되어 떨어지기 직전이거나 떨어진 상태.")
        hwp.HAction.Run("BreakPara")
        insert_text("아크 방전 시 발생하는 충격 등으로 발생할 수 있음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입( os.path.join(os.getcwd(),"표3_그림2.png"), 35 , 29.45 )

        #4열
        hwp.HAction.Run("TableRightCell")
        insert_text("아크혼(그을음)")
        hwp.HAction.Run("BreakPara")
        insert_text("(archorn_soot)")

        hwp.HAction.Run("TableRightCell")
        insert_text("아크 방전이 발생하면서 생긴 그을림 자국.")
        hwp.HAction.Run("BreakPara")
        insert_text("아크 방전 시 발생하는 고온과 연기로 아크혼 주변에 검게 탄 흔적이 남음.")

        hwp.HAction.Run("TableRightCell")
        클립보드로_이미지_삽입( os.path.join(os.getcwd(),"표3_그림3.png"), 35 , 29.45 )

        # 표 종료
        resizeTable(hwp, left_count=15, down_count=5)
        hwp.HAction.Run("Close")
        hwp.HAction.Run("BreakPara")
        hwp.Run("BreakPage")

    def 상태평가요약(project_id):

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
        insert_text(" [표4] 상태평가 요약표")

        # 표4 (요약표)
        data_list = get_insulator_defect_summary(project_id)

        # 헤더 정의
        headers = ["구분", "위치", "점검항목", "점검기준/점검방법", "점검결과", "비고"]

        # 표 생성 (행 = 헤더 1줄 + 데이터 줄 수, 열 = 헤더 길이)
        createTable(1, len(headers))
        # --- 헤더 입력 ---
        for col, head in enumerate(headers):
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_text(head)
            if col < len(headers) - 1:
                hwp.HAction.Run("TableRightCell")

        hwp.HAction.Run("TableAppendRow")
        hwp.Run("TableColBegin")

        # --- 데이터 입력 ---
        for row in data_list:
            for col, value in enumerate(row):
                insert_text(str(value))
                if col < len(headers) - 1:
                    hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableAppendRow")
            hwp.Run("TableColBegin")

        resizeTable(hwp, left_count=15, down_count=2)
        # 마지막 빈 줄 삭제 (마지막 AppendRow 때문에)
        hwp.HAction.Run("TableDeleteRow")
        hwp.HAction.Run("Close")

        hwp.Run("BreakPage")

    def 상태평가결과(project_id):

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
        insert_text(" [표5] 상태평가 결과표")

        # 표5 (결과표)
        data_list = get_insulator_defect_details(project_id)

        # 헤더 정의
        headers = ["구분", "위치", "점검항목", "점검기준/점검방법", "점검결과", "비고"]

        # 표 생성 (행 = 헤더 1줄 + 데이터 줄 수, 열 = 헤더 길이)
        createTable(1, len(headers))
        # --- 헤더 입력 ---
        for col, head in enumerate(headers):
            hwp.HAction.Run("ParagraphShapeAlignCenter")
            insert_text(head)
            if col < len(headers) - 1:
                hwp.HAction.Run("TableRightCell")

        hwp.HAction.Run("TableAppendRow")
        hwp.Run("TableColBegin")

        # --- 데이터 입력 ---
        for row in data_list:
            for col, value in enumerate(row):
                insert_text(str(value))
                if col < len(headers) - 1:
                    hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableAppendRow")
            hwp.Run("TableColBegin")

        resizeTable(hwp, left_count=15, down_count=2)
        # 마지막 빈 줄 삭제 (마지막 AppendRow 때문에)
        hwp.HAction.Run("TableDeleteRow")
        hwp.HAction.Run("Close")

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

        cover_meta, ins_cnt_map, defect_rows = fetch_data_from_db(project_id)

        표지(cover_meta)
        목차()
        상태평가범위()
        결함의정의()
        상태평가요약(project_id)
        상태평가결과(project_id)

        # ===== 파일명 생성 =====
        line_name   = cover_meta["line_name"].replace(" ", "")   # "3호선"
        raw_section_core = cover_meta.get("section_core") or ""        # ~ 대신 _ 로 통일
        section_core_clean = raw_section_core.replace("~", "_").replace(" ", "")  # 

        ## filename_base = f"대구지하철_{line_name}_{section_core_clean}_전체_상태평가보고서"
        filename_base = cover_meta["project_id_desc_en"]


        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        # set save filename
        # 최종 출력 폴더 구성
        base_dir_han = os.path.join(root_dir, str(project_id), "04_REPORT", "REPORT_HAN")
        base_dir_pdf = os.path.join(root_dir, str(project_id), "04_REPORT", "REPORT_PDF")
        os.makedirs(base_dir_han, exist_ok=True)
        os.makedirs(base_dir_pdf, exist_ok=True)

        # HWP 저장
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(base_dir_han, f"{filename_base}_TOTAL.hwp")
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        # PDF 저장
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(base_dir_pdf, f"{filename_base}_TOTAL.pdf")
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        # 저장 파일 경로
        hwp_path = os.path.join(base_dir_han, f"{filename_base}_TOTAL.hwp")
        pdf_path = os.path.join(base_dir_pdf, f"{filename_base}_TOTAL.pdf")

        # 콘솔 출력
        print(f"[INFO] HWP 저장 경로: {hwp_path}")
        print(f"[INFO] PDF 저장 경로: {pdf_path}")

        # hwp.Quit()  # 한글 종료
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        time.sleep(3)
        if hwp:
            hwp.Clear(option=1)  # 오류발생시 한글 버림
            hwp.Quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="대구교통공사 전체 상태평가 보고서 생성기")
    parser.add_argument("--root-dir", required=True, help="루트 디렉토리 경로")
    parser.add_argument("--project-id", required=True, type=int, help="프로젝트 ID")
    args = parser.parse_args()

    root_dir = args.root_dir
    project_id = args.project_id

    makeHwp(root_dir, project_id)

    print("프로그램 종료")
