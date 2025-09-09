# dtro_dtl_statevl.py
import argparse
import json
import os
import sys
import traceback
from datetime import date
import mysql.connector
import pandas as pd
import win32com.client as win32

# --- 콘솔 한글 깨짐 방지 ---
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

LINE_MAP = {"ST1": "1호선", "ST2": "2호선", "ST3": "3호선"}
DEFECT_TYPE_MAP = {
    "crack": "균열",
    "damaged": "파손",
    "archorn_peeling": "아크혼(박리)",
    "archorn_soot": "아크혼(그을음)",
    "stain": "얼룩",
}

def fetch_data_from_db(project_id: int):
    print("[INFO] DB 연결 및 데이터 로딩 시작...")
    connection = mysql.connector.connect(
        host="localhost", port=10645,
        user="deepinspector", password="xoaud17!@",
        database="db_deepinspector"
    )
    cursor = connection.cursor(dictionary=True)

    # 1) 프로젝트 설정
    cursor.execute("""
        SELECT SETTING_VALUE
        FROM PROJECT
        WHERE PROJECT_ID = %s;
    """, (project_id,))
    row = cursor.fetchone()
    setting_value = json.loads(row["SETTING_VALUE"]) if row and row.get("SETTING_VALUE") else {}

    # 2) 결함 원본
    cursor.execute("""
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
    """, (project_id,))
    defect_rows = cursor.fetchall()

    # 3) 역 정보
    cursor.execute("SELECT INITIAL, NAME, LINE, STATION_ORDER FROM SUBWAY_STATIONS;")
    station_rows = cursor.fetchall()
    cursor.close()
    connection.close()

    STATION_MAP  = {r["INITIAL"]: r["NAME"] for r in station_rows}
    ST_ORDER_MAP = {(r["LINE"], r["INITIAL"]): int(r["STATION_ORDER"]) for r in station_rows}

    # 가공
    for r in defect_rows:
        r["DEFECT_TYPE_KR"] = DEFECT_TYPE_MAP.get(r["DEFECT_TYPE"], r["DEFECT_TYPE"])
        parts = (r["IMAGE_NUM"] or "").split("_")
        if len(parts) >= 5:
            line, from_init, to_init, direction, ins_no = parts
            r["LINE"] = line
            r["LINE_KR"] = LINE_MAP.get(line, line)
            r["FROM_INIT"] = from_init
            r["TO_INIT"]   = to_init
            r["FROM_STATION"] = STATION_MAP.get(from_init, from_init)
            r["TO_STATION"]   = STATION_MAP.get(to_init, to_init)
            r["FROM_ORDER"] = ST_ORDER_MAP.get((line, from_init), 10**9)
            r["TO_ORDER"]   = ST_ORDER_MAP.get((line, to_init),   10**9)
            r["DIRECTION"]  = "상행" if direction.upper() == "UP" else ("하행" if direction.upper() == "DOWN" else direction)
            r["INSULATOR_NO"] = str(int(ins_no))
            r["SUB_PROJECT_ID"] = "_".join(parts[:-1])
            r["SUBPROJECT_KR"] = f"{r['FROM_STATION']}~{r['TO_STATION']}({r['DIRECTION']})"

    # 애자 개수 맵
    insulator_count_map = {
        k: v["INSULATOR_COUNT"]
        for k, v in setting_value.items()
        if isinstance(v, dict) and "INSULATOR_COUNT" in v
    }
    print(f"[INFO] 프로젝트 설정 키 수: {len(setting_value)} / 결함행 수: {len(defect_rows)}")
    return setting_value, insulator_count_map, defect_rows

def build_cover_meta_for_prefix(setting_value, filtered_rows, sub_project_prefix: str):
    # 라인명
    line_name = None
    if filtered_rows:
        lines = [r.get("LINE_KR") for r in filtered_rows if r.get("LINE_KR")]
        if lines:
            cnt = {}
            for ln in lines: cnt[ln] = cnt.get(ln, 0) + 1
            line_name = max(cnt.items(), key=lambda x: x[1])[0]
    if not line_name:
        code = (sub_project_prefix.split("_")[0] if sub_project_prefix else
                next((k.split("_")[0] for k in setting_value.keys() if k.startswith("ST")), "ST2"))
        line_name = LINE_MAP.get(code, code)

    # 섹션(시작~끝) 계산 (STATION_ORDER 기반)
    df = pd.DataFrame(filtered_rows)
    if not df.empty:
        df["START_ORDER"] = df[["FROM_ORDER","TO_ORDER"]].min(axis=1)
        df["END_ORDER"]   = df[["FROM_ORDER","TO_ORDER"]].max(axis=1)
        df = df.sort_values(["LINE_KR","START_ORDER","END_ORDER"])
        first, last = df.iloc[0], df.iloc[-1]
        start = str(first["FROM_STATION"] if first["FROM_ORDER"]<=first["TO_ORDER"] else first["TO_STATION"])
        end   = str(last["TO_STATION"]   if last["FROM_ORDER"]<=last["TO_ORDER"] else last["FROM_STATION"])
        section_core = f"{start} ~ {end}"
    else:
        section_core = setting_value.get("subwaySection","구간미상")

    inspector_raw = setting_value.get("inspector","")
    inspector = ", ".join([x.strip() for x in inspector_raw.split(",") if x.strip()])
    cover_meta = {
        "title_line": f"대구 지하철 {line_name.replace('호선','')}호선 {section_core.split('~')[0].strip()} ~ {section_core.split('~')[-1].strip()} 상태평가 보고서",
        "facility_name": setting_value.get("facilityName",""),
        "managed_number": setting_value.get("managedNumber",""),
        "written_date": date.today().strftime("%Y년 %m월 %d일"),
        "place": f"{line_name} {section_core}",
        "inspector": inspector,
        "writer": "Deep Inspector(AI 안전점검 프로그램)",
        "approver": setting_value.get("approver",""),
        "line_name": line_name,
        "section_core": section_core,
    }
    return cover_meta

def init_hwp(visible: bool):
    print("[INFO] 한글(HWP) 초기화...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = bool(visible)
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp

def write_cover(hwp, meta):
    act = hwp.HAction
    print("[INFO] 표지 작성...")
    act.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
    hwp.HParameterSet.HPageNumPos.DrawPos = 5
    act.Execute("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
    act.Run("ParagraphShapeAlignCenter")

    _char = hwp.HParameterSet.HCharShape
    act.GetDefault("CharShape", _char.HSet)
    _char.HSet.SetItem("Bold", 1)
    _char.Height = hwp.PointToHwpUnit(28)
    act.Execute("CharShape", _char.HSet)

    act.Run("BreakPara"); act.Run("BreakPara")
    ins = hwp.HParameterSet.HInsertText
    act.GetDefault("InsertText", ins.HSet)
    ins.Text = meta["title_line"]; act.Execute("InsertText", ins.HSet)

    act.Run("MoveDown")
    act.Run("ParagraphShapeAlignLeft")

    _char.HSet.SetItem("Bold", 1); _char.Height = hwp.PointToHwpUnit(15); act.Execute("CharShape", _char.HSet)
    for _ in range(13): act.Run("BreakPara")

    def put(label, value):
        act.GetDefault("InsertText", ins.HSet)
        ins.Text = f"{label:<6} : {value}"
        act.Execute("InsertText", ins.HSet)
        act.Run("BreakPara")

    put("설비명",    meta["facility_name"])
    put("관리번호",  meta["managed_number"])
    put("작성일",    meta["written_date"])
    put("검사장소",  meta["place"])
    put("점검자",    meta["inspector"])
    put("작성자",    meta["writer"])
    put("확인자",    meta["approver"])

    act.Run("MoveDocEnd"); hwp.Run("BreakPage")

def write_summary_table(hwp, rows):
    print(f"[INFO] 요약표 작성... (행 {len(rows)})")
    act = hwp.HAction
    ins = hwp.HParameterSet.HInsertText
    act.Run("ParagraphShapeAlignLeft")
    # 제목
    _char = hwp.HParameterSet.HCharShape
    act.GetDefault("CharShape", _char.HSet); _char.HSet.SetItem("Bold",1); _char.Height=hwp.PointToHwpUnit(15); act.Execute("CharShape", _char.HSet)
    act.GetDefault("InsertText", ins.HSet); ins.Text="2. 상태평가 결과"; act.Execute("InsertText", ins.HSet); act.Run("BreakPara")
    act.GetDefault("InsertText", ins.HSet); ins.Text="2.1 상태평가 요약표"; act.Execute("InsertText", ins.HSet); act.Run("BreakPara")
    act.Run("ParagraphShapeAlignCenter"); act.GetDefault("InsertText", ins.HSet); ins.Text=" [표4] 상태평가 요약표"; act.Execute("InsertText", ins.HSet); act.Run("BreakPara")

    # 표
    act.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = 1
    hwp.HParameterSet.HTableCreation.Cols = 6
    hwp.HParameterSet.HTableCreation.TableProperties.Width = 16
    act.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    act.Run("TableColWidthEven")

    headers = ["구분","위치","점검항목","점검기준/점검방법","점검결과","비고"]
    for i, head in enumerate(headers):
        act.GetDefault("InsertText", ins.HSet); ins.Text=head; act.Execute("InsertText", ins.HSet)
        if i < len(headers)-1: act.Run("TableRightCell")
    act.Run("TableAppendRow"); hwp.Run("TableColBegin")

    for row in rows:
        for i, val in enumerate(row):
            act.GetDefault("InsertText", ins.HSet); ins.Text=str(val); act.Execute("InsertText", ins.HSet)
            if i < len(headers)-1: act.Run("TableRightCell")
        act.Run("TableAppendRow"); hwp.Run("TableColBegin")
    act.Run("TableDeleteRow"); act.Run("Close"); hwp.Run("BreakPage")

def write_detail_table(hwp, rows):
    print(f"[INFO] 결과표(결함만) 작성... (행 {len(rows)})")
    act = hwp.HAction
    ins = hwp.HParameterSet.HInsertText
    act.Run("ParagraphShapeAlignLeft")
    _char = hwp.HParameterSet.HCharShape
    act.GetDefault("CharShape", _char.HSet); _char.HSet.SetItem("Bold",1); _char.Height=hwp.PointToHwpUnit(15); act.Execute("CharShape", _char.HSet)
    act.GetDefault("InsertText", ins.HSet); ins.Text="2.2 상태평가 결과 (결함 존재 애자만)"; act.Execute("InsertText", ins.HSet); act.Run("BreakPara")
    act.Run("ParagraphShapeAlignCenter"); act.GetDefault("InsertText", ins.HSet); ins.Text=" [표5] 상태평가 결과표(결함)"; act.Execute("InsertText", ins.HSet); act.Run("BreakPara")

    # 표
    act.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = 1
    hwp.HParameterSet.HTableCreation.Cols = 6
    hwp.HParameterSet.HTableCreation.TableProperties.Width = 16
    act.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    act.Run("TableColWidthEven")

    headers = ["구분","위치","점검항목","점검기준/점검방법","점검결과","비고"]
    for i, head in enumerate(headers):
        act.GetDefault("InsertText", ins.HSet); ins.Text=head; act.Execute("InsertText", ins.HSet)
        if i < len(headers)-1: act.Run("TableRightCell")
    act.Run("TableAppendRow"); hwp.Run("TableColBegin")

    for row in rows:
        for i, val in enumerate(row):
            act.GetDefault("InsertText", ins.HSet); ins.Text=str(val); act.Execute("InsertText", ins.HSet)
            if i < len(headers)-1: act.Run("TableRightCell")
        act.Run("TableAppendRow"); hwp.Run("TableColBegin")
    act.Run("TableDeleteRow"); act.Run("Close")

def build_summary_rows(filtered_df, ins_cnt_map, sub_ids_sorted):
    major = {"균열","파손"}
    caution = {"얼룩","아크혼(박리)","아크혼(그을음)"}
    rows = []
    for sub_id in sub_ids_sorted:
        df_sub = filtered_df[filtered_df["SUB_PROJECT_ID"]==sub_id]
        if df_sub.empty: continue
        sub_kr = df_sub.iloc[0]["SUBPROJECT_KR"]
        total_count = ins_cnt_map.get(sub_id, df_sub["INSULATOR_NO"].astype(int).max())
        replace_list, caution_list = [], []
        for ins_no in sorted(df_sub["INSULATOR_NO"].astype(int).unique()):
            dset = set(df_sub[df_sub["INSULATOR_NO"]==str(ins_no)]["DEFECT_TYPE_KR"])
            tag = f"#{ins_no}(" + ", ".join(sorted(dset)) + ")"
            if dset & major: replace_list.append(tag)
            elif dset & caution: caution_list.append(tag)
        replace_text = "교체 필요 애자: " + ", ".join(replace_list) if replace_list else "정상상태"
        caution_text = "주의 필요 애자: " + ", ".join(caution_list) if caution_list else "정상상태"
        rows.append(["지지물", f"{sub_kr} #1~#{total_count}", "지지애자 손상여부", "특이상태 점검/AI점검", replace_text, caution_text])
    return rows

def build_detail_rows_defects_only(filtered_df, ins_cnt_map, sub_ids_sorted):
    major   = {"균열","파손"}
    archorn = {"아크혼(박리)","아크혼(그을음)"}
    stain   = "얼룩"
    rows = []
    for sub_id in sub_ids_sorted:
        df_sub = filtered_df[filtered_df["SUB_PROJECT_ID"]==sub_id]
        if df_sub.empty: continue
        sub_kr = df_sub.iloc[0]["SUBPROJECT_KR"]
        # 결함 있는 애자만
        for ins_no in sorted(df_sub["INSULATOR_NO"].astype(int).unique()):
            dset = set(df_sub[df_sub["INSULATOR_NO"]==str(ins_no)]["DEFECT_TYPE_KR"])
            if not dset: continue
            major_found = sorted(list(dset & major))
            if len(major_found) >= 2:
                result = "균열/파손 검출"
            elif len(major_found) == 1:
                result = f"{major_found[0]} 검출"
            else:
                # 주요 결함 없고 주의만 있으면 결과는 '정상상태' 대신 공란 처리 가능
                result = "정상상태"
            remarks = []
            if dset & archorn: remarks.append("아크혼 검출")
            if stain in dset:  remarks.append("얼룩 검출")
            # 둘 다 없고 주요도 없으면 건너뛸 수도 있음
            if result == "정상상태" and not remarks:
                continue  # 진짜 결함 없음
            rows.append([
                "지지물",
                f"{sub_kr}\n#{ins_no}",
                "지지애자 손상여부",
                "특이상태 점검/AI점검",
                result,
                ", ".join(remarks)
            ])
    return rows

def run(out_dir: str, project_id: int, sub_project_id: str, visible: bool, dry_run: bool):
    print(f"[RUN] 부분 상세 보고서 시작  project_id={project_id}, sub_project_id={sub_project_id}, out={out_dir}, visible={visible}, dry_run={dry_run}")
    setting_value, ins_cnt_map_all, defect_rows = fetch_data_from_db(project_id)

    # 대상 prefix(UP/DOWN 포함)
    targets = {f"{sub_project_id}_UP", f"{sub_project_id}_DOWN"}
    df = pd.DataFrame(defect_rows)
    if df.empty:
        print("[WARN] 결함 데이터가 없습니다.")
    df = df[df["SUB_PROJECT_ID"].isin(targets)].copy()

    # 정렬키(호선/역순/방향)
    if not df.empty:
        df["START_ORDER"] = df[["FROM_ORDER","TO_ORDER"]].min(axis=1)
        df["END_ORDER"]   = df[["FROM_ORDER","TO_ORDER"]].max(axis=1)
        df["DIR_CODE"]    = df["SUB_PROJECT_ID"].str.split("_").str[-1].str.upper()
        df["DIR_ORDER"]   = df["DIR_CODE"].map({"UP":0,"DOWN":1}).fillna(2).astype(int)

    # 표지 메타
    cover_meta = build_cover_meta_for_prefix(setting_value, df.to_dict("records"), sub_project_id)

    # 서브프로젝트 정렬 목록
    sub_order = (
        df[["SUB_PROJECT_ID","SUBPROJECT_KR","LINE_KR","START_ORDER","END_ORDER","DIR_ORDER"]]
        .drop_duplicates()
        .sort_values(["LINE_KR","START_ORDER","END_ORDER","DIR_ORDER"])
        if not df.empty else pd.DataFrame(columns=["SUB_PROJECT_ID","SUBPROJECT_KR"])
    )
    sub_ids_sorted = sub_order["SUB_PROJECT_ID"].tolist()

    # 개수 맵 제한
    ins_cnt_map = {k: v for k, v in ins_cnt_map_all.items() if k in targets}

    # 표 데이터
    summary_rows = build_summary_rows(df, ins_cnt_map, sub_ids_sorted)
    detail_rows  = build_detail_rows_defects_only(df, ins_cnt_map, sub_ids_sorted)

    print(f"[INFO] 요약행 {len(summary_rows)} / 결과행(결함만) {len(detail_rows)}")

    if dry_run:
        print("[DRY-RUN] HWP 생성 스킵. 데이터 미리보기:")
        print("  서브프로젝트:", sub_ids_sorted)
        print("  요약표 1행 예:", summary_rows[:1])
        print("  결과표 1행 예:", detail_rows[:1])
        return

    # ==== HWP 작성 ====
    hwp = None
    try:
        hwp = init_hwp(visible)
        # 간단히 표지 + 요약 + 결과만 (목차/설명은 필요시 추가)
        write_cover(hwp, cover_meta)
        write_summary_table(hwp, summary_rows)
        write_detail_table(hwp, detail_rows)

        # 저장 파일명
        line_name = cover_meta["line_name"].replace(" ", "")
        section_clean = cover_meta["section_core"].replace("~","_").replace(" ","")
        filename_base = f"대구지하철_{line_name}_{section_clean}_부분_상세_상태평가보고서"
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(out_dir, f"{filename_base}.hwp")
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(out_dir, f"{filename_base}.pdf")
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        print(f"[OK] 저장 완료: {os.path.join(out_dir, filename_base)}.hwp/.pdf")

    finally:
        if hwp:
            try:
                hwp.Clear(option=1)
                hwp.Quit()
            except Exception:
                pass

def parse_args():
    p = argparse.ArgumentParser(description="DTRO 부분 상세 상태평가 보고서 생성")
    p.add_argument("--out", required=True, help="출력 폴더 경로")
    p.add_argument("--project-id", type=int, required=True, help="PROJECT_ID")
    p.add_argument("--sub_project_id", required=True, help="예: ST3_MPY_PSS (UP/DOWN 자동 포함)")
    p.add_argument("--visible", type=int, default=1, help="한글 창 표시(1) / 숨김(0)")
    p.add_argument("--dry-run", type=int, default=0, help="HWP 생성없이 데이터만 검증(1)")
    return p.parse_args()

if __name__ == "__main__":
    try:
        args = parse_args()
        print("[BOOT] dtro_dtl_statevl 시작")
        run(
            out_dir=args.out,
            project_id=args.project_id,
            sub_project_id=args.sub_project_id,
            visible=bool(args.visible),
            dry_run=bool(args.dry_run),
        )
        print("[DONE] 종료")
    except Exception as e:
        print("[FATAL] 실행 중 예외 발생:", e)
        traceback.print_exc()
        sys.exit(1)
