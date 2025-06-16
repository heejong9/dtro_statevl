import win32com.client as win32

# HWP 앱 실행
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# 새 문서
hwp.XHwpWindows.Item(0).Visible = True
hwp.Create("Empty")

# 1. 데이터베이스 목록
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 2  # 제목 + 데이터
hwp.HParameterSet.HTableCreation.Cols = 4  # ID, 명칭, 주관부서, 비고
hwp.HParameterSet.HTableCreation.WidthType = 2  # 전체폭 고정
hwp.HParameterSet.HTableCreation.TableWidth = 200
hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 4)
col_widths = [50, 50, 50, 50]
for idx, width in enumerate(col_widths):
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(idx, width)
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

# 첫번째 표 타이틀 입력
titles = ["ID", "명칭", "주관부서", "비고"]
for title in titles:
    hwp.Run("TableCellBlock")
    hwp.Run("TableCellSelect")
    hwp.InsertText(title)
    hwp.Run("MoveRight")

# 줄바꿈
hwp.MovePos(3)
hwp.InsertBreak()

# 2. 데이터베이스 정의
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 2  # 제목 + 데이터
hwp.HParameterSet.HTableCreation.Cols = 10
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HParameterSet.HTableCreation.TableWidth = 250
hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 10)
col_widths2 = [25] * 10
for idx, width in enumerate(col_widths2):
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(idx, width)
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

titles2 = ["Bufferpool", "TS ID", "TS 용량", "테이블 ID", "테이블 명", "Storage Group", "인덱스 BP", "인덱스 ID", "인덱스 용량", "비고"]
for title in titles2:
    hwp.Run("TableCellBlock")
    hwp.Run("TableCellSelect")
    hwp.InsertText(title)
    hwp.Run("MoveRight")

# 파일 저장
output_path = r"C:\Temp\데이터베이스_정의서.hwp"
hwp.SaveAs(output_path)
hwp.Quit()

print(f"파일 저장 완료: {output_path}")
