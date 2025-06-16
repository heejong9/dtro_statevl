import win32com.client as win32

# 한글 오브젝트 호출
hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True



# 텍스트 입력
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
# 문단 모양 설정 기본값 가져오기
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

# 들여쓰기 설정 (첫 줄 들여쓰기만 적용)
hwp.HParameterSet.HParaShape.Indentation = 1000.0

# 문단 모양 적용
# hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

# hwp.HParameterSet.HInsertText.Text = "첫 줄입니다.\\n두 번째 줄입니다."
# hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
# 문단 모양 설정: 첫 줄 들여쓰기만 적용ds





Set = hwp.HParameterSet.HParaShape
hwp.HAction.GetDefault("ParagraphShape", Set.HSet)
tab_def = Set.TabDef
tab_def.CreateItemArray("TabItem", 3)
tab_def.TabItem.SetItem(0, 80000)
tab_def.TabItem.SetItem(1, 3)
tab_def.TabItem.SetItem(2, 0)
hwp.HAction.Execute("ParagraphShape", Set.HSet)

# 텍스트 삽입
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "목차 항목1\t1\n"
hwp.HParameterSet.HInsertText.Text += "목차 항목1\t1\n"
hwp.HParameterSet.HInsertText.Text += "목차 항목1\t1"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

