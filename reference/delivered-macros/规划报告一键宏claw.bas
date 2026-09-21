Attribute VB_Name = "规划报告一键宏claw"

Sub 文档格式标准化_claw()
'=================== CONFIG 配置区 ====================
Private Const CFG_替换文本_1 As String = "旧文本A"
Private Const CFG_替换为_1 As String = "新文本A"
Private Const CFG_启用文本替换_1 As Boolean = False
Private Const CFG_替换文本_2 As String = "旧文本B"
Private Const CFG_替换为_2 As String = "新文本B"
Private Const CFG_启用文本替换_2 As Boolean = False
Private Const CFG_原字体 As String = "仿宋_GB2312"
Private Const CFG_新字体 As String = "宋体"
Private Const CFG_启用字体替换 As Boolean = False
Private Const CFG_原对齐 As Integer = 3  ' 3=两端对齐(wdAlignParagraphJustify)
Private Const CFG_新对齐 As Integer = 0  ' 0=左对齐(wdAlignParagraphLeft)
Private Const CFG_启用段落调整 As Boolean = False
Private Const CFG_原字号 As Single = 12
Private Const CFG_新字号 As Single = 14
Private Const CFG_启用字号 As Boolean = False
'======================================================
'--- 执行体 ---
Application.ScreenUpdating = False
'1) 文本替换1
If CFG_启用文本替换_1 Then
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = CFG_替换文本_1
        .Replacement.ClearFormatting
        .Replacement.Text = CFG_替换为_1
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End If
'2) 文本替换2
If CFG_启用文本替换_2 Then
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = CFG_替换文本_2
        .Replacement.ClearFormatting
        .Replacement.Text = CFG_替换为_2
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End If
'3) 字体替换
If CFG_启用字体替换 Then
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.Name = CFG_原字体
        .Replacement.ClearFormatting
        .Replacement.Font.Name = CFG_新字体
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End If
'4) 段落对齐
If CFG_启用段落调整 Then
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Alignment = CFG_原对齐 Then
            para.Alignment = CFG_新对齐
        End If
    Next
End If
'5) 字号调整
If CFG_启用字号 Then
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.Size = CFG_原字号
        .Replacement.ClearFormatting
        .Replacement.Font.Size = CFG_新字号
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End If
Application.ScreenUpdating = True
End Sub
