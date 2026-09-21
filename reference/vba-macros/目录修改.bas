Attribute VB_Name = "目录修改"
Sub 删除低于当前选定目录层级的目录()

    Dim para As Paragraph
    Dim indentValue As Single
    Dim selectedLevel As Integer
    
    ' 检查当前选择的段落
    If Selection.Type = wdNoSelection Then
        MsgBox "请先选中一个目录项！"
        Exit Sub
    End If

    ' 获取选中段落的缩进值（层级）
    indentValue = Selection.paragraphs(1).LeftIndent
    
    ' 根据缩进值判断选中的目录层级
    Select Case indentValue
        Case 0
            selectedLevel = 1 ' 一级目录
        Case 28
            selectedLevel = 2 ' 二级目录
        Case 56
            selectedLevel = 3 ' 三级目录
        Case Else
            MsgBox "未识别的目录层级"
            Exit Sub
    End Select
    
    ' 遍历文档中的所有段落
    For Each para In ActiveDocument.paragraphs
        indentValue = para.LeftIndent
        
        ' 删除低于选中层级的目录
        If (selectedLevel = 1 And indentValue >= 28) Or _
           (selectedLevel = 2 And indentValue >= 56) Then
            para.Range.Delete
        End If
    Next para

End Sub


