Attribute VB_Name = "规划报告一键宏"
Sub 规划报告一键宏_重构()
    '
    ' 重构版：执行文档格式标准化
    ' 功能：1.替换文本 2.设置段落格式 3.设置字体格式
    '
    
    On Error GoTo ErrorHandler
    
    Dim bkm As Range
    Set bkm = Selection.Range  ' 保存原始位置，处理完恢复
    
    ' 关闭屏幕更新，提高执行速度
    Application.ScreenUpdating = False
    
    ' 1. 文本替换："其它" -> "其他"
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "其它"
        .Replacement.text = "其他"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 2. 文本替换："东流流经" -> "向东流经" (新增功能)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "东流流经"
        .Replacement.text = "向东流经"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 3. 段落格式：两端对齐 -> 左对齐
    ' 方法1：通过查找替换（保留原录制逻辑思路）
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = ""  ' 空文本匹配任意内容
        .ParagraphFormat.alignment = wdAlignParagraphJustify
        .Replacement.text = ""  ' 不改变文本
        .Replacement.ParagraphFormat.alignment = wdAlignParagraphLeft
        .Format = True  ' 重要：应用格式查找
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 方法2：直接遍历所有段落（更高效，推荐）
    ' 如果想用这种方法，替换上面的方法1即可
    ' Dim para As Paragraph
    ' For Each para In ActiveDocument.Paragraphs
    '     If para.alignment = wdAlignParagraphJustify Then
    '         para.alignment = wdAlignParagraphLeft
    '     End If
    ' Next para
    
    ' 4. 字体格式：仿宋 -> 宋体
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = ""  ' 空文本匹配任意内容
        .Font.NameFarEast = "仿宋"  ' 查找中文字体为仿宋
        .Replacement.text = ""  ' 不改变文本
        .Replacement.Font.NameFarEast = "宋体"  ' 改为宋体
        .Format = True  ' 重要：应用格式查找
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 5. 可选：如果需要同时修改英文字体
    ' 添加这段代码可以修改英文字体
    ' 删除下一行的 ' 注释符号启用
    ' ActiveDocument.Content.Font.Name = "Times New Roman"  ' 英文字体
    
    ' 恢复原始选择和屏幕更新
    bkm.Select
    Application.ScreenUpdating = True
    
    MsgBox "文档格式标准化完成！" & vbCrLf & _
           "1. '其它'已替换为'其他'" & vbCrLf & _
           "2. '东流流经'已替换为'向东流经'" & vbCrLf & _
           "3. 段落对齐：两端对齐→左对齐" & vbCrLf & _
           "4. 字体：仿宋→宋体", _
           vbInformation, "操作完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "执行出错：" & Err.Description & vbCrLf & _
           "错误号：" & Err.Number, _
           vbExclamation, "错误"
    Application.ScreenUpdating = True
    bkm.Select
End Sub

