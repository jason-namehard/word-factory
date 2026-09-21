Attribute VB_Name = "格式规范化"
Sub 格式规范化()
    Dim rng As Range, toc As TableOfContents
    Dim counter As Long, bHasTOC As Boolean
    Dim findRng As Range, numRng As Range
    Dim originalStart As Long, originalEnd As Long
    
    ' 初始化
    counter = 0
    bHasTOC = False
    
    ' === 功能1：全文变黑 ===
    On Error Resume Next
    For Each rng In ActiveDocument.StoryRanges
        Do
            rng.Font.Color = wdColorBlack
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng
    On Error GoTo 0
    
    ' === 功能2：去掉高亮 ===
    On Error Resume Next
    For Each rng In ActiveDocument.StoryRanges
        Do
            rng.HighlightColorIndex = wdNoHighlight
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng
    On Error GoTo 0
    
    ' === 功能3：特征文字变为上标 ===
    Set findRng = ActiveDocument.Range
    
    ' 记录原范围，用于查找完成后重置
    originalStart = findRng.Start
    originalEnd = findRng.End
    
    On Error Resume Next
    With findRng.Find
        .ClearFormatting
        .text = "m[0-9]{1,}"  ' 匹配m后1个及以上数字
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue  ' 改为继续查找，避免提前停止
        
        Do While .Execute
            ' 只有当找到匹配项时才处理
            If findRng.Find.Found Then
                ' 复制当前找到的范围
                Set rng = findRng.Duplicate
                
                ' 移动起始位置到数字部分（跳过字母"m"）
                rng.MoveStart wdCharacter, 1
                
                ' 设置数字部分为上标
                rng.Font.Superscript = True
                counter = counter + 1
                
                ' 重置查找范围，避免死循环
                findRng.Collapse wdCollapseEnd
                findRng.End = originalEnd
            Else
                Exit Do
            End If
        Loop
    End With
    On Error GoTo 0
    
    ' === 功能4：更新目录页码 ===
    On Error Resume Next
    If ActiveDocument.TablesOfContents.Count > 0 Then
        Set toc = ActiveDocument.TablesOfContents(1)
        If Not toc Is Nothing Then
            toc.UpdatePageNumbers
            bHasTOC = True
        End If
    End If
    On Error GoTo 0
    
    ' 第一次回车：完成提示
    Dim msg As String
    msg = "格式化完成！" & vbNewLine & vbNewLine
    msg = msg & "执行结果汇总：" & vbNewLine
    msg = msg & "? 全文已设置为黑色" & vbNewLine
    msg = msg & "? 已移除所有高亮" & vbNewLine
    msg = msg & "? 单位上标处理：" & counter & " 处" & vbNewLine
    If bHasTOC Then
        msg = msg & "? 目录页码已更新"
    Else
        msg = msg & "? 未找到目录，跳过页码更新"
    End If
    
    MsgBox msg, vbInformation, "格式化完成"
    
    ' 第二次回车：保存并关闭确认
    If MsgBox("是否保存并关闭当前文档？", vbYesNo + vbQuestion, "保存确认") = vbYes Then
        ' 保存文档
        If ActiveDocument.Path = "" Then
            Application.Dialogs(wdDialogFileSaveAs).Show
        Else
            ActiveDocument.Save
        End If
        
        ' 关闭文档
        ActiveDocument.Close SaveChanges:=wdSaveChanges
        
        ' 完成提示
        MsgBox "文档已保存并关闭。", vbInformation, "操作完成"
    Else
        MsgBox "文档保持打开状态，修改已生效但未保存。", vbInformation, "已取消"
    End If
End Sub
