Attribute VB_Name = "MarkDown语言清除"
Sub MarkDown语言清除()
    ' 检查是否有文本被选中
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选择要转换的 Markdown 文本。", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Dim selectedRange As Range
    Set selectedRange = Selection.Range
    Dim originalText As String
    originalText = selectedRange.text
    
    Dim processedText As String
    processedText = RemoveMarkdownSymbolsWithQuotes(originalText)
    
    selectedRange.text = processedText
    Application.ScreenUpdating = True
    MsgBox "转换完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "转换过程中出现错误：" & Err.Description, vbCritical
End Sub

Function RemoveMarkdownSymbolsWithQuotes(ByVal text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    
    ' 核心修改：处理行内代码，将 `文本` 转换为 “文本”
    regEx.pattern = "`([^`]+)`"
    text = regEx.Replace(text, "“$1”")
    
    ' 其余 Markdown 标记清理规则保持不变
    regEx.pattern = "\*\*([^*]+)\*\*"
    text = regEx.Replace(text, "$1")
    
    regEx.pattern = "\*([^*]+)\*"
    text = regEx.Replace(text, "$1")
    
    Dim paragraphs As Variant
    paragraphs = Split(text, vbCr)
    
    Dim i As Integer
    For i = LBound(paragraphs) To UBound(paragraphs)
        Dim para As String
        para = paragraphs(i)
        
        If Left(Trim(para), 1) = "#" Then
            Dim tempPara As String
            tempPara = Trim(para)
            Do While Left(tempPara, 1) = "#"
                tempPara = Trim(Mid(tempPara, 2))
            Loop
            para = tempPara
        End If
        
        para = LTrim(para)
        If Left(para, 1) = "*" Or Left(para, 1) = "-" Or Left(para, 1) = "+" Then
            para = Trim(Mid(para, 2))
        End If
        
        para = Replace(para, "*", "")
        
        If para Like "#.*" Or para Like "##.*" Then
            Dim dotPos As Integer
            dotPos = InStr(para, ".")
            If dotPos > 0 Then
                Dim beforeDot As String
                Dim afterDot As String
                beforeDot = Left(para, dotPos)
                afterDot = Trim(Mid(para, dotPos + 1))
                para = beforeDot & afterDot
            End If
        End If
        
        Do While InStr(para, "  ") > 0
            para = Replace(para, "  ", " ")
        Loop
        
        para = Trim(para)
        paragraphs(i) = para
    Next i
    
    RemoveMarkdownSymbolsWithQuotes = Join(paragraphs, vbCr)
End Function
