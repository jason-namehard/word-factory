Attribute VB_Name = "段落配方生成器"
Sub 段落配方生成器()
    Dim doc As Document
    Dim selectionType As Integer
    Dim specChar As String
    Dim recipeName As String
    Dim excelPath As String
    Dim sheetName As String
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim shouldCreateExcel As Boolean
    
    Set doc = ActiveDocument
    
    ' 检查是否有选中的内容
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "请先选中要生成配方的段落！", vbExclamation
        Exit Sub
    End If
    
    ' 获取用户选择
    selectionType = GetUserSelection()
    If selectionType = 0 Then Exit Sub
    
    If selectionType = 2 Then
        specChar = InputBox("请输入要识别的特定字符（如xx）：", "特定字符输入")
        If specChar = "" Then Exit Sub
    End If
    
    ' 询问是否创建Excel文件
    shouldCreateExcel = MsgBox("是否自动创建Excel表格？", vbYesNo + vbQuestion, "创建Excel") = vbYes
    
    If shouldCreateExcel Then
        ' 自动生成Excel文件
        excelPath = InputBox("请输入Excel文件名称（不含路径）：", "Excel文件", "数据表.xlsx")
        If excelPath = "" Then excelPath = "数据表.xlsx"
        
        sheetName = InputBox("请输入Sheet名称：", "Sheet名称", "Sheet1")
        If sheetName = "" Then sheetName = "Sheet1"
        
        ' 确保文件后缀
        If Not (Right(excelPath, 4) = ".xls" Or Right(excelPath, 5) = ".xlsx") Then
            excelPath = excelPath & ".xlsx"
        End If
    Else
        ' 使用现有Excel文件
        excelPath = InputBox("请输入Excel文件名称（不含路径）：", "Excel文件", "数据表.xlsx")
        If excelPath = "" Then excelPath = "数据表.xlsx"
        
        sheetName = InputBox("请输入Sheet名称：", "Sheet名称", "Sheet1")
        If sheetName = "" Then sheetName = "Sheet1"
    End If
    
    recipeName = InputBox("请输入段落配方名称：", "配方命名")
    If recipeName = "" Then recipeName = "默认配方"
    
    ' 处理选中的段落
    ProcessSelectedRange selectionType, specChar, excelPath, sheetName, recipeName, shouldCreateExcel
End Sub

Function GetUserSelection() As Integer
    Dim result As Variant
    
    result = InputBox("请选择识别模式：" & vbCrLf & _
                     "1 - 识别高亮文本" & vbCrLf & _
                     "2 - 识别特定字符", "模式选择", "1")
    
    If result = "" Then
        GetUserSelection = 0
    ElseIf result = "1" Then
        GetUserSelection = 1
    ElseIf result = "2" Then
        GetUserSelection = 2
    Else
        MsgBox "请输入有效的选择（1或2）", vbExclamation
        GetUserSelection = 0
    End If
End Function

Sub ProcessSelectedRange(selectionType As Integer, specChar As String, excelPath As String, sheetName As String, recipeName As String, shouldCreateExcel As Boolean)
    Dim selectedRng As Range
    Dim recipeRng As Range
    Dim varCount As Integer
    Dim recipeText As String
    Dim varPrefixes As Collection
    Dim varTexts As Collection
    
    Set selectedRng = Selection.Range
    Set recipeRng = ActiveDocument.Range
    recipeRng.SetRange ActiveDocument.Content.End, ActiveDocument.Content.End
    
    ' 初始化集合来存储变量前缀和文本
    Set varPrefixes = New Collection
    Set varTexts = New Collection
    
    ' 添加配方标题
    recipeRng.text = vbCrLf & "=== 段落配方 [" & recipeName & "] ===" & vbCrLf
    recipeRng.Collapse 0 ' wdCollapseEnd
    
    varCount = 0
    recipeText = ""
    
    If selectionType = 1 Then
        ' 高亮模式 - 使用新的隐形修正机制
        ProcessHighlightedTextWithFix selectedRng, excelPath, sheetName, varCount, recipeText, varPrefixes, varTexts
    Else
        ' 特定字符模式
        ProcessSpecialChars selectedRng, specChar, excelPath, sheetName, varCount, recipeText, varPrefixes, varTexts
    End If
    
    ' 保存配方文本
    recipeRng.text = recipeText & vbCrLf
    recipeRng.Collapse 0
    
    ' 添加配方信息头
    recipeRng.text = "EXCEL_FILE:" & excelPath & vbCrLf
    recipeRng.Collapse 0
    recipeRng.text = "SHEET_NAME:" & sheetName & vbCrLf
    recipeRng.Collapse 0
    recipeRng.text = "VARIABLE_COUNT:" & varCount & vbCrLf
    recipeRng.Collapse 0
    
    recipeRng.text = "=== 配方结束 ===" & vbCrLf & vbCrLf
    
    ' 如果需要创建Excel，则自动生成
    If shouldCreateExcel Then
        CreateExcelWithData excelPath, sheetName, varPrefixes, varTexts
    End If
    
    MsgBox "段落配方生成完成！共识别 " & varCount & " 个变量。" & IIf(shouldCreateExcel, " Excel表格已自动创建。", ""), vbInformation
End Sub

Sub ProcessHighlightedTextWithFix(selectedRng As Range, excelPath As String, sheetName As String, ByRef varCount As Integer, ByRef recipeText As String, ByRef varPrefixes As Collection, ByRef varTexts As Collection)
    Dim doc As Document
    Dim rng As Range
    Dim currentPos As Long
    Dim highlightStart As Long
    Dim isHighlighted As Boolean
    Dim prevHighlighted As Boolean
    Dim tempText As String
    Dim lastTextSegment As String
    
    Set doc = ActiveDocument
    Set rng = selectedRng.Duplicate
    currentPos = rng.Start
    prevHighlighted = False
    lastTextSegment = ""
    
    Debug.Print "=== 隐形修正高亮版本开始 ==="
    Debug.Print "原始文本: " & rng.text
    
    ' 遍历选中范围的每个字符
    For currentPos = rng.Start To rng.End
        Dim testRng As Range
        Set testRng = doc.Range(currentPos, currentPos)
        
        ' 检测高亮：wdNoHighlight = 0
        isHighlighted = (testRng.HighlightColorIndex <> wdNoHighlight)
        
        ' 高亮状态发生变化
        If isHighlighted <> prevHighlighted Then
            If isHighlighted Then
                ' 从非高亮切换到高亮
                If currentPos > rng.Start Then
                    ' 添加前面的非高亮文本
                    Dim textRng As Range
                    Set textRng = doc.Range(rng.Start, currentPos - 1)
                    If textRng.text <> "" Then
                        lastTextSegment = cleanText(textRng.text)
                        recipeText = recipeText & "TEXT:" & lastTextSegment & vbCrLf
                        Debug.Print "添加文本: '" & lastTextSegment & "'"
                    End If
                End If
                highlightStart = currentPos
            Else
                ' 从高亮切换到非高亮 - 应用隐形修正
                ' 将高亮区间缩减一个字符（最后一个字符划归到后面的文本）
                Dim correctedEnd As Long
                correctedEnd = currentPos - 1
                
                ' 确保高亮区间至少有一个字符
                If highlightStart < correctedEnd Then
                    ' 缩减高亮区间（去掉最后一个字符）
                    Dim varRng As Range
                    Set varRng = doc.Range(highlightStart, correctedEnd - 1)
                    
                    varCount = varCount + 1
                    recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
                    Debug.Print "添加变量: VAR:" & varCount & ", 内容: '" & varRng.text & "'"
                    
                    ' 保存变量前缀和文本
                    varPrefixes.Add lastTextSegment
                    varTexts.Add varRng.text
                    lastTextSegment = ""
                    
                    ' 更新范围起始位置（跳过被修正的字符）
                    rng.Start = correctedEnd
                Else
                    ' 如果高亮区间只有一个字符，不进行缩减
                    Set varRng = doc.Range(highlightStart, correctedEnd)
                    
                    varCount = varCount + 1
                    recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
                    Debug.Print "添加变量: VAR:" & varCount & ", 内容: '" & varRng.text & "'"
                    
                    ' 保存变量前缀和文本
                    varPrefixes.Add lastTextSegment
                    varTexts.Add varRng.text
                    lastTextSegment = ""
                    
                    rng.Start = currentPos
                End If
            End If
            prevHighlighted = isHighlighted
        End If
    Next currentPos
    
    ' 处理末尾
    If prevHighlighted Then
        ' 末尾是高亮 - 应用隐形修正
        Dim finalCorrectedEnd As Long
        finalCorrectedEnd = rng.End
        
        ' 确保高亮区间至少有一个字符
        If highlightStart < finalCorrectedEnd Then
            ' 缩减高亮区间（去掉最后一个字符）
            Set varRng = doc.Range(highlightStart, finalCorrectedEnd - 1)
            
            varCount = varCount + 1
            recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
            Debug.Print "添加末尾变量: VAR:" & varCount & ", 内容: '" & varRng.text & "'"
            
            ' 保存变量前缀和文本
            varPrefixes.Add lastTextSegment
            varTexts.Add varRng.text
            lastTextSegment = ""
            
            ' 将缩减的字符添加到后面的文本
            Set textRng = doc.Range(finalCorrectedEnd, finalCorrectedEnd)
            If textRng.text <> "" Then
                lastTextSegment = cleanText(textRng.text)
                recipeText = recipeText & "TEXT:" & lastTextSegment & vbCrLf
                Debug.Print "添加修正文本: '" & lastTextSegment & "'"
            End If
        Else
            ' 如果高亮区间只有一个字符，不进行缩减
            Set varRng = doc.Range(highlightStart, finalCorrectedEnd)
            
            varCount = varCount + 1
            recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
            Debug.Print "添加末尾变量: VAR:" & varCount & ", 内容: '" & varRng.text & "'"
            
            ' 保存变量前缀和文本
            varPrefixes.Add lastTextSegment
            varTexts.Add varRng.text
            lastTextSegment = ""
        End If
    ElseIf rng.Start <= rng.End Then
        ' 末尾是非高亮文本
        Set textRng = doc.Range(rng.Start, rng.End)
        If textRng.text <> "" Then
            lastTextSegment = cleanText(textRng.text)
            recipeText = recipeText & "TEXT:" & lastTextSegment & vbCrLf
            Debug.Print "添加末尾文本: '" & lastTextSegment & "'"
        End If
    End If
    
    Debug.Print "=== 隐形修正高亮版本结束 ==="
End Sub

Sub ProcessSpecialChars(selectedRng As Range, specChar As String, excelPath As String, sheetName As String, ByRef varCount As Integer, ByRef recipeText As String, ByRef varPrefixes As Collection, ByRef varTexts As Collection)
    Dim fullText As String
    Dim paragraphs As Variant
    Dim i As Long
    Dim paraText As String
    Dim foundPos As Long
    Dim currentPos As Long
    Dim textBefore As String
    Dim lastTextSegment As String
    
    ' 按段落分割处理
    fullText = selectedRng.text
    paragraphs = Split(fullText, vbCrLf)
    lastTextSegment = ""
    
    For i = 0 To UBound(paragraphs)
        paraText = paragraphs(i)
        currentPos = 1
        
        ' 如果段落为空，添加空行标记
        If Trim(paraText) = "" Then
            recipeText = recipeText & "TEXT:" & vbCrLf
            GoTo NextParagraph
        End If
        
        ' 处理段落内的特定字符
        Do
            foundPos = InStr(currentPos, paraText, specChar)
            
            If foundPos > 0 Then
                ' 添加特定字符之前的文本
                If foundPos > currentPos Then
                    textBefore = Mid(paraText, currentPos, foundPos - currentPos)
                    lastTextSegment = cleanText(textBefore)
                    recipeText = recipeText & "TEXT:" & lastTextSegment & vbCrLf
                End If
                
                ' 添加变量
                varCount = varCount + 1
                recipeText = recipeText & "VAR:" & varCount & "|" & excelPath & "!" & sheetName & "!B" & (varCount + 1) & vbCrLf
                
                ' 保存变量前缀和文本
                varPrefixes.Add lastTextSegment
                varTexts.Add specChar
                lastTextSegment = ""
                
                currentPos = foundPos + Len(specChar)
            Else
                ' 添加最后一段文本
                If currentPos <= Len(paraText) Then
                    textBefore = Mid(paraText, currentPos)
                    lastTextSegment = cleanText(textBefore)
                    recipeText = recipeText & "TEXT:" & lastTextSegment & vbCrLf
                End If
                Exit Do
            End If
        Loop
        
NextParagraph:
        ' 如果不是最后一段，添加段落分隔符
        If i < UBound(paragraphs) Then
            recipeText = recipeText & "TEXT:" & vbCrLf
        End If
    Next i
End Sub

Sub CreateExcelWithData(excelPath As String, sheetName As String, varPrefixes As Collection, varTexts As Collection)
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim fullPath As String
    Dim i As Integer
    
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    If excelApp Is Nothing Then
        MsgBox "无法创建Excel应用程序对象，请确保已安装Microsoft Excel。", vbExclamation
        Exit Sub
    End If
    
    ' 创建新工作簿
    Set excelWorkbook = excelApp.Workbooks.Add
    Set excelWorksheet = excelWorkbook.Worksheets(1)
    excelWorksheet.Name = sheetName
    
    ' 写入表头
    excelWorksheet.Cells(1, 1).Value = "项目"
    excelWorksheet.Cells(1, 2).Value = "数值"
    
    ' 写入变量前缀和初始值
    For i = 1 To varPrefixes.Count
        excelWorksheet.Cells(i + 1, 1).Value = varPrefixes(i)
        excelWorksheet.Cells(i + 1, 2).Value = varTexts(i)
    Next i
    
    ' 自动调整列宽
    excelWorksheet.Columns(1).AutoFit
    excelWorksheet.Columns(2).AutoFit
    
    ' 构建完整路径（保存到文档所在目录）
    fullPath = ActiveDocument.Path
    If fullPath = "" Then
        fullPath = Environ("USERPROFILE") & "\Documents"
    End If
    fullPath = fullPath & "\" & excelPath
    
    ' 保存工作簿
    On Error Resume Next
    excelWorkbook.SaveAs fullPath
    If Err.Number <> 0 Then
        MsgBox "保存Excel文件时出错: " & Err.Description, vbExclamation
        excelWorkbook.Close False
        excelApp.Quit
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 关闭工作簿和Excel
    excelWorkbook.Close
    excelApp.Quit
    
    ' 释放对象
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    
    MsgBox "Excel表格已创建并保存到: " & fullPath, vbInformation
End Sub

Function cleanText(text As String) As String
    ' 清理文本，移除多余的换行和空格
    Dim result As String
    result = Trim(text)
    ' 移除末尾的段落标记
    If Right(result, 1) = Chr(13) Then
        result = Left(result, Len(result) - 1)
    End If
    If Right(result, 1) = Chr(10) Then
        result = Left(result, Len(result) - 1)
    End If
    cleanText = result
End Function

