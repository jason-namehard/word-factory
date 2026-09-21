Attribute VB_Name = "段落重配"
Sub 段落重配()
    Dim doc As Document
    Dim selectedRng As Range
    Dim recipeText As String
    Dim excelPath As String
    Dim sheetName As String
    Dim variableCount As Integer
    
    Set doc = ActiveDocument
    
    ' 检查是否有选中的内容
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "请先选中段落配方！", vbExclamation
        Exit Sub
    End If
    
    Set selectedRng = Selection.Range
    recipeText = selectedRng.text
    
    ' 解析配方信息
    If Not ParseRecipeInfo(recipeText, excelPath, sheetName, variableCount) Then
        MsgBox "未能正确解析段落配方！" & vbCrLf & _
               "请确认选中了完整的配方文本", vbExclamation
        Exit Sub
    End If
    
    ' 获取Excel数据
    Dim excelData() As String
    ReDim excelData(1 To variableCount)
    
    If Not GetExcelData(excelPath, sheetName, variableCount, excelData) Then
        MsgBox "无法从Excel获取数据！请检查文件路径和Sheet名称。", vbExclamation
        Exit Sub
    End If
    
    ' 重建段落
    ReconstructParagraph recipeText, excelData, variableCount
End Sub

Function ParseRecipeInfo(recipeText As String, ByRef excelPath As String, ByRef sheetName As String, ByRef variableCount As Integer) As Boolean
    Dim lines() As String
    Dim i As Long
    Dim line As String
    
    ' 处理换行符
    recipeText = Replace(recipeText, vbCrLf, "?")
    recipeText = Replace(recipeText, vbCr, "?")
    recipeText = Replace(recipeText, vbLf, "?")
    lines = Split(recipeText, "?")
    
    excelPath = ""
    sheetName = ""
    variableCount = 0
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line = "" Then GoTo NextLine
        
        If Left(line, 11) = "EXCEL_FILE:" Then
            excelPath = Trim(Mid(line, 12))
        ElseIf Left(line, 11) = "SHEET_NAME:" Then
            sheetName = Trim(Mid(line, 12))
        ElseIf Left(line, 15) = "VARIABLE_COUNT:" Then
            Dim countStr As String
            countStr = Trim(Mid(line, 16))
            If IsNumeric(countStr) Then
                variableCount = Val(countStr)
            End If
        End If
        
NextLine:
    Next i
    
    ParseRecipeInfo = (excelPath <> "" And sheetName <> "" And variableCount > 0)
End Function

Function GetExcelData(excelPath As String, sheetName As String, variableCount As Integer, ByRef excelData() As String) As Boolean
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim i As Long
    Dim fullPath As String
    
    On Error GoTo ErrorHandler
    
    ' 获取完整路径
    fullPath = GetFullExcelPath(excelPath)
    If fullPath = "" Then
        GetExcelData = False
        Exit Function
    End If
    
    ' 启动Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    
    ' 打开工作簿
    Set excelWorkbook = excelApp.Workbooks.Open(fullPath)
    
    ' 获取工作表
    On Error Resume Next
    Set excelWorksheet = excelWorkbook.Sheets(sheetName)
    If excelWorksheet Is Nothing Then
        excelWorkbook.Close False
        excelApp.Quit
        GetExcelData = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' 读取数据
    For i = 1 To variableCount
        Dim cellAddress As String
        cellAddress = "B" & (i + 1)
        
        Dim cellValue As Variant
        cellValue = excelWorksheet.Range(cellAddress).Value
        
        If IsError(cellValue) Then
            excelData(i) = "#错误#"
        ElseIf IsNull(cellValue) Or cellValue = "" Then
            excelData(i) = ""
        Else
            excelData(i) = CStr(cellValue)
        End If
    Next i
    
    ' 关闭Excel
    excelWorkbook.Close False
    excelApp.Quit
    
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    
    GetExcelData = True
    Exit Function
    
ErrorHandler:
    If Not excelWorkbook Is Nothing Then
        On Error Resume Next
        excelWorkbook.Close False
    End If
    If Not excelApp Is Nothing Then
        On Error Resume Next
        excelApp.Quit
    End If
    GetExcelData = False
End Function

Function GetFullExcelPath(fileName As String) As String
    Dim basePath As String
    Dim testPath As String
    
    ' 检查是否已经是完整路径
    If InStr(fileName, "\") > 0 Then
        If Dir(fileName) <> "" Then
            GetFullExcelPath = fileName
            Exit Function
        End If
    End If
    
    ' 检查文档所在目录
    basePath = ActiveDocument.Path
    If basePath <> "" Then
        If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
        testPath = basePath & fileName
        If Dir(testPath) <> "" Then
            GetFullExcelPath = testPath
            Exit Function
        End If
    End If
    
    ' 检查当前目录
    If Dir(fileName) <> "" Then
        GetFullExcelPath = fileName
        Exit Function
    End If
    
    ' 没找到文件
    GetFullExcelPath = ""
End Function

Sub ReconstructParagraph(recipeText As String, excelData() As String, variableCount As Integer)
    Dim lines() As String
    Dim i As Long
    Dim currentVar As Integer
    Dim outputText As String
    Dim outputRng As Range
    Dim inParagraphSection As Boolean
    Dim previousLineWasText As Boolean
    
    ' 处理换行符
    recipeText = Replace(recipeText, vbCrLf, "?")
    recipeText = Replace(recipeText, vbCr, "?")
    recipeText = Replace(recipeText, vbLf, "?")
    lines = Split(recipeText, "?")
    
    outputText = ""
    currentVar = 0
    inParagraphSection = False
    previousLineWasText = False
    
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        If line = "" Then
            ' 空行表示段落分隔
            If inParagraphSection Then
                outputText = outputText & vbCrLf
                previousLineWasText = False
            End If
            GoTo NextLine
        End If
        
        ' 检查是否进入段落部分
        If InStr(line, "=== 段落配方") > 0 Then
            inParagraphSection = True
            GoTo NextLine
        ElseIf InStr(line, "=== 配方结束") > 0 Then
            inParagraphSection = False
            GoTo NextLine
        End If
        
        If Not inParagraphSection Then
            GoTo NextLine
        End If
        
        ' 跳过信息行
        If Left(line, 11) = "EXCEL_FILE:" Or Left(line, 11) = "SHEET_NAME:" Or Left(line, 15) = "VARIABLE_COUNT:" Then
            GoTo NextLine
        End If
        
        ' 处理文本和变量行
        If Left(line, 5) = "TEXT:" Then
            Dim textContent As String
            textContent = Mid(line, 6)
            
            ' 处理分段逻辑
            If textContent = "" Then
                ' 空的TEXT:表示换行
                outputText = outputText & vbCrLf
                previousLineWasText = False
            Else
                ' 如果前一行不是TEXT，且当前TEXT以数字开头（如"2、土壤改良工程"），则添加换行
                If Not previousLineWasText And IsNumeric(Left(textContent, 1)) Then
                    outputText = outputText & vbCrLf
                End If
                
                outputText = outputText & textContent
                previousLineWasText = True
            End If
        ElseIf Left(line, 4) = "VAR:" Then
            currentVar = currentVar + 1
            If currentVar <= variableCount Then
                outputText = outputText & excelData(currentVar)
            Else
                outputText = outputText & "#数据缺失#"
            End If
            previousLineWasText = False
        Else
            ' 处理没有TEXT:前缀的文本行（可能是配方中的换行）
            If line <> "" Then
                outputText = outputText & vbCrLf & line
                previousLineWasText = True
            End If
        End If
        
NextLine:
    Next i
    
    ' 创建新的段落
    Set outputRng = ActiveDocument.Range
    outputRng.SetRange ActiveDocument.Content.End, ActiveDocument.Content.End
    outputRng.text = vbCrLf & vbCrLf & "=== 重建段落 ===" & vbCrLf & outputText & vbCrLf & "=== 结束 ===" & vbCrLf
    
    MsgBox "段落重建完成！共使用了 " & currentVar & " 个变量。", vbInformation
End Sub

