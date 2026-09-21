Attribute VB_Name = "智能上下标"
' =======================================================
' 智能上下标工具 - 修复等号边界版
' 功能：智能识别并应用上下标格式
' 特点：
'   - 等号作为明确边界字符
'   - 精确符号识别
'   - 修复错误格式化问题
' =======================================================

' 定义状态枚举
Enum ParseState
    STATE_START = 0
    STATE_IN_SYMBOL = 1
    STATE_IN_NUMBER = 2
    STATE_IN_SPECIAL = 3
    STATE_IN_UNIT = 4 ' 单位符号状态
End Enum

' 定义常量
Const MAX_CHAR_LIMIT As Long = 1230 ' 固定字符数限制

Sub 智能上下标调整()
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选中要格式化的文本", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 获取选中文本
    Dim selRange As Range
    Set selRange = Selection.Range
    
    ' 使用Word精确字符计数（包括空格）
    Dim charCount As Long
    charCount = selRange.ComputeStatistics(wdStatisticCharacters)
    
    ' 检查字符数限制
    If charCount > MAX_CHAR_LIMIT Then
        MsgBox "选中文本超过 " & MAX_CHAR_LIMIT & " 个字符，无法处理。" & vbCrLf & _
               "请选择不超过 " & MAX_CHAR_LIMIT & " 字符的文本后重试。", vbExclamation, "字符数过多"
        Exit Sub
    End If
    
    ' 获取选中文本内容
    Dim selectedText As String
    selectedText = selRange.text
    
    ' 如果文本以段落标记结尾，去除它
    If Right(selectedText, 1) = vbCr Then
        selectedText = Left(selectedText, Len(selectedText) - 1)
    End If
    
    ' 初始化规则字典
    Dim commonRules As Object
    Dim specialRules As Object
    Set commonRules = CreateObject("Scripting.Dictionary")
    Set specialRules = CreateObject("Scripting.Dictionary")
    初始化规则字典 commonRules, specialRules
    
    ' 使用FSM识别符号
    Dim symbols As Collection
    Set symbols = 识别符号(selectedText, commonRules, specialRules)
    
    ' 应用格式规则
    If symbols.Count > 0 Then
        应用格式规则 selRange, symbols, commonRules, specialRules
    End If
    
    ' 显示完成提示
    MsgBox "处理完成", vbInformation, "完成"
End Sub

' ==================== 规则字典 ====================
Private Sub 初始化规则字典(commonRules As Object, specialRules As Object)
    ' 普通符号规则（字母开头），N正常、S上标、B下标
    commonRules.Add "Vmax", "NBBB"   ' Vmax
    commonRules.Add "Qpl", "NBB"    ' Qpl (pl下标)
     
    ' 单位符号规则（数字+单位）
    commonRules.Add "m2", "NS"      ' m2
    commonRules.Add "m3", "NS"      ' m3
    commonRules.Add "m4", "NS"      ' m?
    commonRules.Add "m5", "NS"      ' m?
    commonRules.Add "cm2", "NNS"    ' cm2
    commonRules.Add "km2", "NNS"    ' km2
    commonRules.Add "qm", "NB" ' kg/m3
    commonRules.Add "Qm", "NB" ' kg/m3
    commonRules.Add "KP", "NB" ' kg/m3
    commonRules.Add "CV", "NB" ' kg/m3
    commonRules.Add "H24P", "NBBB" ' kg/m3
    commonRules.Add "hR", "NB" ' kg/m3
    commonRules.Add "H24", "NBB" ' kg/m3
    
    ' 特殊符号规则（非字母开头）
    specialRules.Add "°C", "SN"    ' 摄氏度（°C）
    specialRules.Add "mm", "NN"    ' 毫米
End Sub

' ==================== 标点符号边界检测（包含等号） ====================
Private Function IsBoundaryChar(char As String) As Boolean
    ' 定义边界标点符号（包含等号）
    Dim boundaryChars As String
    boundaryChars = "，。！？；：""""=-+*×/÷（）【】《》＝％"
    
    ' 检查是否为边界字符
    If InStr(boundaryChars, char) > 0 Then
        IsBoundaryChar = True
    Else
        IsBoundaryChar = False
    End If
End Function

' ==================== 改进的FSM符号识别 ====================
Private Function 识别符号(text As String, commonRules As Object, specialRules As Object) As Collection
    Dim symbols As New Collection
    Dim state As ParseState
    state = STATE_START
    Dim startPos As Long
    startPos = 1
    Dim i As Long
    Dim currentChar As String
    Dim unitStartPos As Long ' 记录单位开始位置
    
    For i = 1 To Len(text)
        currentChar = Mid(text, i, 1)
        
        ' 检查是否为边界字符（包含等号）
        If IsBoundaryChar(currentChar) Then
            ' 边界字符，结束当前状态
            Select Case state
                Case STATE_IN_SYMBOL
                    Dim symbolText As String
                    symbolText = Mid(text, startPos, i - startPos)
                    If Len(symbolText) > 0 Then
                        symbols.Add Array(startPos, i - 1, symbolText, False)
                    End If
                    state = STATE_START
                Case STATE_IN_UNIT
                    Dim unitText As String
                    unitText = Mid(text, unitStartPos, i - unitStartPos)
                    If Len(unitText) > 0 Then
                        symbols.Add Array(unitStartPos, i - 1, unitText, False)
                    End If
                    state = STATE_START
                Case STATE_IN_NUMBER
                    state = STATE_START
            End Select
            ' 跳过边界字符
            GoTo Continue
        End If
        
        Select Case state
            Case STATE_START
                ' 检查是否进入特殊符号
                If specialRules.Exists(currentChar) Then
                    ' 单字符特殊符号
                    symbols.Add Array(i, i, currentChar, True)
                ElseIf i < Len(text) Then
                    Dim twoChar As String
                    twoChar = Mid(text, i, 2)
                    If specialRules.Exists(twoChar) Then
                        symbols.Add Array(i, i + 1, twoChar, True)
                        i = i + 1 ' 跳过下一个字符
                    Else
                        ' 检查是否进入数字开头的单位
                        If IsDigit(currentChar) Then
                            state = STATE_IN_NUMBER
                            startPos = i
                        ' 检查是否进入普通符号
                        ElseIf IsLetter(currentChar) Then
                            state = STATE_IN_SYMBOL
                            startPos = i
                        End If
                    End If
                End If
                
            Case STATE_IN_NUMBER
                ' 数字后遇到字母，进入单位状态
                If IsLetter(currentChar) Then
                    state = STATE_IN_UNIT
                    unitStartPos = i ' 记录单位开始位置
                ' 数字后遇到非字母数字，结束
                ElseIf Not (IsDigit(currentChar) Or IsLetter(currentChar)) Then
                    state = STATE_START
                End If
                
            Case STATE_IN_UNIT
                ' 单位符号结束条件
                If Not (IsLetter(currentChar) Or IsDigit(currentChar)) Then
                    ' 提取单位部分
                    unitText = Mid(text, unitStartPos, i - unitStartPos)
                    symbols.Add Array(unitStartPos, i - 1, unitText, False)
                    state = STATE_START
                End If
                
            Case STATE_IN_SYMBOL
                ' 符号结束条件（不包含等号）
                If Not (IsLetter(currentChar) Or IsDigit(currentChar)) Then
                    ' 符号结束
                    symbolText = Mid(text, startPos, i - startPos)
                    symbols.Add Array(startPos, i - 1, symbolText, False)
                    state = STATE_START
                End If
        End Select
        
Continue:
    Next i
    
    ' 处理最后一个符号 - 添加严格边界检查
    If state = STATE_IN_UNIT Then
        If unitStartPos <= Len(text) Then
            unitText = Mid(text, unitStartPos, Len(text) - unitStartPos + 1)
            symbols.Add Array(unitStartPos, Len(text), unitText, False)
        End If
    ElseIf state = STATE_IN_SYMBOL Then
        If startPos <= Len(text) Then
            symbolText = Mid(text, startPos, Len(text) - startPos + 1)
            symbols.Add Array(startPos, Len(text), symbolText, False)
        End If
    ElseIf state = STATE_IN_NUMBER Then
        If startPos <= Len(text) Then
            symbolText = Mid(text, startPos, Len(text) - startPos + 1)
            symbols.Add Array(startPos, Len(text), symbolText, False)
        End If
    End If
    
    Set 识别符号 = symbols
End Function

' ==================== 应用格式规则 ====================
Private Sub 应用格式规则(selRange As Range, symbols As Collection, commonRules As Object, specialRules As Object)
    Dim i As Long
    For i = 1 To symbols.Count
        Dim symbolInfo As Variant
        symbolInfo = symbols(i)
        
        Dim startPos As Long
        Dim endPos As Long
        Dim symbolText As String
        Dim isSpecial As Boolean
        
        startPos = symbolInfo(0)
        endPos = symbolInfo(1)
        symbolText = symbolInfo(2)
        isSpecial = symbolInfo(3)
        
        ' 严格位置有效性检查
        If startPos < 1 Or endPos < startPos Or endPos > Len(selRange.text) Then
            GoTo Continue
        End If
        
        ' 获取格式规则
        Dim rule As String
        rule = 获取格式规则(symbolText, isSpecial, commonRules, specialRules)
        
        ' 只对找到规则的符号应用格式
        If rule <> "" Then
            On Error Resume Next ' 防止范围错误
            Dim symbolRange As Range
            Set symbolRange = selRange.Duplicate
            symbolRange.SetRange _
                Start:=selRange.Start + startPos - 1, _
                End:=selRange.Start + endPos
            
            If Err.Number = 0 Then
                应用规则到范围 symbolRange, rule
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
Continue:
    Next i
End Sub

' ==================== 改进的规则匹配 ====================
Private Function 获取格式规则(symbolText As String, isSpecial As Boolean, _
                             commonRules As Object, specialRules As Object) As String
    ' 1. 优先匹配特殊符号
    If isSpecial Then
        If specialRules.Exists(symbolText) Then
            获取格式规则 = specialRules(symbolText)
            Exit Function
        End If
    Else
        If commonRules.Exists(symbolText) Then
            获取格式规则 = commonRules(symbolText)
            Exit Function
        End If
    End If
    
    ' 2. 尝试清除格式后匹配
    Dim cleanText As String
    cleanText = 清除格式(symbolText)
    
    If isSpecial Then
        If specialRules.Exists(cleanText) Then
            获取格式规则 = specialRules(cleanText)
            Exit Function
        End If
    Else
        If commonRules.Exists(cleanText) Then
            获取格式规则 = commonRules(cleanText)
            Exit Function
        End If
    End If
    
    ' 3. 智能单位检测（严格限制）
    If Not isSpecial Then
        ' 只对明确匹配的模式应用智能检测
        If cleanText = "m2" Then
            获取格式规则 = "NS"
            Exit Function
        End If
        
        If cleanText = "m3" Then
            获取格式规则 = "NS"
            Exit Function
        End If
        
        ' 严格限制：仅对明确以m开头且长度为2的符号应用上标
        If Left(cleanText, 1) = "m" And Len(cleanText) = 2 And IsDigit(Mid(cleanText, 2, 1)) Then
            获取格式规则 = "NS"
            Exit Function
        End If
        
        ' 严格限制：仅对明确以2结尾且长度大于1的符号应用上标
        If Right(cleanText, 1) = "2" And Len(cleanText) > 1 And IsLetter(Left(cleanText, 1)) Then
            获取格式规则 = String(Len(cleanText) - 1, "N") & "S"
            Exit Function
        End If
        
        ' 严格限制：仅对明确以3结尾且长度大于1的符号应用上标
        If Right(cleanText, 1) = "3" And Len(cleanText) > 1 And IsLetter(Left(cleanText, 1)) Then
            获取格式规则 = String(Len(cleanText) - 1, "N") & "S"
            Exit Function
        End If
    End If
    
    ' 无匹配
    获取格式规则 = ""
End Function

Private Function 清除格式(text As String) As String
    ' 移除所有非字母数字字符（保留字母和数字）
    Dim result As String
    Dim i As Long
    result = ""
    
    For i = 1 To Len(text)
        Dim char As String
        char = Mid(text, i, 1)
        
        If (char >= "a" And char <= "z") Or _
           (char >= "A" And char <= "Z") Or _
           (char >= "0" And char <= "9") Then
            result = result & char
        End If
    Next i
    
    清除格式 = result
End Function

' ==================== 应用规则到范围 ====================
Private Sub 应用规则到范围(rng As Range, rule As String)
    ' 应用新格式
    Dim i As Long
    For i = 1 To Len(rule)
        If i <= rng.Characters.Count Then
            Select Case Mid(rule, i, 1)
                Case "S"  ' 上标
                    rng.Characters(i).Font.Superscript = True
                    rng.Characters(i).Font.Subscript = False
                Case "B"  ' 下标
                    rng.Characters(i).Font.Subscript = True
                    rng.Characters(i).Font.Superscript = False
                Case "N"  ' 正常
                    rng.Characters(i).Font.Superscript = False
                    rng.Characters(i).Font.Subscript = False
            End Select
        End If
    Next i
End Sub

' ==================== 辅助函数 ====================
Private Function IsLetter(char As String) As Boolean
    IsLetter = (char >= "a" And char <= "z") Or (char >= "A" And char <= "Z")
End Function

Private Function IsDigit(char As String) As Boolean
    IsDigit = (char >= "0" And char <= "9")
End Function

