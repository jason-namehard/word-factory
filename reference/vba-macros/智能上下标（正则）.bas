Attribute VB_Name = "智能上下标"
' =======================================================
' 智能上下标工具 - 正则表达式重构版
' 功能：基于正则表达式智能识别并应用上下标格式
' 特点：
'   - 特殊符号最高优先级
'   - 边界符号优先级低于字典符号
'   - 清晰的区间排除逻辑
'   - 保留性能优化
' =======================================================

' 定义常量
Const MAX_CHAR_LIMIT As Long = 1230 ' 固定字符数限制

Sub 智能上下标调整()
    On Error GoTo ErrorHandler
    
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
    
    ' 检查字符数限制（保留原有性能优化）
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
    
    ' 使用正则表达式识别符号
    Dim symbols As Collection
    Set symbols = 正则表达式识别符号(selectedText, commonRules, specialRules)
    
    ' 应用格式规则
    If symbols.Count > 0 Then
        应用格式规则 selRange, symbols, commonRules, specialRules
    End If
    
    ' 显示完成提示
    MsgBox "处理完成，共识别并格式化 " & symbols.Count & " 个符号", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "发生错误：" & Err.Description & vbCrLf & "错误号：" & Err.Number, vbCritical, "错误"
End Sub

' ==================== 规则字典初始化 ====================
Private Sub 初始化规则字典(commonRules As Object, specialRules As Object)
    ' 清除可能存在的重复键
    commonRules.RemoveAll
    specialRules.RemoveAll
    
    ' 特殊符号规则（最高优先级）- 即使与其他符号冲突也优先匹配
    ' 修改点：特殊符号拥有最高权限，即使它是其他符号的子集
    specialRules.Add "°C", "SN"    ' 摄氏度（°C）
    specialRules.Add "mm", "NN"    ' 毫米
    ' 可以继续添加其他特殊符号...
    
    ' 普通符号规则（字母开头），N正常、S上标、B下标
    
    commonRules.Add "Vmax", "NBBB"   ' Vmax
    commonRules.Add "Qpl", "NBB"    ' Qpl (pl下标)
    commonRules.Add "m2", "NS"      ' m2
    commonRules.Add "m3", "NS"      ' m3
    commonRules.Add "m4", "NS"      ' m4
    commonRules.Add "m5", "NS"      ' m5
    commonRules.Add "cm2", "NNS"    ' cm2
    commonRules.Add "km2", "NNS"    ' km2
    commonRules.Add "qm", "NB"      ' qm
    commonRules.Add "Qm", "NB"      ' Qm
    commonRules.Add "KP", "NB"      ' KP
    commonRules.Add "CV", "NB"      ' CV
    commonRules.Add "H24P", "NBBB"  ' H24P
    commonRules.Add "hR", "NB"      ' hR
    commonRules.Add "H24", "NBB"    ' H24
    commonRules.Add "Q4ml", "NBSS"   ' Q4ml
    commonRules.Add "Q4el", "NBSS"   ' Q4el
    commonRules.Add "Pt", "NB"       ' Pt
    commonRules.Add "Q4al+pl", "NBSSSSS" ' Q4al+pl（包含+的复杂符号）
End Sub

' ==================== 正则表达式识别符号（核心重构） ====================
Private Function 正则表达式识别符号(text As String, commonRules As Object, specialRules As Object) As Collection
    Dim symbols As New Collection
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 构建正则表达式模式：特殊符号优先，然后普通符号
    Dim pattern As String
    pattern = 构建正则模式(commonRules, specialRules)
    
    ' 诊断信息：显示构建的正则模式
    Debug.Print "构建的正则模式: " & pattern
    
    With regEx
        .Global = True
        .IgnoreCase = False ' 大小写敏感
        .MultiLine = True
        .pattern = pattern
    End With
    
    ' 执行匹配
    Dim matches As Object
    Dim match As Object
    Dim matchText As String
    Dim isSpecial As Boolean
    
    If regEx.Test(text) Then
        Set matches = regEx.Execute(text)
        
        For Each match In matches
            matchText = match.Value
            isSpecial = specialRules.Exists(matchText) ' 检查是否为特殊符号
            
            ' 诊断信息：显示每个匹配项
            Debug.Print "匹配到: """ & matchText & """, 位置: " & (match.FirstIndex + 1) & "-" & (match.FirstIndex + match.Length) & ", 特殊符号: " & isSpecial
            
            ' 添加到符号集合
            symbols.Add Array(match.FirstIndex + 1, match.FirstIndex + match.Length, matchText, isSpecial)
        Next match
    End If
    
    ' 诊断信息：显示总匹配数
    Debug.Print "总匹配符号数: " & symbols.Count
    
    Set 正则表达式识别符号 = symbols
    Set regEx = Nothing
End Function

' ==================== 构建正则模式（核心逻辑） ====================
Private Function 构建正则模式(commonRules As Object, specialRules As Object) As String
    Dim pattern As String
    Dim key As Variant
    Dim keys() As String
    Dim i As Long, j As Long
    Dim temp As String
    
    ' 收集所有键（特殊符号 + 普通符号）
    Dim allKeys As Collection
    Set allKeys = New Collection
    
    ' 首先添加特殊符号（最高优先级）
    For Each key In specialRules.keys
        allKeys.Add key
    Next key
    
    ' 然后添加普通符号
    For Each key In commonRules.keys
        ' 检查是否已经是特殊符号，避免重复
        If Not specialRules.Exists(key) Then
            allKeys.Add key
        End If
    Next key
    
    ' 按键长度排序（长符号优先，避免短符号错误匹配长符号的一部分）
    ReDim keys(1 To allKeys.Count)
    For i = 1 To allKeys.Count
        keys(i) = allKeys(i)
    Next i
    
    ' 简单冒泡排序（按长度降序）
    For i = 1 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If Len(keys(i)) < Len(keys(j)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    ' 构建正则模式，对每个键进行转义
    pattern = ""
    For i = 1 To UBound(keys)
        pattern = pattern & "|" & 转义正则表达式(keys(i))
    Next i
    
    If Len(pattern) > 0 Then
        pattern = "(" & Mid(pattern, 2) & ")" ' 移除开头的"|"
    End If
    
    构建正则模式 = pattern
End Function

' ==================== 正则表达式转义 ====================
Private Function 转义正则表达式(text As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    
    ' 需要转义的正则元字符
    Dim specialChars As String
    specialChars = "\+.?*|{}[]()^$"
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        ' 如果是元字符，则转义
        If InStr(specialChars, char) > 0 Then
            result = result & "\" & char
        Else
            result = result & char
        End If
    Next i
    
    转义正则表达式 = result
End Function

' ==================== 应用格式规则（优化版） ====================
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
        
        ' 获取格式规则（优先特殊符号）
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

' ==================== 获取格式规则（优化版） ====================
Private Function 获取格式规则(symbolText As String, isSpecial As Boolean, _
                             commonRules As Object, specialRules As Object) As String
    ' 修改点1：特殊符号绝对优先
    If isSpecial Then
        If specialRules.Exists(symbolText) Then
            获取格式规则 = specialRules(symbolText)
            Exit Function
        End If
    End If
    
    ' 修改点2：普通符号匹配
    If commonRules.Exists(symbolText) Then
        获取格式规则 = commonRules(symbolText)
        Exit Function
    End If
    
    ' 修改点3：清除格式后再次尝试匹配
    Dim cleanText As String
    cleanText = 清除格式(symbolText)
    
    If cleanText <> symbolText Then
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
    End If
    
    ' 无匹配规则
    获取格式规则 = ""
End Function

' ==================== 清除格式（保留原逻辑） ====================
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

' ==================== 应用规则到范围（保留原逻辑） ====================
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

' ==================== 辅助函数（保留） ====================
Private Function IsLetter(char As String) As Boolean
    IsLetter = (char >= "a" And char <= "z") Or (char >= "A" And char <= "Z")
End Function

Private Function IsDigit(char As String) As Boolean
    IsDigit = (char >= "0" And char <= "9")
End Function

