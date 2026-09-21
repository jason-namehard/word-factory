Attribute VB_Name = "表格空格回车删除"
' 优化后的统一表格空格回车删除宏
' 提供两个入口点：全文修改和选定修改
' 基于用户需求合并共同逻辑，避免代码重复[3](@ref)

Option Explicit

' 主入口点1：处理文档中所有表格
Sub 表格空格回车删除_全文修改()
    Dim removeOption As Integer
    
    ' 获取用户选择
    removeOption = GetUserRemoveOption()
    If removeOption = 0 Then Exit Sub ' 用户取消操作
    
    ' 调用核心处理函数，False表示处理全文
    ProcessTablesInDocument removeOption, False
End Sub

' 主入口点2：仅处理选定区域内的表格
Sub 表格空格回车删除_选定修改()
    Dim removeOption As Integer
    
    ' 获取用户选择
    removeOption = GetUserRemoveOption()
    If removeOption = 0 Then Exit Sub ' 用户取消操作
    
    ' 调用核心处理函数，True表示处理选定区域
    ProcessTablesInDocument removeOption, True
End Sub

' 核心处理函数：根据参数决定处理范围
' removeOption: 删除选项（1=仅空格, 2=仅回车, 3=两者）
' processSelection: 是否只处理选定区域
Private Sub ProcessTablesInDocument(removeOption As Integer, processSelection As Boolean)
    Dim tbl As Table
    Dim currentCell As cell
    Dim cellRange As Range
    Dim originalText As String
    Dim newText As String
    Dim totalCellsProcessed As Long
    Dim modifiedCellsCount As Long
    Dim tableCollection As New Collection
    Dim tableItem As Variant
    Dim cellItem As Variant
    Dim hasValidTables As Boolean
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    totalCellsProcessed = 0
    modifiedCellsCount = 0
    hasValidTables = False
    
    ' 步骤1：收集要处理的表格
    For Each tbl In ActiveDocument.Tables
        If processSelection Then
            ' 只处理与选定区域相交的表格[2](@ref)
            If Not (Selection.Start > tbl.Range.End Or Selection.End < tbl.Range.Start) Then
                tableCollection.Add tbl
                hasValidTables = True
            End If
        Else
            ' 处理所有表格
            tableCollection.Add tbl
            hasValidTables = True
        End If
    Next tbl
    
    ' 检查是否找到有效表格
    If tableCollection.Count = 0 Then
        If processSelection Then
            MsgBox "在当前选定范围内没有找到表格！", vbExclamation, "提示"
        Else
            MsgBox "文档中没有找到任何表格！", vbExclamation, "提示"
        End If
        GoTo ExitSub
    End If
    
    ' 步骤2：处理收集到的表格中的单元格
    For Each tableItem In tableCollection
        Set tbl = tableItem
        For Each currentCell In tbl.Range.Cells
            ' 如果是选定区域模式，检查单元格是否在选定范围内
            If processSelection Then
                If Selection.Start <= currentCell.Range.End And _
                   Selection.End >= currentCell.Range.Start Then
                    totalCellsProcessed = totalCellsProcessed + 1
                    If ProcessSingleCell(currentCell, removeOption) Then
                        modifiedCellsCount = modifiedCellsCount + 1
                    End If
                End If
            Else
                ' 全文模式，处理所有单元格
                totalCellsProcessed = totalCellsProcessed + 1
                If ProcessSingleCell(currentCell, removeOption) Then
                    modifiedCellsCount = modifiedCellsCount + 1
                End If
            End If
        Next currentCell
    Next tableItem
    
    ' 步骤3：显示处理结果
    Dim resultMsg As String
    If processSelection Then
        resultMsg = "选定区域处理完成！"
    Else
        resultMsg = "全文处理完成！"
    End If
    resultMsg = resultMsg & "扫描了 " & totalCellsProcessed & " 个单元格，其中 " & modifiedCellsCount & " 个单元格的内容被清理。"
    
    MsgBox resultMsg, vbInformation, "完成"
    
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "处理过程中出现错误: " & Err.Description, vbCritical, "错误"
    Resume ExitSub
End Sub

' 处理单个单元格的核心函数
' 返回Boolean值表示单元格内容是否被修改
Private Function ProcessSingleCell(targetCell As cell, removeOption As Integer) As Boolean
    Dim cellRange As Range
    Dim originalText As String
    Dim newText As String
    Dim textChanged As Boolean
    
    textChanged = False
    On Error GoTo ErrorHandler
    
    Set cellRange = targetCell.Range
    ' 将范围缩小到单元格内容，排除单元格结束标记
    cellRange.MoveEnd wdCharacter, -1
    
    originalText = cellRange.text
    newText = originalText
    
    ' 只在有内容时处理
    If Len(originalText) > 0 Then
        ' 根据用户选择执行相应的清理操作
        Select Case removeOption
            Case 1  ' 仅删除空格
                newText = Replace(newText, " ", "")
                newText = Replace(newText, Chr(160), "") ' 不间断空格
            Case 2  ' 仅删除回车
                newText = Replace(newText, Chr(13), "")   ' 段落标记
                newText = Replace(newText, Chr(11), "")   ' 手动换行符
                newText = Replace(newText, Chr(7), "")    ' 表格标记
            Case 3  ' 删除空格和回车
                newText = Replace(newText, " ", "")
                newText = Replace(newText, Chr(160), "")
                newText = Replace(newText, Chr(13), "")
                newText = Replace(newText, Chr(11), "")
                newText = Replace(newText, Chr(7), "")
        End Select
        
        ' 检查文本是否发生变化
        If newText <> originalText Then
            cellRange.text = newText
            textChanged = True
        End If
    End If
    
    ProcessSingleCell = textChanged
    Exit Function
    
ErrorHandler:
    ProcessSingleCell = False
End Function

' 获取用户删除选项的通用函数
Private Function GetUserRemoveOption() As Integer
    Dim userInput As String
    Dim userChoice As Integer
    
    userInput = InputBox("请选择要删除的内容：" & vbCrLf & _
                       "1 - 仅删除空格" & vbCrLf & _
                       "2 - 仅删除回车" & vbCrLf & _
                       "3 - 删除空格和回车", "去除空格和回车", "3")
    
    If userInput = "" Then
        GetUserRemoveOption = 0 ' 用户取消
    Else
        ' 验证输入有效性
        If IsNumeric(userInput) Then
            userChoice = CInt(userInput)
            If userChoice >= 1 And userChoice <= 3 Then
                GetUserRemoveOption = userChoice
            Else
                MsgBox "请输入有效的选项（1-3）！", vbExclamation, "输入错误"
                GetUserRemoveOption = 0
            End If
        Else
            MsgBox "请输入数字选项（1-3）！", vbExclamation, "输入错误"
            GetUserRemoveOption = 0
        End If
    End If
End Function

