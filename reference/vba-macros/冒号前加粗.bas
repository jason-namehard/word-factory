Attribute VB_Name = "冒号前加粗"
Sub 冒号前加粗()
    Dim selRange As Range
    Set selRange = Selection.Range
    
    If selRange.Start = selRange.End Then
        MsgBox "请先选择文本"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 处理每一句话
    For Each sentence In selRange.Sentences
        Dim txt As String
        txt = sentence.text
        
        ' 查找冒号位置
        Dim colonPos As Long
        colonPos = InStr(txt, "：")  ' 中文冒号
        If colonPos = 0 Then colonPos = InStr(txt, ":")  ' 英文冒号
        
        If colonPos > 0 Then
            ' 创建加粗范围（包括冒号）
            Dim boldPart As Range
            Set boldPart = sentence.Duplicate
            boldPart.End = boldPart.Start + colonPos  ' 包括冒号
            
            ' 加粗处理
            boldPart.Font.Bold = True
        End If
    Next sentence
    
    Application.ScreenUpdating = True
    MsgBox "冒号及前面文字已加粗!", vbInformation
End Sub
