Attribute VB_Name = "单位上标"
Sub 单位上标()
    Dim rng As Range
    Dim counter As Long
    
    Set rng = ActiveDocument.Range
    counter = 0
    
    With rng.Find
        .ClearFormatting
        .text = "m[2-9]{1,}"  ' 匹配m后1个及以上数字
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            ' 定位到数字部分（不移动结束位置！）
            rng.MoveStart wdCharacter, 1
            
            ' 记录数字部分的范围
            Dim numRng As Range
            Set numRng = rng.Duplicate
            numRng.End = rng.End
            
            ' 设置数字为上标
            numRng.Font.Superscript = True
            counter = counter + 1
            
            ' 重置范围继续搜索
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    MsgBox "单位上标转换完成！共处理 " & counter & " 处", vbInformation
End Sub

