Attribute VB_Name = "表格空格回车删除claw"

Sub 表格空格删除_claw()
Dim tbl As Table, r As Long, c As Long
If ActiveDocument.Tables.Count = 0 Then Exit Sub
Application.ScreenUpdating = False
For Each tbl In ActiveDocument.Tables
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim txt As String
            txt = tbl.Cell(r, c).Range.Text
            Do While Len(txt) > 0
                Dim ch As String: ch = Right(txt, 1)
                If ch = Chr(13) Or ch = Chr(7) Or ch = " " Or ch = vbTab Then
                    txt = Left(txt, Len(txt) - 1)
                Else: Exit Do: End If
            Loop
            If Len(txt) > 0 Then tbl.Cell(r, c).Range.Text = txt
        Next
    Next
Next
Application.ScreenUpdating = True
End Sub
