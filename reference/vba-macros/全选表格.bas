Attribute VB_Name = "全选表格"
Sub 全选表格()
    Dim tbl As Table
    Application.ScreenUpdating = False
    For Each tbl In ActiveDocument.Tables
        tbl.Range.Editors.Add wdEditorEveryone
    Next
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    Application.ScreenUpdating = True
End Sub

