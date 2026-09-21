Attribute VB_Name = "表格边框调整claw"

Sub 表格边框调整_claw()
Dim tbl As Table: Application.ScreenUpdating = False
For Each tbl In ActiveDocument.Tables
    With tbl.Borders(wdBorderTop): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderBottom): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderLeft): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderRight): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt: End With
    With tbl.Borders(wdBorderInsideHorizontal): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: End With
    With tbl.Borders(wdBorderInsideVertical): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: End With
Next: Application.ScreenUpdating = True
End Sub
