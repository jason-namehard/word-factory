Attribute VB_Name = "目录修改_claw"

Sub 静默更新目录页码_claw()
Dim toc As TableOfContents
For Each toc In ActiveDocument.TablesOfContents
    toc.UpdatePageNumbers
Next
End Sub
