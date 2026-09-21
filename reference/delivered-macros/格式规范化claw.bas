Attribute VB_Name = "格式规范化claw"

Sub 格式规范化_claw()
Dim rng As Object, st As Integer
For st = 1 To 11
    On Error Resume Next: Set rng = ActiveDocument.StoryRanges(st)
    Do While Not rng Is Nothing: rng.Font.ColorIndex = wdBlack: Set rng = rng.NextStoryRange: Loop
Next
For st = 1 To 11
    On Error Resume Next: Set rng = ActiveDocument.StoryRanges(st)
    Do While Not rng Is Nothing: rng.HighlightColorIndex = wdNoHighlight: Set rng = rng.NextStoryRange: Loop
Next
Dim fr As Object: Set fr = ActiveDocument.Content
With fr.Find: .ClearFormatting: .Text = "m[0-9]{1,}": .MatchWildcards = True: .Forward = True: .Wrap = wdFindStop: End With
Dim oe As Long: oe = fr.End
Do While fr.Find.Execute
    If fr.Find.Found Then
        Dim s As Long, e As Long: s = fr.Start + 1: e = fr.End
        If s < e Then ActiveDocument.Range(s, e).Font.Superscript = True
        fr.Start = fr.End: fr.End = oe
        If fr.Start >= oe Then Exit Do
    Else: Exit Do: End If
Loop
End Sub
