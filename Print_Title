Private Sub Workbook_BeforePrint(Cancel As Boolean)
 Dim Sh As Worksheet
 For Each Sh In Me.Worksheets
    Sh.PageSetup.LeftHeader = Sh.Range("A2").Value
    'I'm considering that the title It's on A2 cell
End Sub
