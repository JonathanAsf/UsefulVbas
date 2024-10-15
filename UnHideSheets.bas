Sub UnHiddeSheets()
    Dim ws As Worksheet
  
    For Each ws In Sheets
        If Not ws.Name = ActiveSheet.Name And ws.Visible = xlSheetHidden Then
            ws.Visible = True
        End If
    Next
'Change visibility to Visible of all Sheets that was Hidden
End Sub
