Sub HideSheets()
    Dim ws As Worksheet
    For Each ws In Sheets
        If Not ws.Name = ActiveSheet.Name And ws.Visible = xlSheetVisible Then
            ws.Visible = xlSheetHidden
        End If
    Next
'Change visibility of all sheets to Hide configurations (Except the activite sheet)
End Sub
