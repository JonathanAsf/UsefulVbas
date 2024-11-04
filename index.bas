
Sub Tabela_Conteudo()

    Dim StartCell As Range 'For inputbox to select range
    Dim Endcell As Range 'For message box as info
    Dim sh As Worksheet
    Dim ShName As String
    Dim msgConfirm As VBA.VbMsgBoxResult 'Confirm button
    

    
    On Error Resume Next
    
    Set StartCell = Excel.Application.InputBox("Where do you want to insert the content tables?" _
    & vbNewLine & "Please, select a cell:", "Insert Table of contents", , , , , , 8)
    If Err.Number = 424 Then Exit Sub
    On Error GoTo Handle
    Set StartCell = StartCell.Cells(1, 1)
    Set Endcell = StartCell.Offset(Worksheets.Count - 2, 1)
    
    'get confirmation
    msgConfirm = VBA.MsgBox("The value of the range cells: " & vbNewLine & StartCell.Address & Endcell.Address & _
    " Will be overwritten." & vbNewLine & "Do you want to continue?", vbDefaultButton2, "Confirmation")
    
    If msgConfirm = vbCancel Then Exit Sub
   
    For Each sh In Worksheets
        ShName = sh.Name
        If ActiveSheet.Name <> ShName Then
            If sh.Visible = xlSheetVisible Then
                ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _
                "'" & ShName & "'!A1", TextToDisplay:=ShName
                StartCell.Offset(0, 1).Value = sh.Range("A1").Value
                Set StartCell = StartCell.Offset(1, 0)
            End If
        End If
    Next sh
    Exit Sub
Handle:
MsgBox "Unfortunately an error has occurred. Contact the program developer"
End Sub


