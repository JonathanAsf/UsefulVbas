
Sub Tabela_Conteudo()

    Dim StartCell As Range 'for inputbox to select range
    Dim Endcell As Range 'For message box as info
    Dim sh As Worksheet
    Dim ShName As String
    Dim msgConfirm As VBA.VbMsgBoxResult 'Confirm button
    

    
    On Error Resume Next
    
    Set StartCell = Excel.Application.InputBox("Onde voce deseja inserir a tabela de conteudos?" _
    & vbNewLine & "Por favor selecione uma celula:", "Insert Table of contents", , , , , , 8)
    If Err.Number = 424 Then Exit Sub
    On Err GoTo Handle
    Set StartCell = StartCell.Cells(1, 1)
    Set Endcell = StartCell.Offset(Worksheets.Count - 2, 1)
    
    'get confirmation
    msgConfirm = VBA.MsgBox("Os valores da celulas de: " & vbNewLine & StartCell.Address & " ate " & Endcell.Address & _
    " podem ser reescritos." & vbNewLine & "Tem certeza que deseja prosseguir?", vbDefaultButton2, "Confirmacao necessaria")
    
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
MsgBox "Infelizmente um erro ocorreu. Contate o Desenvolvedor do programa"
End Sub


