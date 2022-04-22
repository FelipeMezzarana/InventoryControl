Private Sub classe_afterupdate()

If Me.classe = "" Then
    Exit Sub
End If

If Me.classe = "A" Or Me.classe = "B" Or Me.classe = "C" Then
    Exit Sub
Else
    MsgBox "Invalid class." & vbCrLf & vbCrLf & "Choose between: A, B or C", vbCritical, "Attention"
    Me.classe = ""
End If

End Sub

Private Sub codigo_Change()
Call carregaitem
End Sub

Private Sub CommandButton20_Click()

Call salva_edicao
Edit_Userform.Image47.Picture = Images_Userform.certo.Picture
MsgBox "Editing done successfully!", vbInformation, Edit_Userform.codigo

End Sub




Private Sub tipo_Change()

If Me.tipo = "" Then
    Exit Sub
End If

If Me.tipo = "M" Or Me.tipo = "F" Or Me.tipo = "E" Then
    Exit Sub
Else
    MsgBox "Invalid type." & vbCrLf & vbCrLf & "Valid types: " & "M --> Mecânico " & vbCrLf & "F --> Facilities " & vbCrLf & " E --> Elétrica ", vbCritical, "Attention"
    Me.tipo = ""
End If
End Sub

Private Sub UserForm_Initialize()


'função para habilitar botao max min
Call HabilitaBotoes(Me)

Edit_Userform.classe.AddItem "A"
Edit_Userform.classe.AddItem "B"
Edit_Userform.classe.AddItem "C"

Edit_Userform.tipo.AddItem "M"
Edit_Userform.tipo.AddItem "F"
Edit_Userform.tipo.AddItem "E"

Call carrega_combobox_cod_editar
End Sub