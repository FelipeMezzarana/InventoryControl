
'Userform
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
Edição.Image47.Picture = Imagens.certo.Picture
MsgBox "Editing done successfully!", vbInformation, Edição.codigo

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

Edição.classe.AddItem "A"
Edição.classe.AddItem "B"
Edição.classe.AddItem "C"

Edição.tipo.AddItem "M"
Edição.tipo.AddItem "F"
Edição.tipo.AddItem "E"

Call carrega_combobox_cod_editar
End Sub
