'Userform
Private Sub classe_Change()

If Me.classe = "" Then
    Exit Sub
End If


If Me.classe = "A" Or Me.classe = "B" Or Me.classe = "C" Then
    Exit Sub
Else
    MsgBox "Invalid class." & vbCrLf & vbCrLf & "Choose between: A, B or C", vbCritical, "Atenção"
    Me.classe = ""
End If


End Sub

Private Sub CommandButton1_Click()
Call gerar_codigo

End Sub

Private Sub CommandButton20_Click()

For Each ctl In item_novo.Controls
    Select Case TypeName(ctl)
        Case "TextBox", "combobox"
            If ctl.Text = "" Then
                MsgBox "Record NOT SAVED" & vbCrLf & vbCrLf & "All text boxes must be filled in", vbCritical, "Attention"
                Exit Sub
            End If
    End Select
Next ctl



Call novo_cadastro
Me.Image47.Picture = Imagens.certo.Picture
MsgBox "Item successfully registered!", vbInformation, "Register"

Call apagar_item_novo

End Sub


Private Sub est_max_Change()
If est_max = "" Then
    Exit Sub
End If

If Not IsNumeric(est_max) Then
    MsgBox "Please enter a numeric value", vbCritical, "Attention"
    est_max = ""
    Exit Sub
End If
End Sub

Private Sub est_min_Change()

If est_min = "" Then
    Exit Sub
End If

If Not IsNumeric(est_min) Then
    MsgBox "Please enter a numeric value", vbCritical, "Attention"
    est_min = ""
    Exit Sub
End If
End Sub

Private Sub Label4_Click()

End Sub

Private Sub saldo_Change()
If saldo = "" Then
    Exit Sub
End If

If Not IsNumeric(saldo) Then
    MsgBox "Please enter a numeric value", vbCritical, "Attention"
    saldo = ""
    Exit Sub
End If
End Sub

Private Sub tipo_Change()

If Me.tipo = "" Then
    Exit Sub
End If

If Me.tipo = "M" Or Me.tipo = "F" Or Me.tipo = "E" Then
    Exit Sub
Else
    MsgBox "Invalid type." & vbCrLf & vbCrLf & "Valid types: " & vbCrLf & "M --> Mecânico " & vbCrLf & "F --> Facilities " & vbCrLf & "E --> Elétrica ", vbCritical, "Attention"
    Me.tipo = ""
End If
End Sub

Private Sub UserForm_Initialize()

'função para habilitar botao max min
Call HabilitaBotoes(Me)

'carrega combobox
Call carrega_combobox_aplic_new
Call carrega_combobox_local
item_novo.classe.AddItem "A"
item_novo.classe.AddItem "B"
item_novo.classe.AddItem "C"

item_novo.tipo.AddItem "M"
item_novo.tipo.AddItem "F"
item_novo.tipo.AddItem "E"

End Sub
