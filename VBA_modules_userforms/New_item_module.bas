
Sub checagem()

For Each ctl In new_item_Userform.Controls
    Select Case TypeName(ctl)
        Case "TextBox", "combobox"
            If ctl.Text = "" Then
                MsgBox "Record NOT SAVED" & vbCrLf & vbCrLf & "All text boxes must be filled in", vbCritical, "Attention"
                Exit Sub
            End If
    End Select
    Next ctl
End Sub


Sub apagar_item_novo()

For Each ctl In new_item_Userform.Controls
    If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
        ctl.Value = ""
    End If
    Next ctl
    
new_item_Userform.Image47.Picture = Images_Userform.errado.Picture

End Sub


Sub carrega_combobox_aplic_new()

ThisWorkbook.Activate
On Error Resume Next
Dim intComboItem As Integer
Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
Dim item As String
Dim co As Integer
'Initialise the collection class variable

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

    
Do Until rst.EOF = True
    item = rst.Fields("APLICAÇÃO")
    co = 0
    For intComboItem = 0 To new_item_Userform.aplicacao.ListCount - 1
        If item = new_item_Userform.aplicacao.List(intComboItem) Then
            co = 1
        End If
    Next
    If co = 0 Then
        new_item_Userform.aplicacao.AddItem (rst.Fields("APLICAÇÃO"))
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

End Sub

Sub carrega_combobox_local()

ThisWorkbook.Activate
On Error Resume Next
Dim intComboItem As Integer
Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
Dim item As String
Dim co As Integer
'Initialise the collection class variable

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

    
Do Until rst.EOF = True
    item = rst.Fields("LOCAL")
    co = 0
    For intComboItem = 0 To new_item_Userform.locall.ListCount - 1
        If item = new_item_Userform.locall.List(intComboItem) Then
            co = 1
        End If
    Next
    If co = 0 Then
        new_item_Userform.locall.AddItem (rst.Fields("LOCAL"))
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

End Sub


Sub gerar_codigo()

Dim cod As String

If new_item_Userform.classe = "" Or new_item_Userform.tipo = "" Then
    MsgBox "Select class and type to generate code", vbCritical, "Attention"
    Exit Sub
End If

cod = "AM." & new_item_Userform.classe & "." & new_item_Userform.tipo & "."

'busca numero
'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String
Dim counter As Integer
Dim digitos As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "CLASSE ='" & new_item_Userform.classe & "' AND TIPO ='" & new_item_Userform.tipo & "'"

counter = 1
Do Until rst.EOF = True
    counter = counter + 1
    rst.MoveNext
Loop
rst.Close

counter = counter + 100

If counter <= 9 Then
    digitos = "0000" & counter
End If

If counter <= 99 And counter > 9 Then
    digitos = "000" & counter
End If

If counter <= 999 And counter > 99 Then
    digitos = "00" & counter
End If

If counter <= 9990 And counter > 999 Then
    digitos = "0" & counter
End If


cod = cod & digitos

new_item_Userform.codigo = cod

End Sub


Sub novo_cadastro()

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim saldo As Integer
Dim saldo_final As Integer

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic, Options:=adCmdTable

rst.AddNew
rst.Fields("CODIGO") = new_item_Userform.codigo
rst.Fields("APLICAÇÃO") = new_item_Userform.aplicacao
rst.Fields("DESCRIÇÃO") = new_item_Userform.descrição
rst.Fields("LOCAL") = new_item_Userform.locall
rst.Fields("CLASSE") = new_item_Userform.classe
rst.Fields("TIPO") = new_item_Userform.tipo
rst.Fields("UM") = new_item_Userform.um
rst.Fields("ESTOQUE_MINIMO") = new_item_Userform.est_min
rst.Fields("ESTOQUE_MAXIMO") = new_item_Userform.est_max
rst.Fields("SALDO") = new_item_Userform.saldo
rst.Update

UserForm_Initialize_Exit:
On Error Resume Next
rst.Close
cnn.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub
err:
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub

