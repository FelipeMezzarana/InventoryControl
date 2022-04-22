
Sub checagem()

For Each ctl In item_novo.Controls
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

For Each ctl In item_novo.Controls
    If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
        ctl.Value = ""
    End If
    Next ctl
    
item_novo.Image47.Picture = Imagens.errado.Picture

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
    For intComboItem = 0 To item_novo.aplicacao.ListCount - 1
        If item = item_novo.aplicacao.List(intComboItem) Then
            co = 1
        End If
    Next
    If co = 0 Then
        item_novo.aplicacao.AddItem (rst.Fields("APLICAÇÃO"))
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
    For intComboItem = 0 To item_novo.locall.ListCount - 1
        If item = item_novo.locall.List(intComboItem) Then
            co = 1
        End If
    Next
    If co = 0 Then
        item_novo.locall.AddItem (rst.Fields("LOCAL"))
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

If item_novo.classe = "" Or item_novo.tipo = "" Then
    MsgBox "Select class and type to generate code", vbCritical, "Attention"
    Exit Sub
End If

cod = "AM." & item_novo.classe & "." & item_novo.tipo & "."

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
rst.Filter = "CLASSE ='" & item_novo.classe & "' AND TIPO ='" & item_novo.tipo & "'"

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

item_novo.codigo = cod

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
rst.Fields("CODIGO") = item_novo.codigo
rst.Fields("APLICAÇÃO") = item_novo.aplicacao
rst.Fields("DESCRIÇÃO") = item_novo.descrição
rst.Fields("LOCAL") = item_novo.locall
rst.Fields("CLASSE") = item_novo.classe
rst.Fields("TIPO") = item_novo.tipo
rst.Fields("UM") = item_novo.um
rst.Fields("ESTOQUE_MINIMO") = item_novo.est_min
rst.Fields("ESTOQUE_MAXIMO") = item_novo.est_max
rst.Fields("SALDO") = item_novo.saldo
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

