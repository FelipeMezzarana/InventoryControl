Sub carrega_combobox_cod_editar()

ThisWorkbook.Activate
'On Error GoTo errHandler:

Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
'Initialise the collection class variable

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

Do Until rst.EOF = True
    Edit_Userform.codigo.AddItem (rst.Fields("CODIGO"))
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

errHandler:
'clear memory
Set rst = Nothing
Set cnn = Nothing
MsgBox "Error Request assistance" & err.Number & " (" & err.Description & ") in the procedure carrega_combobox"

End Sub


Sub carrega_combobox_aplic_edit()

ThisWorkbook.Activate
'On Error GoTo errHandler:
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
    For intComboItem = 0 To Estoque.pesquisa_app.ListCount - 1
        If item = Estoque.pesquisa_app.List(intComboItem) Then
            co = 1
        End If
    Next
    If co = 0 Then
        Estoque.pesquisa_app.AddItem (rst.Fields("APLICAÇÃO"))
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

errHandler:
'clear memory
Set rst = Nothing
Set cnn = Nothing
MsgBox "Error Request assistance " & err.Number & " (" & err.Description & ") in the procedure carrega_combobox"

End Sub


Sub carregaitem()
'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim saldo As Integer
Dim saldo_final As Integer

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset


rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[CODIGO] ='" & Edit_Userform.codigo & "'"

Edit_Userform.aplicação = rst.Fields("APLICAÇÃO")
Edit_Userform.descrição = rst.Fields("DESCRIÇÃO")
Edit_Userform.locall = rst.Fields("LOCAL")
Edit_Userform.classe = rst.Fields("CLASSE")
Edit_Userform.tipo = rst.Fields("TIPO")
Edit_Userform.um = rst.Fields("UM")
Edit_Userform.est_min = rst.Fields("ESTOQUE_MINIMO")
Edit_Userform.est_max = rst.Fields("ESTOQUE_MAXIMO")
Edit_Userform.saldo = rst.Fields("SALDO")

rst.Close
Set rst = Nothing
UserForm_Initialize_Exit:
On Error Resume Next
rst.Close
cnn.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub
err:
MsgBox "ERRO - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub



Sub salva_edicao()
'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim saldo As Integer
Dim saldo_final As Integer

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset


rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[CODIGO] ='" & Edit_Userform.codigo & "'"

rst.Fields("APLICAÇÃO") = Edit_Userform.aplicação
rst.Fields("DESCRIÇÃO") = Edit_Userform.descrição
rst.Fields("LOCAL") = Edit_Userform.locall
rst.Fields("CLASSE") = Edit_Userform.classe
rst.Fields("TIPO") = Edit_Userform.tipo
rst.Fields("UM") = Edit_Userform.um
rst.Fields("ESTOQUE_MINIMO") = Edit_Userform.est_min
rst.Fields("ESTOQUE_MAXIMO") = Edit_Userform.est_max
rst.Fields("SALDO") = Edit_Userform.saldo
rst.Update
rst.Close
Set rst = Nothing
UserForm_Initialize_Exit:
On Error Resume Next
rst.Close
cnn.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub
err:
MsgBox "ERRO - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub
