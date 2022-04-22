Sub carrega_combobox()

ThisWorkbook.Activate
On Error Resume Next

Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
'Initialise the collection class variable

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

Do Until rst.EOF = True
    Estoque.pesquisa.AddItem (rst.Fields("DESCRIÇÃO"))
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

End Sub


Sub carrega_combobox_aplic()

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

End Sub

Sub carrega_combobox_cod()

ThisWorkbook.Activate
On Error Resume Next

Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
'Initialise the collection class variable

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset

rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

Do Until rst.EOF = True
    Estoque.pesquisa_cod.AddItem (rst.Fields("CODIGO"))
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub

End Sub

Sub imprimir()


'setup da pagina de impressão
With ThisWorkbook.Worksheets("etiqueta").PageSetup
    .PrintArea = "$E$6:$J$8"
    '.Zoom = 40
    .CenterHorizontally = True
    .CenterVertically = True
    
End With

ThisWorkbook.Worksheets("etiqueta").Activate
Application.Dialogs(xlDialogPrinterSetup).Show
'ThisWorkbook.Worksheets("etiqueta").Range("$E$6:$J$8").PrintOut
Application.Dialogs(xlDialogPageSetup).Show


ThisWorkbook.Worksheets("Menu").Activate

End Sub



