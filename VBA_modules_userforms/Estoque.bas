
Private Sub CommandButton11_Click()

On Error GoTo err
Dim desc As String
desc = Me.pesquisa.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String
Dim sp_1 As String
Dim sp_2 As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset



rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    With ListBox1
        .AddItem
        .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
        .List(LinhaListbox, 1) = " | "
        .List(LinhaListbox, 2) = rst.Fields("CODIGO")
        .List(LinhaListbox, 3) = " | "
        .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
        .List(LinhaListbox, 5) = " | "
        .List(LinhaListbox, 6) = rst.Fields("LOCAL")
        .List(LinhaListbox, 7) = " | "
        .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
        
        LinhaListbox = LinhaListbox + 1
    End With
rst.MoveNext
Loop

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
MsgBox "ERROR - request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit


End Sub

Private Sub CommandButton12_Click()

On Error GoTo err
Dim desc As String
desc = Me.pesquisa.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String
Dim sp_1 As String
Dim sp_2 As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset



rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    If CDbl(rst.Fields("SALDO")) = 0 Then
        If rst.Fields("CLASSE") = "A" Or rst.Fields("CLASSE") = "B" Then
            With ListBox1
                .AddItem
                .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
                .List(LinhaListbox, 1) = " | "
                .List(LinhaListbox, 2) = rst.Fields("CODIGO")
                .List(LinhaListbox, 3) = " | "
                .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
                .List(LinhaListbox, 5) = " | "
                .List(LinhaListbox, 6) = rst.Fields("LOCAL")
                .List(LinhaListbox, 7) = " | "
                .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
                LinhaListbox = LinhaListbox + 1
            End With
        End If
    End If
    rst.MoveNext
Loop

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
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit

End Sub

Private Sub CommandButton13_Click()
On Error GoTo err
Dim desc As String
desc = Me.pesquisa.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String
Dim sp_1 As String
Dim sp_2 As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset



rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable

If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    If CDbl(rst.Fields("SALDO")) = 0 Then
        With ListBox1
            .AddItem
            .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
            .List(LinhaListbox, 1) = " | "
            .List(LinhaListbox, 2) = rst.Fields("CODIGO")
            .List(LinhaListbox, 3) = " | "
            .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
            .List(LinhaListbox, 5) = " | "
            .List(LinhaListbox, 6) = rst.Fields("LOCAL")
            .List(LinhaListbox, 7) = " | "
            .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
            LinhaListbox = LinhaListbox + 1
        End With
    End If
    rst.MoveNext
Loop

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
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit

End Sub

Private Sub CommandButton14_Click()
On Error GoTo err
Dim cod As String
cod = Me.pesquisa_cod.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;30;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset


rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[CODIGO] ='" & cod & "'"


If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    With ListBox1
        .AddItem
        .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
        .List(LinhaListbox, 1) = " | "
        .List(LinhaListbox, 2) = rst.Fields("CODIGO")
        .List(LinhaListbox, 3) = " | "
        .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
        .List(LinhaListbox, 5) = " | "
        .List(LinhaListbox, 6) = rst.Fields("LOCAL")
        .List(LinhaListbox, 7) = " | "
        .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
        LinhaListbox = LinhaListbox + 1
    End With
rst.MoveNext
Loop

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
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub

Private Sub CommandButton15_Click()

On Error GoTo err

Dim it As String
Dim nome As String

Dim i As Integer
    If ListBox1.ListIndex = -1 Then
        it = "Nothing"
        MsgBox "Select an item", vbCritical, "Atenção"
        Exit Sub
    Else
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
              it = ListBox1.List(i, 2)
              nome = ListBox1.List(i, 0)
            End If
        Next i
    End If

Dim myValue As Variant

myValue = InputBox("Item: " & nome & vbCrLf & vbCrLf & "Código: " & it & vbCrLf & vbCrLf & "Digite a quantidade do item a ser ADICIONADA no estoque", "ENTRADA")

If Not IsNumeric(myValue) Then
    MsgBox "REGISTRATION NOT COMPLETED" & vbCrLf & vbCrLf & "Please enter a numeric value", vbCritical, "Attention"
    Exit Sub
End If


'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim saldo As Integer
Dim saldo_final As Integer
Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset


rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[CODIGO] ='" & it & "'"

saldo = CDbl(rst.Fields("SALDO"))
rst.Fields("SALDO") = saldo + myValue
saldo_final = rst.Fields("SALDO")

rst.Update
rst.Close
Set rst = Nothing
MsgBox "Registration successful" & vbCrLf & vbCrLf & "Final balance " & it & ": " & saldo_final, vbInformation, "Registration"
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


Private Sub CommandButton16_Click()
On Error GoTo err

Dim it As String
Dim nome As String

Dim i As Integer
    If ListBox1.ListIndex = -1 Then
        it = "Nothing"
        MsgBox "Select an item", vbCritical, "Attention"
        Exit Sub
    Else
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
              it = ListBox1.List(i, 2)
              nome = ListBox1.List(i, 0)
            End If
        Next i
    End If

Dim myValue As Variant

myValue = InputBox("Item: " & nome & vbCrLf & vbCrLf & "Código: " & it & vbCrLf & vbCrLf & "Digite a quantidade do item a ser RETIRADA no estoque", "ENTRADA")

If Not IsNumeric(myValue) Then
    MsgBox "REGISTRATION NOT COMPLETED" & vbCrLf & vbCrLf & "Please enter a numeric value", vbCritical, "Attention"
    Exit Sub
End If


'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim saldo As Integer
Dim saldo_final As Integer

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset


rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[CODIGO] ='" & it & "'"

saldo = CDbl(rst.Fields("SALDO"))
rst.Fields("SALDO") = saldo - myValue
saldo_final = rst.Fields("SALDO")

rst.Update
rst.Close
Set rst = Nothing
MsgBox "Registration successful" & vbCrLf & vbCrLf & "Final balance " & it & ": " & saldo_final, vbInformation, "Registration"
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

Private Sub CommandButton17_Click()
On Error GoTo err

Dim it As String
Dim nome As String

Dim i As Integer
    If ListBox1.ListIndex = -1 Then
        it = "Nothing"
        MsgBox "Select an item", vbCritical, "Attention"
        Exit Sub
    Else
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
              it = ListBox1.List(i, 2)
              nome = ListBox1.List(i, 0)
            End If
        Next i
    End If

Edição.Show vbModeless
Edição.codigo = it


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

Private Sub CommandButton18_Click()
'On Error GoTo err

Dim codigo As String
Dim descricao As String
Dim aplicacao As String
Dim codigo_bar As String
Dim locall As String


Dim i As Integer
    If ListBox1.ListIndex = -1 Then
        it = "Nothing"
        MsgBox "Select an item", vbCritical, "Attention"
        Exit Sub
    Else
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
              codigo = ListBox1.List(i, 2)
              descricao = ListBox1.List(i, 0)
              aplicacao = ListBox1.List(i, 4)
              locall = ListBox1.List(i, 6)
            End If
        Next i
    End If

codigo_bar = "*" & codigo & "*"

'preenche folha de impressão

Dim wb As Workbook
Dim ws As Worksheet

Set wb = ThisWorkbook
Set ws = wb.Worksheets("etiqueta")

ws.Range("E6") = descricao
ws.Range("E7") = locall
ws.Range("F7") = codigo_bar
ws.Range("E8") = aplicacao
ws.Range("H8") = codigo

'Call imprimir
Unload Estoque
ThisWorkbook.Worksheets("etiqueta").Select
Exit Sub
err:
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
End Sub

Private Sub CommandButton6_Click()
Consulta_OT.Zoom = Consulta_OT.Zoom + 5
End Sub

Private Sub CommandButton7_Click()
Consulta_OT.Zoom = Consulta_OT.Zoom - 5
End Sub

Private Sub CommandButton8_Click()
On Error GoTo err
Dim desc As String
desc = Me.pesquisa.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;30;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset



rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[DESCRIÇÃO] ='" & desc & "'"


If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    With ListBox1
        .AddItem
        .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
        .List(LinhaListbox, 1) = " | "
        .List(LinhaListbox, 2) = rst.Fields("CODIGO")
        .List(LinhaListbox, 3) = " | "
        .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
        .List(LinhaListbox, 5) = " | "
        .List(LinhaListbox, 6) = rst.Fields("LOCAL")
        .List(LinhaListbox, 7) = " | "
        .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
        LinhaListbox = LinhaListbox + 1
    End With
rst.MoveNext
Loop

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
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub




Private Sub CommandButton9_Click()
On Error GoTo err
Dim desc As String
desc = Me.pesquisa_app.Text
'zera e configura listbox
Dim LinhaListbox As Integer
    LinhaListbox = 0
    ListBox1.Clear
    Me.ListBox1.ColumnCount = 9
    ListBox1.ColumnWidths = "250;20;100;20;150;20;100;20;30;40"

'cria e abre conexão
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim s As String
Dim n As String

Set cnn = New ADODB.Connection
cnn.Open AlmoxarifadoDataBase()
Set rst = New ADODB.Recordset



rst.Open Source:="Estoque", ActiveConnection:=cnn, CursorType:=adOpenStatic, LockType:=adLockOptimistic, Options:=adCmdTable
rst.Filter = "[APLICAÇÃO] ='" & desc & "'"


If rst.EOF = True Then
    MsgBox "No record found", vbInformation, "OT"
    Exit Sub
Else
    rst.MoveFirst
End If

Do Until rst.EOF
    With ListBox1
        .AddItem
        .List(LinhaListbox, 0) = rst.Fields("DESCRIÇÃO")
        .List(LinhaListbox, 1) = " | "
        .List(LinhaListbox, 2) = rst.Fields("CODIGO")
        .List(LinhaListbox, 3) = " | "
        .List(LinhaListbox, 4) = rst.Fields("APLICAÇÃO")
        .List(LinhaListbox, 5) = " | "
        .List(LinhaListbox, 6) = rst.Fields("LOCAL")
        .List(LinhaListbox, 7) = " | "
        .List(LinhaListbox, 8) = rst.Fields("ESTOQUE_MINIMO") & "  /  " & rst.Fields("ESTOQUE_MAXIMO") & "  /  " & rst.Fields("SALDO")
        LinhaListbox = LinhaListbox + 1
    End With
rst.MoveNext
Loop

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
MsgBox "ERROR - Request assistance" & err.Number & vbCrLf & err.Description, vbCritical, "Error!"
Resume UserForm_Initialize_Exit
End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
'maximizar userform
Application.WindowState = xlMaximized
Me.Height = Application.Height
Me.Width = Application.Width
Me.Left = Application.Left
Me.Top = Application.Top
Me.StartUpPosition = 3

'função para habilitar botao max min
Call HabilitaBotoes(Me)
'subs para preencher combobox
Call carrega_combobox
Call carrega_combobox_aplic
Call carrega_combobox_cod

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
ThisWorkbook.Worksheets("Menu").Select
End Sub