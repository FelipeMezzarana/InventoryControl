
'Função que retornará o nome da classe e o nome do UserForm
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Função que recupera as informações sobre o nome da classe e o estilo da janela do UserForm
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'Função que altera o estilo da janela do UserForm
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Sub que irá obter o nome do UserForm (ObjForm)
Sub HabilitaBotoes(objForm As Object)

    'Código que atribui os botões minimizar e maximizar e possibilita redimensionar o UserForm
    SetWindowLong FindWindow("ThunderDFrame", objForm.Caption), -16, GetWindowLong(FindWindow("ThunderDFrame", objForm.Caption), -16) Or &H10000 Or &H20000 Or &H40000

End Sub
