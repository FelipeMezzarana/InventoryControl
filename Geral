Sub VOLTAR()

ThisWorkbook.Worksheets("Menu").Select

End Sub
Sub show_estoque()

Estoque.Show vbModeless

End Sub

Sub show_editar()

Edição.Show vbModeless

End Sub


Sub show_novo()

item_novo.Show vbModeless

End Sub

Sub iniciar()
'Remove linhas, colunas e barras de rolagem
ThisWorkbook.Activate
Dim wsSheet As Worksheet

Worksheets("Menu").Unprotect "al"
    Application.ScreenUpdating = False

    For Each wsSheet In ThisWorkbook.Worksheets

        If Not wsSheet.Name = "Blank" Then

            wsSheet.Activate

            With ActiveWindow

                .DisplayHeadings = False

                .DisplayWorkbookTabs = False

                .DisplayHorizontalScrollBar = False

                .DisplayVerticalScrollBar = False

            End With

        End If

    Next wsSheet

        
    'Fullscreen
    Application.DisplayFullScreen = True
    Worksheets("Menu").Select
    
    
'seleciona e protege menu

Worksheets(1).Select

With Worksheets(1)

.EnableSelection = xlNoSelection

.Protect Contents:=True, UserInterfaceOnly:=True

End With

'trava tela(limita rolatem)
Worksheets(1).ScrollArea = "$A$1:$a$1"

'esconde formula bar
Application.DisplayFormulaBar = False

Worksheets(1).Select
ActiveWindow.Zoom = 78

Worksheets("Menu").Protect "al"
 Application.ScreenUpdating = True
End Sub

Sub desproteger()
'Remove linhas, colunas e barras de rolagem
Dim wsSheet As Worksheet
    Application.ScreenUpdating = False
    
Worksheets("Auxiliar").Unprotect "al"
Worksheets("Menu").Unprotect "al"

    For Each wsSheet In ThisWorkbook.Worksheets

        If Not wsSheet.Name = "Blank" Then

            wsSheet.Activate

            With ActiveWindow

                .DisplayHeadings = True

                .DisplayWorkbookTabs = True

                .DisplayHorizontalScrollBar = True

                .DisplayVerticalScrollBar = True

            End With

        End If

    Next wsSheet

        
    'Fullscreen
    Application.DisplayFullScreen = False
    Worksheets("Menu").Select
    
    
'seleciona e protege menu

Worksheets(1).Select

With Worksheets(1)

.EnableSelection = xlNoRestrictions

.Protect Contents:=False, UserInterfaceOnly:=False


End With

Worksheets("Auxiliar").Select
With Worksheets("Auxiliar")

.EnableSelection = xlNoRestrictions

.Protect Contents:=False, UserInterfaceOnly:=False

End With
Worksheets("Auxiliar").Select


'trava tela(limita rolatem)
Worksheets(1).ScrollArea = ""

'esconde formula bar
Application.DisplayFormulaBar = True

Worksheets(1).Select
Worksheets("Auxiliar").Unprotect "al"
Worksheets("Menu").Unprotect "al"

 Application.ScreenUpdating = True
End Sub
Sub destravar_e_cadastrar()

Dim wb As Workbook
Dim ws As Worksheet
Dim li As Integer
Dim cont As Integer
Dim cn As String
Dim resposta As Variant

Set wb = ThisWorkbook
Set ws = wb.Worksheets("Auxiliar")

li = ws.Range("a10000").End(xlUp).Row
cont = 0

resposta = InputBox("Insira seu E-Number", "Acesso restrito")
cn = resposta

'cria marcador para checar cadastro, cont = 1 é cadastrado, cont = 0 não é
Do While li > 0
    If ws.Cells(li, 1).Text = cn Then
    cont = 1
    End If
li = li - 1
Loop

If cont = 1 Then
    Call desproteger
    ws.Select
Else
    MsgBox "Password not registered" & vbCrLf & vbCrLf & "Access denied", vbCritical
End If


End Sub


Sub zoon_plus()

'Worksheets("Menu").Activate
ActiveWindow.Zoom = ActiveWindow.Zoom + 3

End Sub

Sub zoon_minus()

'Worksheets("Menu").Activate
ActiveWindow.Zoom = ActiveWindow.Zoom - 3
End Sub