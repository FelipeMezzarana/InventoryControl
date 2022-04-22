Public Function AlmoxarifadoDataBase() As String
'função que facilita o acesso a DataBase MI-Database

Dim wb As Workbook
Dim wsc As Worksheet

Set wb = ThisWorkbook
Set wsc = wb.Worksheets("Auxiliar")
 
Dim pasta As String
Dim folderPath As String
folderPath = Application.ActiveWorkbook.Path

pasta = folderPath & "\" & wsc.Range("j2").Value & "\" & wsc.Range("j3").Value & ".accdb"
AlmoxarifadoDataBase = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pasta
End Function