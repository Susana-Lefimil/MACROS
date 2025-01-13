Attribute VB_Name = "copiarArchivosCerrados2"
'Copia y pega celdas con información en otro archivo que esté cerrado"

Sub copiarArchivosCerrado2()
Dim archivo1 As Workbook
Dim archivo2 As Workbook

Set archivo1 = ThisWorkbook
Set archivo2 = Workbooks.Open("C:\Users\slefimil\OneDrive - Universidad de los Andes\Escritorio\Macro_importar.xlsx")

archivo1.Activate
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

archivo2.Activate
Range("A1").PasteSpecial xlPasteAll

archivo2.Close savechanges:=True

MsgBox "Completado"


End Sub


