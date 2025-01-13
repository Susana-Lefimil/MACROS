Attribute VB_Name = "CopiarArchivosCerrados"
'Copia y pega un rango específico de celdas en otro archivo que esté cerrado'
Sub copiarArchivosCerrado()
Dim archivo1 As Workbook
Dim archivo2 As Workbook

Set archivo1 = ThisWorkbook
Set archivo2 = Workbooks.Open("C:\Users\slefimil\OneDrive - Universidad de los Andes\Escritorio\Macro_importar.xlsx")

archivo1.Activate
Range("A1:H44").Copy

archivo2.Activate
Range("A1").PasteSpecial xlPasteAll

archivo2.Close savechanges:=True

MsgBox "Completado"


End Sub



