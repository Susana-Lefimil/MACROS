Attribute VB_Name = "RangosDinamicos"
'Sirve para determinar que celdas tienen información'

Sub copiarRangoDinamico()
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Hoja2").Activate
Range("A1").PasteSpecial xlPasteAll
End Sub




Sub copiarRangoHoja2()
Range("A1:H48").Copy
Sheets("Hoja2").Activate
Range("A1").PasteSpecial xlPasteAll
End Sub
