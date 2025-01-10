Attribute VB_Name = "Separar_hoja_por_columnaFin"
Sub SepararHojaEnFunciónDeUnaColumna()
    'Declarar variables'
    Dim libro As Excel.Worksheet
    Dim nLastRow, nRow, nNextRow As Integer
    Dim colStr As String
    Dim diccionario As Object
    Dim colValues As Variant
    Dim colValue As Variant
    Dim hoja As Excel.Worksheet
    Dim lastLine As Long
    Dim Rng, CI As Range
    
    'ocultamos el procedimiento
    Application.ScreenUpdating = False
    'Change Active sheet name
    ActiveSheet.Name = "Export"
    'Borrar celdas de la columna B vacías
    Sheets("Export").Select
    Range("B2:K15000").Select
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    
    
    'Insertar columna en A'
    lastLine = Cells(Rows.Count, 1).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Activar en columna A la función substitute para eliminar / y :'
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(MID(RC[2],1,23),""/"",""""),"":"",""""),"","", """")"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & lastLine)
    Range("A2:A" & lastLine).Select
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    
             
     'Activar hoja y seleccionar toda la información'
     Set libro = ActiveSheet
     nLastRow = libro.Range("A" & libro.Rows.Count).End(xlUp).Row
     'Crear diccionario'
     Set diccionario = CreateObject("Scripting.Dictionary")
     'Recorrer columna A hasta el último registro, si es nuevo el valor agregar sino no'
     For nRow = 2 To nLastRow
         colStr = libro.Range("A" & nRow).Value
         If diccionario.Exists(colStr) = False Then
            diccionario.Add colStr, 2
         End If
     Next
     'Asignar valor de columnas como las llaves'
     colValues = diccionario.Keys
     
     For i = LBound(colValues) To UBound(colValues)
         colValue = colValues(i)
         Set hoja = Worksheets.Add(After:=Worksheets(Worksheets.Count))
         hoja.Name = colValue
         libro.Rows(1).EntireRow.Copy hoja.Rows(1)
         For nRow = 2 To nLastRow
             If CStr(libro.Range("A" & nRow).Value) = CStr(colValue) Then
                 libro.Rows(nRow).EntireRow.Copy
                 nNextRow = hoja.Range("B" & hoja.Rows.Count).End(xlUp).Row + 1
                 hoja.Range("A" & nNextRow).PasteSpecial xlPasteValuesAndNumberFormats
             End If
         Next
         Columns("A:A").Select
         Selection.Delete Shift:=xlToLeft
         hoja.Columns("A:F").AutoFit
     Next

 
    Sheets("Export").Visible = False
    'mostramos el procedimiento
    Application.ScreenUpdating = True
    MsgBox "Fin procedimiento"
   
End Sub
