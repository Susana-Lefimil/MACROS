Attribute VB_Name = "Actualizacion1"
Sub Actualizar()
    Dim libro As Excel.Worksheet
    Dim nLastRow, nRow, nNextRow As Integer
    Dim colStr As String
    Dim diccionario As Object
    Dim colValues As Variant
    Dim colValue As Variant
    Dim hoja As Excel.Worksheet
    Dim lastLine As Long
    Dim Rng, CI As Range
    Dim lastColumn
    Dim desde() As Variant
    Dim i As Long

    Dim ws As Worksheet
    
    
    'ocultamos el procedimiento
    Application.ScreenUpdating = False
    
    'Change Active sheet name
    ActiveSheet.Name = "Export"
    

    'Funcion para eliminar cualquier hoja menos  export
    Worksheets("Export").Activate
  
    Application.DisplayAlerts = False
  
    For Each ws In Sheets
    
      If ws.Name <> ActiveSheet.Name Then
      
        ws.Delete
        
      End If
      
    Next ws
    
    'Borramos la columna oculta despues de la primera carga.
    lastColumn = 1
    For iCntr = lastColumn To 1 Step -1
        If Columns(iCntr).Hidden = True Then Columns(iCntr).EntireColumn.Delete
    Next
  
    'en la primera carga de informaci�n y despu�s de una actualizacion siempre se comineza de cero en este paso
  
    
    'Borrar celdas de la columna B vac�as
    Sheets("Export").Select
    Range("B2:K15000").Select
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
   
    
    
    'Insertar columna en A'
    lastLine = Cells(Rows.Count, 1).End(xlUp).Row
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Activar en columna A la funci�n substitute para eliminar / y :'
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCAT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(MID(RC[2],1,26),""/"",""""),"":"",""""),"","",""""), "" TAGBORRAR"")"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A" & lastLine)
    Range("A2:A" & lastLine).Select
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Se juntan columnas Hernia, laminectomia y fijaci�n
    
    desde = Array("Hernia N�cleo Pulposo Est TAGBORRAR", "Laminectom�a Descompresiva TAGBORRAR", "Fijaci�n De Columna CS Os TAGBORRAR")
    
    For i = LBound(desde) To UBound(desde)
        Worksheets("Export").Columns("A").Replace _
            What:=desde(i), Replacement:="Hernia Laminectomia Fijacion", _
            SearchOrder:=xlByColumns, MatchCase:=True
    Next
    
    For i = LBound(desde) To UBound(desde)
        Worksheets("Export").Columns("A").Replace _
            What:=" TAGBORRAR", Replacement:="", _
            SearchOrder:=xlByColumns, MatchCase:=True
    Next
    
    
     'Activar hoja y seleccionar toda la informaci�n'
     Set libro = ActiveSheet
     nLastRow = libro.Range("A" & libro.Rows.Count).End(xlUp).Row
     'Crear diccionario'
     Set diccionario = CreateObject("Scripting.Dictionary")
     'Recorrer columna A hasta el �ltimo registro, si es nuevo el valor agregar sino no'
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
    
    
 
    
    'mostramos el procedimiento
    Application.ScreenUpdating = True
    Worksheets("Export").Activate
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Nombre Hoja"
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    
    
    MsgBox "Fin Procedimiento"
  


End Sub
