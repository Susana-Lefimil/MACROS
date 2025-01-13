Attribute VB_Name = "Separar_hojas"
Sub separar()
Dim ancho As Integer
Dim nfilas As Integer
Dim nintervalos As Integer
Dim ncolumnas As Integer
Dim columna_inicio As Integer
Dim columna_fin As Integer


ancho = 50
nfilas = Range("tabla1").Roes.Count
nintervalos = Application.WorksheetFunction.RoundDown(nfilas / ancho, 0)
inicio = Range("Tabla1[#Headers]").Row
ncolumnas = Application.WorksheetFunction.CountA(Range("Tabla1[#Headers]"))
columna_inicio = Range("Tabla1[[#Headers],[id]").Column
column_fin = Range("Tabla1[[#Headers],[palabra]]").Column

For i = 1 To nintervalos
    Worksheets.Add.Name = "Lista" & i
    Range("Tabla1[#Headers]").Copy Destination:=Sheets("Lista" & i).Range("A1")
    Sheets("hoja1").Select
    Range(Cells(ancho * (i - 1) + (inicio + 1), columna_inicio), Cells(ancho * i + inicio, columna_fin)).Copy Destination:=Sheets("Lista" & i).Range("A2")
Next i

Worksheets.Add.Name = "Lista" & i
Range("Tabla1[#Headers]").Copy Destination:=Sheets("Lista" & i).Range("A1")
Sheets("hoja1").Select
Range(Cells(ancho * (i - 1) + (inicio + 1), columna_inicio), Cells(nfilas + inicio, columna_fin)).Copy Destination:=Sheets("Lista" & i).Range("A2")

    End Sub
