Attribute VB_Name = "Eliminar_filas_vacias"
Sub ELIMINAR_FILAS_VACIAS()

'ocultamos el procedimiento
    Application.ScreenUpdating = False
    'suprondremos que vamos a inspeccionar 1.000 filas,
    'en busca de todas las que haya en blanco
    
    fila = Cells(Rows.Count, 1).End(xlUp).Row
    Columns("A").Select
    For i = 1 To fila
    'si la celda está vacía...
        If ActiveCell = "" Or IsNull(ActiveCell) Then
        Hoja1.Rows(i).Delete
        End If
    'pasamos a la siguiente fila
    ActiveCell.Offset(1, 0).Select
    Next
    'mostramos el procedimiento
    Application.ScreenUpdating = True
End Sub
