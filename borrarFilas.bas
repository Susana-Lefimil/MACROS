Attribute VB_Name = "borrarFilas"
Sub borrarFilas()
    Dim ultimaFila As Long
    Dim fila As Long
      ultimaFila = Cells(Rows.Count, "E").End(xlUp).Row
      'Hoja1.Cells(Rows.Count, "E").End(xlUp).Row
    
    For fila = 8 To ultimaFila
        If Hoja1.Cells(fila, 5) = Empty Then
            Hoja1.Rows(fila).EntireRow.Delete
        End If
        If ultimaFila = fila Then
            MsgBox "mk"
            Exit Sub
        End If
    Next fila
    
    MsgBox ultimaFila
    
End Sub
