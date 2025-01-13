Attribute VB_Name = "EliminarFilasVacias"
Sub EliminarFilasVacias()
   Dim fila As Long, i As Long
   
   fila = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
   
   For i = 2 To fila
    If Hoja1.Cells(i, 1) = Empty Then
    Hoja1.Rows(i).Delete
    End If
    
Next i
      
End Sub
