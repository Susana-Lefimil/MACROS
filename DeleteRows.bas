Attribute VB_Name = "DeleteRows"
Sub DeleteRows()
   Sheets("Hoja1").Select
   Range("B2:K15000").Select
   Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
End Sub
