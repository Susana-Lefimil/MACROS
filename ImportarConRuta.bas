Attribute VB_Name = "ImportarConRuta"
'Genera una ruta para ubicar archivo que se quiere importar'
Sub GENERAR_RUTA()
Dim Ruta As String, EXTENSION2 As String
Ruta = Application.GetOpenFilename
Hoja1.Cells(1, 5) = Ruta

End Sub
'Sirve para importar informacion sin duplicados'
Sub IMPORTAR_DATOS()

Dim LIBRO_QUE_ENVIA As Workbook, LIBRO_QUE_RECIBE As Workbook
Dim HOJA_QUE_ENVIA, HOJA_QUE_RECIBE As Object
Dim RUTAS As String, REPETIDO As String, RECI As String
Dim fila, FINAL, i, fila2, FINAL2, I2 As Long, EXISTE As Long


RUTAS = Hoja1.Cells(1, 5)
Application.ScreenUpdating = False

'libro al que se le extraera la informacion

Set LIBRO_QUE_ENVIA = Workbooks.Open(RUTAS)

Set HOJA_QUE_ENVIA = LIBRO_QUE_ENVIA.Sheets(1)

'libro donde se ingresara la informacion

Set LIBRO_QUE_RECIBE = ThisWorkbook
Set HOJA_QUE_RECIBE = LIBRO_QUE_RECIBE.Sheets(1)


fila = HOJA_QUE_ENVIA.Range("A" & Rows.Count).End(xlUp).Row + 1
FINAL = fila - 1

fila2 = HOJA_QUE_RECIBE.Range("A" & Rows.Count).End(xlUp).Row + 1
FINAL2 = fila2 - 1

For I2 = 2 To FINAL
    EXISTE = 0
        REPETIDO = HOJA_QUE_ENVIA.Cells(I2, 9)
        
        
        For i = 2 To FINAL2
        RECI = HOJA_QUE_RECIBE.Cells(i, 2)
            If LCase(REPETIDO) = LCase(RECI) Then
                EXISTE = 1
                Exit For
            End If
        Next i
        
        If EXISTE = 0 Then
        HOJA_QUE_RECIBE.Cells(fila2, 1) = HOJA_QUE_ENVIA.Cells(I2, 2)
        HOJA_QUE_RECIBE.Cells(fila2, 2) = HOJA_QUE_ENVIA.Cells(I2, 9)
        fila2 = fila2 + 1
        End If
        
        
Next I2

LIBRO_QUE_ENVIA.Close
ThisWorkbook.Save
Application.ScreenUpdating = True
        


End Sub
