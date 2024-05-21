Sub BUSCAR()
With Hoja2

' recuperando codigo de articulo
'variable que almacena el codigo del articulo

Dim codigo As Long

codigo = Hoja1.Cells(2, 3)
'MsgBox (codigo)

'buscando y pintando lineas

Dim fila, filamax As Long

'calculando maximo de filas grabadas
filamax = .UsedRange.Rows.Count
'MsgBox (filamax)

For fila = 3 To filamax

    If (Application.WorksheetFunction.CountIf(.Columns(1), codigo)) >= 1 Then
    
        Hoja1.Cells(4, 3) = Application.WorksheetFunction.VLookup(codigo, .Range("A:F"), 2, 0)
        
        Hoja1.Cells(6, 3) = Application.WorksheetFunction.VLookup(codigo, .Range("A:F"), 3, 0)
        
        Hoja1.Cells(8, 3) = Application.WorksheetFunction.VLookup(codigo, .Range("A:F"), 4, 0)
        
        Hoja1.Cells(10, 3) = Application.WorksheetFunction.VLookup(codigo, .Range("A:F"), 5, 0)
        
        Hoja1.Cells(12, 3) = Application.WorksheetFunction.VLookup(codigo, .Range("A:F"), 6, 0)
        
        
    End If
    
Next fila

'posicionar en celda codigo
Hoja1.Cells(2, 3).Select


End With

End Sub
