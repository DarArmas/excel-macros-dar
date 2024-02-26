Sub AjustarSaldoRestante()

    Dim saldoRestante As Variant
    Dim ultimaFilaNueva As Long

    ultimaFilaNueva = ActiveSheet.Cells(ActiveSheet.Rows.Count, "I").End(xlUp).Row
    saldoRestante = ActiveSheet.Cells(ultimaFilaNueva, "I").Value

    ActiveSheet.Cells(ultimaFilaNueva, "E").Value = saldoRestante

End Sub

Sub EliminarFilasNegativas(ultimaFila As Long, fila As Integer)
    Dim i As Long
    
    For i = ultimaFila To fila + 1 Step -1
       
        If IsNumeric(ActiveSheet.Cells(i, "I").Value) Then
           
            Dim valorCelda As Double
            valorCelda = CDbl(ActiveSheet.Cells(i, "I").Value)
            
            If valorCelda < 0 Then
                ActiveSheet.Rows(i).EntireRow.Delete
            End If
        End If
    Next i
End Sub