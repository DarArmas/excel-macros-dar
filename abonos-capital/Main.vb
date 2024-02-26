Sub CalcularAbono()
    Dim filaInput As String
    Dim fila AS Integer
    Dim valor As Variant
    Dim entradaValida As Boolean
    Dim limite As Integer
    Dim ultimaFila As Long
    Dim i As Long
    Dim filasAntesPeriodo As Integer
    Dim columnaSaldos As String
    Dim celdaNumPlazos As String

    celdaNumPlazos = "F4"
    columnaSaldos = "I"
    entradaValida = False
    filasAntesPeriodo = 9
    limite = CInt(ActiveSheet.Range(celdaNumPlazos).Value) 
    
    Do While Not entradaValida
        filaInput = InputBox("Selecciona el periodo (1 - " & limite & ") en el que deseas abonar:")
        
        If filaInput = "" Then
            Exit Sub
        End If
        
        
        If IsNumeric(filaInput) Then
            fila = CInt(filaInput)
            If fila >= 1 And fila <= limite Then
                entradaValida = True
            End If
        Else
            MsgBox "Entrada inválida. Selecciona un periodo válido (1 - " & limite & ").", vbExclamation
        End If

    Loop
    
    valor = InputBox("Cantidad a abonar:")
    
    If valor = "" Then
        Exit Sub
    End If

    ActiveSheet.Cells(fila + filasAntesPeriodo, 5).Value = valor
    MsgBox "Abono agregado", vbInformation
    
    ultimaFila = ActiveSheet.Cells(ActiveSheet.Rows.Count, "I").End(xlUp).Row

    Utils.EliminarFilasNegativas ultimaFila, fila

    Utils.AjustarSaldoRestante
End Sub