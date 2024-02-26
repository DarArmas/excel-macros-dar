Sub CalcularAbono()
    Dim filaInput As String
    Dim fila AS Integer
    Dim valor As Variant
    Dim periodoValido As Boolean
    Dim abonoValido As Boolean
    Dim limite As Integer
    Dim ultimaFila As Long
    Dim i As Long
    Dim filasAntesPeriodo As Integer
    Dim columnaSaldos As String
    Dim columnaExtras As String
    Dim celdaNumPlazos As String
    Dim saldoCorriente As Double

    celdaNumPlazos = "F4"
    columnaSaldos = "I"
    columnaExtras = "E"
    periodoValido = False
    abonoValido = False
    filasAntesPeriodo = 9
    limite = CInt(ActiveSheet.Range(celdaNumPlazos).Value)
    
    Do While Not periodoValido
        filaInput = InputBox("Selecciona el periodo (1 - " & limite & ") en el que deseas abonar:")
        
        If filaInput = "" Then
            Exit Sub
        End If
        
        
        If IsNumeric(filaInput) Then
            fila = CInt(filaInput)
            If fila >= 1 And fila <= limite Then
                periodoValido = True
            End If
        Else
            MsgBox "Entrada inválida. Selecciona un periodo válido (1 - " & limite & ").", vbExclamation
        End If

    Loop

    saldoCorriente = CDbl(ActiveSheet.Cells((fila + filasAntesPeriodo) - 1, columnaSaldos).Value)
    
    Do While Not abonoValido
        valor = InputBox("Cantidad a abonar:")
        
        If valor = "" Then
            Exit Sub
        End If
        
        If CDbl(valor) <= saldoCorriente Then
            abonoValido = true
            ActiveSheet.Cells(fila + filasAntesPeriodo, "E").Value = valor
            MsgBox "Abono agregado", vbInformation
        Else
            MsgBox "El bono sobrepasa lo que debes", vbExclamation
        End If
    Loop

    
    ultimaFila = ActiveSheet.Cells(ActiveSheet.Rows.Count, "I").End(xlUp).Row

    Utils.EliminarFilasNegativas ultimaFila, fila

    Utils.AjustarSaldoRestante

End Sub


Sub DeshacerCalcularAbono()
    Dim nombreHoja As String
    Dim nombreTemporal As String
    
    nombreHoja = ActiveSheet.Name
    
    If Not SheetExists("Macro backup") Then
        MsgBox "La hoja 'macro backup' no existe.", vbExclamation
        Exit Sub
    End If
    
    nombreTemporal = nombreHoja & "_Temp"
    
    Sheets("macro backup").Copy Before:=Sheets(1)
    Sheets(1).Name = nombreTemporal
    
    Application.DisplayAlerts = False 
    Sheets(nombreHoja).Delete
    Application.DisplayAlerts = True 
    
    Sheets(nombreTemporal).Name = nombreHoja
End Sub

Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function