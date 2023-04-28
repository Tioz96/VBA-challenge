Sub Ticker()
    Dim arrDatos() As Variant
    Dim dictDatos As Object
    Set dictDatos = CreateObject("Scripting.Dictionary")
    
    
    'Leer datos de las columnas A a F en un arreglo
    arrDatos = Range("A1:F" & Range("A" & Rows.Count).End(xlUp).Row).Value
    
    'Procesar datos del arreglo
    For i = LBound(arrDatos) To UBound(arrDatos)
        'Agregar el valor de la columna A al diccionario si no existe ya
        If Not dictDatos.Exists(arrDatos(i, 1)) Then
            dictDatos.Add arrDatos(i, 1), 1
        End If
    Next i
    
    'Escribir los valores únicos de la columna A en la columna I
    Range("I1").Resize(dictDatos.Count, 1).Value = WorksheetFunction.Transpose(dictDatos.keys)
    Range("I1").Value = "Ticker"
ActiveSheet.UsedRange.Columns.AutoFit
    
End Sub

Sub Opening_Closing_Price()

    Dim arrDatos() As Variant
    Dim dictDatos As Object
    Set dictDatos = CreateObject("Scripting.Dictionary")
    
    'Leer datos de las columnas A a F en un arreglo
    arrDatos = Range("A1:G" & Range("A" & Rows.Count).End(xlUp).Row).Value
    
    'Procesar datos del arreglo
    For i = LBound(arrDatos) To UBound(arrDatos)
        'Agregar el valor de la columna A al diccionario si no existe ya
        If Not dictDatos.Exists(arrDatos(i, 1)) Then
            dictDatos.Add arrDatos(i, 1), 1
        End If
    Next i
    
    'Escribir los valores únicos de la columna A en la columna I
    For i = 1 To dictDatos.Count
        'Obtener la etiqueta en la columna A
        etiqueta = Range("I" & i).Value
        
        'Iterar a través de las filas del arreglo para buscar la etiqueta y la fecha correspondiente
        For j = LBound(arrDatos) To UBound(arrDatos)
            'Verificar si la etiqueta y la fecha corresponden
            If arrDatos(j, 1) = etiqueta And Right(arrDatos(j, 2), 4) = "0102" Then
                'Escribir el valor de la columna C en la columna V de la misma fila
                Range("V" & i).Value = arrDatos(j, 3)
                Exit For 'Salir del bucle si se encuentra la primera coincidencia
            End If
        Next j
        
        'Iterar a través de las filas del arreglo para buscar la etiqueta y la fecha correspondiente
        For j = LBound(arrDatos) To UBound(arrDatos)
            'Verificar si la etiqueta y la fecha corresponden
            If arrDatos(j, 1) = etiqueta And Right(arrDatos(j, 2), 4) = "1231" Then
                'Escribir el valor de la columna F en la columna W de la misma fila
                Range("W" & i).Value = arrDatos(j, 6)
                Exit For 'Salir del bucle si se encuentra la primera coincidencia
            End If
        Next j
    Next i
    
ActiveSheet.UsedRange.Columns.AutoFit

End Sub

Sub Calculate_YearlyChange_Percetnchange()
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "V").End(xlUp).Row
    
    Range("J1").Value = "YearlyChange"
Range("K1").Value = "PercentChange"

    
    For i = 2 To ultimaFila 'comenzando en la fila 2, asumiendo que la fila 1 tiene encabezados de columna
        'Obtener los valores de apertura y cierre
        openingprice = Cells(i, "V").Value
        closingprice = Cells(i, "W").Value
        
        'Calcular Yearly Change y Percent Change
        yearlychange = closingprice - openingprice
        percentchange = (closingprice / openingprice) - 1
        
        'Escribir los resultados en las columnas J y K
        Cells(i, "J").Value = yearlychange
        Cells(i, "K").Value = percentchange
    Next i

Columns("K:K").NumberFormat = "0.00%"

Dim rango As Range
    Dim celda As Range
    
    Set rango = Range("J2:J" & Cells(Rows.Count, "J").End(xlUp).Row) ' Definir rango de celdas en la columna J
    
    For Each celda In rango
        If celda.Value < 0 Then
            celda.Interior.Color = vbRed ' Cambiar el color de fondo a rojo si el valor es negativo
        ElseIf celda.Value > 0 Then
            celda.Interior.Color = vbGreen ' Cambiar el color de fondo a verde si el valor es positivo
        End If
    Next celda

Columns("V:W").Clear

ActiveSheet.UsedRange.Columns.AutoFit

End Sub

Sub Total_Stock_Volume()

    Dim lastRowA As Long, lastRowI As Long, i As Long, j As Long
    Dim sumRange() As Variant, compareRange() As Variant, uniqueValues() As Variant, sumValues() As Variant
    
    Range("L1").Value = "Stock Total Volume"
        
    lastRowA = Cells(Rows.Count, "A").End(xlUp).Row
    lastRowI = Cells(Rows.Count, "I").End(xlUp).Row
   
    compareRange = Range("A1:A" & lastRowA).Value
    sumRange = Range("G1:G" & lastRowA).Value
    uniqueValues = WorksheetFunction.Transpose(WorksheetFunction.Unique(Range("I2:I" & lastRowI)))
    ReDim sumValues(1 To UBound(uniqueValues))
    
    For i = 1 To UBound(uniqueValues)
        For j = 1 To lastRowA
            If compareRange(j, 1) = uniqueValues(i) Then
                sumValues(i) = sumValues(i) + sumRange(j, 1)
            End If
        Next j
    Next i

    Range("L2:L" & UBound(uniqueValues) + 1).Value = WorksheetFunction.Transpose(sumValues)

ActiveSheet.UsedRange.Columns.AutoFit

End Sub

Sub PorcentajeMinMax()
    Dim lastRow As Long
    Dim minValue As Double
    Dim maxValue As Double
    Dim maxValueL As Double
    Dim maxRow As Long

Range("O2").Value = "Gratest % Increased"
Range("O3").Value = "Greatest % Decreased"
Range("O4").Value = "Gratest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

    lastRow = Cells(Rows.Count, "K").End(xlUp).Row
    minValue = WorksheetFunction.Min(Range("K1:K" & lastRow))
    maxValue = WorksheetFunction.Max(Range("K1:K" & lastRow))
    maxValueL = WorksheetFunction.Max(Range("L1:L" & lastRow))
    maxRow = WorksheetFunction.Match(maxValue, Range("K1:K" & lastRow), 0)

    Range("Q2").Value = maxValue
    Range("Q4").Value = Cells(maxRow, "L").Value
    Range("Q3").Value = minValue
    Range("P2:P4").Value = ""
    Range("P2").Value = Cells(maxRow, "I").Value
    Range("P3").Value = Cells(WorksheetFunction.Match(minValue, Range("K1:K" & lastRow), 0), "I").Value
    Range("P4").Value = Cells(WorksheetFunction.Match(maxValueL, Range("L1:L" & lastRow), 0), "I").Value
    Range("Q2:Q3").NumberFormat = "0.00%"

ActiveSheet.UsedRange.Columns.AutoFit
    
End Sub

