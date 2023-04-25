VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Opening_Closing_Price()
    Dim arrDatos() As Variant
    Dim dictDatos As Object
    Dim i As Long, j As Long
    Dim etiqueta As String
    Set dictDatos = CreateObject("Scripting.Dictionary")
    
    'Leer datos de las columnas A a G en un arreglo
    arrDatos = Range("A1:G" & Range("A" & Rows.Count).End(xlUp).Row).Value
    
    'Procesar datos del arreglo para obtener valores únicos de la columna A
    For i = LBound(arrDatos) To UBound(arrDatos)
        If Not dictDatos.Exists(arrDatos(i, 1)) Then
            dictDatos.Add arrDatos(i, 1), 1
        End If
    Next i
    
    'Escribir los valores únicos de la columna A en la columna I
    Range("I1").Resize(dictDatos.Count, 1).Value = WorksheetFunction.Transpose(dictDatos.keys)
    
    'Iterar a través de los valores únicos de la columna A en la columna I
    For i = 1 To dictDatos.Count
        etiqueta = Range("I" & i).Value 'Obtener la etiqueta en la columna A
        
        'Iterar a través de las filas del arreglo para buscar la etiqueta y la fecha correspondiente
        For j = LBound(arrDatos) To UBound(arrDatos)
            'Verificar si la etiqueta y la fecha corresponden al 01 de febrero
            If arrDatos(j, 1) = etiqueta And Right(arrDatos(j, 2), 4) = "0102" Then
                Range("V" & i).Value = arrDatos(j, 3) 'Escribir el valor de la columna C en la columna V de la misma fila
                Exit For 'Salir del bucle si se encuentra la primera coincidencia
            End If
        Next j
        
        'Iterar a través de las filas del arreglo para buscar la etiqueta y la fecha correspondiente
        For j = LBound(arrDatos) To UBound(arrDatos)
            'Verificar si la etiqueta y la fecha corresponden al 31 de diciembre
            If arrDatos(j, 1) = etiqueta And Right(arrDatos(j, 2), 4) = "1231" Then
                Range("W" & i).Value = arrDatos(j, 6) 'Escribir el valor de la columna F en la columna W de la misma fila
                Exit For 'Salir del bucle si se encuentra la primera coincidencia
            End If
        Next j
    Next i
End Sub

