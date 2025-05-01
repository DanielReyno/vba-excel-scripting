Attribute VB_Name = "Module1"
Sub custom_macro()
    
    Dim hoja As Worksheet
    Set hoja = ThisWorkbook.Sheets(1)
    
    
    Dim listado As Variant
    
    listado = Array("Daniel", "Jose", "Mario", "Yanet")
    
    
    listado_length = UBound(listado) - LBound(listado) + 1
    
    
    Call demoFunction(hoja)
    
    'Cells(10, 10).value = "Something"
    'Range("A1", "G30").Activate
    
    
    'Dim numeroDeElementos As Integer
    'numeroDeElementos = longitudDeUnaColumna(hoja)
    
    'MsgBox numeroDeElementos
    
    
End Sub

Public Function llenarCeldasATravesArray(listado As Variant)
    Dim control1 As Integer
    For control1 = 0 To listado_length - 1
        Cells(control1 + 1, 1).value = listado(control1)
    Next
    
    Exit Function
End Function


Public Function mostrarSecuenciaDeNumeros()
    Dim pivot, control1 As Integer
    pivot = InputBox("Digite la cantidad de numeros que quieras contar")
    For control1 = 1 To pivot
        MsgBox "Secuencia: " & control1
    Next
    
    Exit Function
End Function

Public Function contarValoresEnUnaColumna(context As Worksheet)
    Dim items As Long
    items = context.Columns.Count
    MsgBox "Elementos: " & items
    Exit Function
End Function


Public Function longitudDeUnaColumna(context As Worksheet) As Integer
    Dim conteo As Integer
    Dim valor As String
    
    conteo = 1
    valor = context.Cells(conteo, 1).value
    While valor <> ""
        Cells(conteo, 1).Activate
        'MsgBox "Celda #" & conteo & ": " & valor
        
        conteo = conteo + 1
        valor = context.Cells(conteo, 1).value
        'MsgBox conteo
    Wend
    'MsgBox "The last index of the list is:" & conteo - 1 & "--->" & valor
    MsgBox "final item is: " & ActiveCell.value
    'MsgBox "Loop ended correctly"
    'items = context.Columns.Count
    'MsgBox "Elementos: " & items
    longitudDeUnaColumna = conteo - 1
    Exit Function
End Function


Public Function demoFunction(context As Worksheet)
    Dim variable As Range
    
    Set variable = context.Cells.FindNext("Something")
    MsgBox variable
    Exit Function
End Function
