Attribute VB_Name = "zTrabajoConFormatos"
'formatos() obtiene el color de fondo de celda seleccionada(incluso en formateo condicional)
'puede ser modificado para obtener color de fuente
'aplicaFormato() aplica numero de formato
'aplicaFormatoAlBloque() aplica un formato de letra determinado a un bloque

Sub formatos()
    Dim resultado As Long
    resultado = ColorCelda(Selection)
    Debug.Print resultado
    
End Sub

Sub aplicaFormato()
    Selection.Interior.ColorIndex = 35
    
End Sub

Sub aplicaFormatoAlBloque()
    With Range("B6:H20").Font
        .Name = "Arial"
        .Size = 9
        .Italic = True
    End With
End Sub

'función copiada de sitio web
Function ColorCelda(celda As Range, _
    Optional ColorRellenoCelda As Boolean = True, _
    Optional ReturnColorIndex As Long = True) As Long
    
    ' ColorRellenoCelda - Opcional, tipo Boolean (valor por defecto = TRUE)
    '                true -> color de Relleno, según la propiedad de .Color o .ColorIndex
    '                determinada en el valor del tercer argumento ReturnColorIndex
    '                FALSO  -> color de la Fuente,
    '                según la propiedad de .Color o .ColorIndex
    ' ReturnColorIndex - Opcional, tipo Boolean (valor por defecto = TRUE)
    '                    VERDADERO -> hará que nuestra función emplee la propiedad .ColorIndex
    '                    FALSO -> .Color
    Dim X As Long
    Dim Test As Boolean
    Dim CeldaActiva As String
    
    Application.Volatile
    CeldaActiva = ActiveCell.Address
    'recorremos todos los formatos condicionales existentes en la celda de estudio
    For X = 1 To celda.FormatConditions.Count
        With celda.FormatConditions(X)
            'si la condición responde al Valor de una celda
            'sabiendo que .Formula1 y .Formula2 son los valores que podemos informar al configurar nuestro formato condicional
            If .Type = xlCellValue Then
                Select Case .Operator
                    'Evaluate equivaldría a la función INDIRECTO de la hoja de cálculo
                    'convierte, por tanto, un nombre de Microsoft Excel en un objeto, valor o referencia.
                    Case xlBetween:      Test = celda.Value >= Evaluate(.Formula1) And celda.Value <= Evaluate(.Formula2)
                    Case xlNotBetween:   Test = celda.Value <= Evaluate(.Formula1) Or celda.Value >= Evaluate(.Formula2)
                    Case xlEqual:        Test = Evaluate(.Formula1) = celda.Value
                    Case xlNotEqual:     Test = Evaluate(.Formula1) <> celda.Value
                    Case xlGreater:      Test = celda.Value > Evaluate(.Formula1)
                    Case xlLess:         Test = celda.Value < Evaluate(.Formula1)
                    Case xlGreaterEqual: Test = celda.Value >= Evaluate(.Formula1)
                    Case xlLessEqual:    Test = celda.Value <= Evaluate(.Formula1)
                End Select
            'si por contra la condición corresponde a una expresión o fórmula
            ElseIf .Type = xlExpression Then
                Application.ScreenUpdating = False
                celda.Select
                Test = Evaluate(.Formula1)
                Range(CeldaActiva).Select
                Application.ScreenUpdating = True
            End If
            
            'Verificamos nuestra comparativa Test, cuando sea VERDADERA
            If Test Then
                If ColorRellenoCelda Then
                    'y además nuestro segundo y tercer argumento se han informado como CIERTOS
                    'controlando el tercer argumento con la función VBA IIf
                    ColorCelda = IIf(ReturnColorIndex, .Interior.ColorIndex, .Interior.Color)
                Else
                    'pero si nuestro segundo argumento es FALSO y el tercer argumento se ha informado como CIERTO
                    'controlando el tercer argumento con la función VBA IIf
                    ColorCelda = IIf(ReturnColorIndex, .Font.ColorIndex, .Font.Color)
                End If
                'salimos de la función...
                Exit Function
            End If
        End With
    Next

    'si no tuviera Formato condicional aplicado...
    If ColorRellenoCelda Then
        ColorCelda = IIf(ReturnColorIndex, celda.Interior.ColorIndex, celda.Interior.Color)
    Else
        ColorCelda = IIf(ReturnColorIndex, celda.Font.ColorIndex, celda.Font.Color)
    End If
End Function

