Attribute VB_Name = "g_lapsosVacios"
Option Explicit

'corregir
'   comentario al inicio de línea en Function buscaInicio
'Se apica a 1 hoja
'importante: se usa la palabra "nada" como lectura de la planilla
'Descripción de la macro:
    'Cuando hay datos que faltan
        'borra las celdas centrales sobrantes
        'y arregla las formulas para los lapsos
    'selecciona los datos importantes

Public leastRowSheet As Integer
    
Sub lapsosVaciosMain()
    Dim maxFilasLapso As Integer
    Dim maxLapsos As Integer
    leastRowSheet = 80
    maxFilasLapso = 40  'cifra de seguridad
    maxLapsos = 20      'cifra de seguridad
    
    Dim inicioMes As String
    Dim celdaInicial As String
    Dim celdaFinal As String
    Dim lapsos As Integer
    Dim celdasVacias As Integer
    Dim numeroFila As String
    Dim inicioMesActual As Boolean
    'recoge variable celdaInicial y selecciona
    'celda "diferencia" en esa fila
    '¿ el lapso se inicia en este mes o en mes anterior?
    inicioMes = BuscaInicioMes()    'numero fila
    celdaInicial = buscaInicio(inicioMes)
    If (Left(celdaInicial, 2) = "si") Then
        inicioMesActual = True
    Else
        inicioMesActual = False
    End If
    celdaInicial = Right(celdaInicial, Len(celdaInicial) - 2)
    celdaFinal = Range(celdaInicial).Offset(0, 4).Address

    For lapsos = 1 To maxLapsos
        'cuenta cantidad de filas del siguiente lapso
        celdasVacias = cuentaCeldasVacias(maxFilasLapso, leastRowSheet)
        'si hay más de cierta cantidad de datos faltantes,
        'se asume final de datos recogidos
        'se selecciona datos para solicitar al GPS y termina Sub
        If (celdasVacias > maxFilasLapso) Then
            Range(celdaInicial, celdaFinal).Select
            Exit Sub
        End If
        'define nueva celdaFinal y
        'coloca formulas en fila final(columnas centrales)
        celdaFinal = Selection.Offset(0, -1).Address
        Call arreglaFormulas(ActiveCell.Row, celdasVacias, inicioMesActual)
        'borra celdas centrales del lapso y prepara búsqueda
        'del siguiente lapso
        Call borrar(celdasVacias)
        Range(celdaFinal).Offset(0, 1).Select
        'después del primer lapso, se inicia siempre en mes actual
        inicioMesActual = True
    Next lapsos
End Sub

Function BuscaInicioMes() As String
    Range("B10").Select
    Do While Not (Selection = "")
        Selection.Offset(1, 0).Select
    Loop
    Selection.End(xlDown).Select
    BuscaInicioMes = ActiveCell.Row
End Function

Function buscaInicio(iniMes As String) As String
    Dim filaInicial As String
    Range("G" & leastRowSheet).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
'Aquí hay que comprobar que la fia tenga un valor, porque si no
'hay que buscar la siguiente celda con contenido si la hay
    Selection.Offset(1, -1).Select
    Do While Not (Selection = "")
        Selection.Offset(1, 0).Select
    Loop
    Selection.Offset(-1, 1).Select
    filaInicial = ActiveCell.Row
    buscaInicio = "siB" & filaInicial
    
    'si el dato inicial está en el mes anterior
    If (filaInicial = iniMes) And (Range("F" & iniMes).Value = "") Then
        Range("F" & iniMes).Select
        Selection.End(xlUp).Select
        If (ActiveCell.Row > 9) Then
            'hay otroInicio, buscar
            filaInicial = ActiveCell.Row
            buscaInicio = "noB" & filaInicial
        End If
    End If
    Range("G" & filaInicial).Select
End Function

Function cuentaCeldasVacias(maxFil As Integer, leRowSheet As Integer) As Integer
    ' selecciona primera celda vacía
    Dim myRange As Range
    Set myRange = Selection
    myRange.Offset(1, -1).Select
    Do While Not (Selection = "")
        Selection.Offset(1, 0).Select
    Loop
    Set myRange = Selection
    ' comienza a contar desde celda 1
    cuentaCeldasVacias = 1
    Do
        Selection.Offset(1, 0).Select
        If (Selection.Row > leRowSheet) Then
            cuentaCeldasVacias = (maxFil + 10)
            Exit Do
        End If
        If Not (Selection.Value = "") Then
            If (InStr(Selection.Value, "nada") > 0) Then
                cuentaCeldasVacias = (maxFil + 10)
                Exit Do
            End If
            Selection.Offset(0, 1).Select
            Exit Do
        End If
        cuentaCeldasVacias = cuentaCeldasVacias + 1
    Loop While (cuentaCeldasVacias < (maxFil + 5))
End Function

Sub arreglaFormulas(numeroFila As String, celVacias As Integer, iniMesAct As Boolean)
    Dim sumDifAcum As Integer
    Selection.FormulaLocal = "=SUMA(E" & numeroFila - celVacias & ":E" & numeroFila & ")-F" & numeroFila
    Selection.Offset(0, 1).Select
    'si inicio del lapso está en este mes
    If (iniMesAct = True) Then
        sumDifAcum = (numeroFila - (celVacias + 1))
    ElseIf (iniMesAct = False) Then
        sumDifAcum = (numeroFila - 1)
    End If
    Selection.FormulaLocal = "=G" & numeroFila & "+H" & sumDifAcum
End Sub

Sub borrar(celVacias)
    Dim finSel, iniSel As String
    finSel = Selection.Offset(-1, 0).Address
    iniSel = Selection.Offset(-celVacias, -1).Address
    Range(Range(iniSel), Range(finSel)).Select
    Selection.ClearContents
End Sub
