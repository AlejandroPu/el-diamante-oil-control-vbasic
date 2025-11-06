Attribute VB_Name = "m_cerrarInforme"
Option Explicit

'es necesario cerrar la columna F para el mes siguiente
'si no hay datos, tal vez no tiene sentido cerrar el mes

Public ultimaFilaInforme As Integer
Public ultimaFilaMesActual As Integer
Public ultimaFilaSigMes As Integer
Public mesActual As Integer
Public mesSiguiente As Integer

Sub cerrarInformeMain()
Attribute cerrarInformeMain.VB_Description = "Cerrar informe"
Attribute cerrarInformeMain.VB_ProcData.VB_Invoke_Func = " \n14"
    'Última fila del informe
    ultimaFilaInforme = 80          'esta fila no debe tener datos
    If (Range("B" & ultimaFilaInforme).Value <> "") Then
        MsgBox ("La fila " & ultimaFilaInforme & " no debe tener datos para aplicar esta macro.")
        Exit Sub
    End If
    
    'Definir filas importantes
    Dim filasImportantes() As Integer
    Dim primeraFilSigMesConDato As Integer
    ultimaFilaSigMes = Range("B" & ultimaFilaInforme).End(xlUp).Row
    ultimaFilaMesActual = buscaUltFilMesAct()
    'Encuentra primer dato del siguiente mes, que tiene consumo GPS
    Range("F" & ultimaFilaMesActual).Offset(1, 0).Select
    If (Selection.Value = "") Then
        Selection.End(xlDown).Select
        primeraFilSigMesConDato = Selection.Row
    Else
        primeraFilSigMesConDato = Selection.Row
    End If
    'borra final
    Range("B" & (primeraFilSigMesConDato + 1) & ":Q" & ultimaFilaInforme).Select
    Selection.ClearContents
    filasImportantes = buscaFilasImportantes()
    
    'Borrar celdas sobrantes en rango G:H y K:Q
    'y agregar datos de resumen
    Call datosFinales(filasImportantes, ultimaFilaInforme)
    
    'seleccionar dato inicial y final para paso1txt
    Dim a As Integer
    Dim b As Integer
    a = (filasImportantes(0) - 1)
    b = (filasImportantes(1))
    Range("B" & a & ":F" & a & ",B" & b & ":F" & b).Select
    'repetir para todos los meses

End Sub

Function buscaFilasImportantes() As Integer()
    Dim filasArray(3) As Integer
    Range("F9").End(xlDown).Offset(1, -1).Select
    filasArray(0) = Selection.Row
    Range("F" & ultimaFilaMesActual).Offset(1, 0).Select
    If (Selection.Value = "") Then
        Selection.End(xlDown).Select
    End If
    filasArray(1) = Selection.Row
    Range("B" & ultimaFilaInforme).End(xlUp).Offset(1, 0).Select
    filasArray(2) = Selection.Row
    buscaFilasImportantes = filasArray()
End Function

Sub datosFinales(filImp() As Integer, ultFilInf As Integer)
    'Agregar suma consumos columna E
    Range("E" & filImp(2)).FormulaLocal = "=suma(E" & filImp(0) & ":E" & filImp(1) & ")"
    Range("F" & filImp(2)).FormulaLocal = "=suma(F" & filImp(0) & ":F" & filImp(1) & ")"
    Range("G" & filImp(2)).FormulaLocal = "Total PLANILLA"
    Range("G" & (filImp(2) + 1)).FormulaLocal = "Total GPS"
    Range("G" & filImp(2)).Font.Bold = True
    Range("G" & (filImp(2) + 1)).Font.Bold = True
    Range("G" & filImp(2)).FormatConditions.Delete
    Range("G" & (filImp(2) + 1)).FormatConditions.Delete
    Range("G" & filImp(2)).HorizontalAlignment = xlLeft
    Range("G" & (filImp(2) + 1)).HorizontalAlignment = xlLeft
    'Agregar porcentaje diferencia entre datos de consumo total del mes
    Range("Q" & filImp(2)).FormulaLocal = "=(E" & filImp(2) & "-F" & filImp(2) & ")/E" & filImp(2)
    Range("Q" & filImp(2)).Font.Bold = True
    Range("Q" & filImp(2)).NumberFormat = "0.00%"
End Sub

Function buscaUltFilMesAct() As Integer
    Dim rowReview As Integer
    mesActual = month(Range("B" & ultimaFilaSigMes).End(xlUp).Value)
    Range("B" & ultimaFilaSigMes).Select
    mesSiguiente = month(Selection.Value)
    For rowReview = ultimaFilaSigMes To 9 Step -1
        Selection.Offset(-1, 0).Select
        'Debug.Print (Month(Selection) & " = " & mesActual & " ?")
        If (month(Selection) = mesActual) Then
            buscaUltFilMesAct = Selection.Row
            Exit For
        End If
    Next rowReview
End Function
