Attribute VB_Name = "o_FormateaSiguienteMes"
Option Explicit

'Mejorar parte que dice "Total PLANILLA"
'No es necesario borrar porque se borrará al final, y pueden ser 2 filas con total
'la planilla debe estar cerrada en la columna F. un "nada hasta aquí" sirve

Public lastRowSheet As Integer
Public mesActual As Integer
Public mesSiguiente As Integer
Public ultimaFilaSigMes As Integer
Public ultimaFilaMesActual As Integer
Public ultFilActConDatoGps As Integer
Public primeraFilSigMesConDato As Integer

    
Sub FormateaSiguienteMesMain()
    lastRowSheet = 80           'esta fila no debe tener datos
    If (Range("B" & lastRowSheet).Value <> "") Then
        MsgBox ("La fila " & lastRowSheet & " no debe tener datos para aplicar esta macro.")
        Exit Sub
    End If
    
    'Encuentra última fila siguiente mes
    ultimaFilaSigMes = ultFilSigMes()

    'Encuetra ultimo dato del mes actual
    ultimaFilaMesActual = buscaUltFilMesAct()
    
    'Encuetra ultimo dato del mes actual, que tiene consumo GPS
    Range("F" & ultimaFilaMesActual).Select
    If (Selection.Value = "") Then
        Selection.Offset(1, 0).Select
        ultFilActConDatoGps = Selection.End(xlUp).Row
    Else
        ultFilActConDatoGps = Selection.Row
    End If
    
    
    'Encuentra primer dato del siguiente mes, que tiene consumo GPS
    Range("F" & ultimaFilaMesActual).Offset(1, 0).Select
    If (Selection.Value = "") Then
        primeraFilSigMesConDato = Selection.End(xlDown).Row
    Else
        primeraFilSigMesConDato = Selection.Row
    End If
    
    'muestra en consola las variables recogidas
    Debug.Print "Mes actual: " & mesActual & " - Última fila mes actual: " & ultimaFilaMesActual
    Debug.Print "Última fila mes actual con dato GPS: " & ultFilActConDatoGps
    Debug.Print "Siguiente mes: " & mesSiguiente & " - Última fila siguiente mes: " & ultimaFilaSigMes
    Debug.Print "Primera fila siguiente mes con dato GPS: " & primeraFilSigMesConDato
        
    'Selecciona sección a copiar para poner al inicio de la planilla
    Range("B" & ultFilActConDatoGps & ":J" & ultimaFilaMesActual).Select
    Selection.Copy
    Range("B10").PasteSpecial xlPasteAll
    'borra espacio entre mes inicio de planilla y aquel del que se ocupará la actual planilla
    Dim inicioRangoVacio As Integer
    inicioRangoVacio = 11 + (ultimaFilaMesActual - ultFilActConDatoGps)
    Range("B" & inicioRangoVacio & ":Q" & (inicioRangoVacio + 1)).Select
    Selection.ClearContents
    
    'Selecciona sección a copiar que será el mes actual de la planilla
    Range("B" & (ultimaFilaMesActual + 1) & ":J" & ultimaFilaSigMes).Select
    Selection.Copy
    Range("B" & (inicioRangoVacio + 2)).PasteSpecial xlPasteAll
    Dim iniSigRangoVacio As Integer
    iniSigRangoVacio = (inicioRangoVacio + 2) + (ultimaFilaSigMes - ultimaFilaMesActual)
    Range("B" & iniSigRangoVacio & ":Q" & lastRowSheet).Select
    Selection.ClearContents
    
    'arregla formulas
    Dim celdaNuevaFormula As Integer
    celdaNuevaFormula = primeraFilSigMesConDato - (ultimaFilaMesActual + 1)
    Range("G" & (inicioRangoVacio + 2 + celdaNuevaFormula)).Select
    Call formulasIniMes
    Range("G" & (inicioRangoVacio + 2)).Select
    Selection.Offset(0, 4).Select
    Call formulasDerecha
    
    'Ubicarse en título del mes
    Range("A1").Select
    
'    Worksheets(ActiveSheet.Index + 1).Select
    
    End Sub

Function ultFilSigMes() As Integer
'    Dim filaACopiar As Integer
    Range("B" & lastRowSheet).End(xlUp).Offset(1, 5).Select
    ultFilSigMes = Selection.Row - 1
'    If (Range("G" & lastRowSheet).End(xlUp).Value = "Total PLANILLA") Then
'        Range("E" & filaACopiar & ":Q" & filaACopiar).Copy
'        Range("E" & (filaACopiar - 1)).Select
'        ActiveSheet.Paste
'        formateaResumen = filaACopiar - 2
'    Else
'        formateaResumen = filaACopiar - 1
'    End If
End Function

Function buscaUltFilMesAct() As Integer
    Dim rowReview As Integer
    mesActual = month(Range("B" & ultimaFilaSigMes).End(xlUp).Value)
    mesSiguiente = month(Range("B" & lastRowSheet).End(xlUp).Value)
    Range("B" & lastRowSheet).End(xlUp).Select
    For rowReview = ultimaFilaSigMes To 9 Step -1
        Selection.Offset(-1, 0).Select
        'Debug.Print (Month(Selection) & " = " & mesActual & " ?")
        If (month(Selection) = mesActual) Then
            buscaUltFilMesAct = Selection.Row
            Exit For
        End If
    Next rowReview
End Function

'Esta sub se aplica a partir de la celda seleccionada actual
'inicia el rango de suma 2 filas antes
'en la celda de la derecha arregla la formula a su estandar
'borra todas las celdas centrales de filas anteriores
Sub formulasIniMes()
    Dim cadena As String
    Dim resto As String
    Dim nuevaIniFormula As String
    Dim ultFilaBorrado As Integer
    
    cadena = Selection.Formula
    If InStr(cadena, ":") > 0 Then
        resto = Right(cadena, Len(cadena) - InStr(cadena, ":"))
        cadena = Right(cadena, Len(cadena) - InStr(cadena, "(") - 1)
        nuevaIniFormula = CInt(Left(cadena, InStr(cadena, ":") - 1)) - 2

    
        'corrección de celdas
        Selection.Formula = "=SUM(E" & nuevaIniFormula & ":" & resto
    Else
        Selection.Formula = "=E" & Selection.Row & "-F" & Selection.Row
    End If
    Selection.Offset(0, 1).Formula = "=G" & Selection.Row & "+H" & (Selection.Row - 1)
    
    'borrado
    ultFilaBorrado = Selection.Offset(-1, 0).Row
    Range("G9:H" & ultFilaBorrado).Select
    Selection.ClearContents
    Range("G" & ultFilaBorrado).Offset(1, 0).Select
End Sub

'a partir de la celda seleccionada
Sub formulasDerecha()
lastRowSheet = 80
    Dim filaInicial As Integer
    Dim fila As Integer
    
    fila = Selection.Row
    filaInicial = fila
    
    Range("K" & fila).Formula = "=J" & fila & "-I" & fila
    Range("L" & fila & ":Q" & fila).ClearContents
    fila = fila + 1
    Range("K" & fila).Formula = "=J" & fila & "-I" & fila
    Range("L" & fila).Formula = "=I" & fila & "-I" & (fila - 1)
    Range("M" & fila).Formula = "=F" & fila & "/L" & fila
    Range("N" & fila).ClearContents
    Range("O" & fila).Formula = "=E" & fila & "/(J" & fila & "-J" & fila & ")"
    Range("P" & fila).ClearContents
    Range("Q" & fila).Formula = "=(G" & fila & "*100)/E" & fila
    
    Range("K" & fila & ":Q" & fila).Select
    Selection.Copy
    Selection.Offset(1, 0).Select
    Selection.PasteSpecial xlPasteAll
    Selection.Offset(1, 0).Select
    Selection.PasteSpecial xlPasteAll
    Selection.Offset(1, 0).Select
    Selection.PasteSpecial xlPasteAll
    fila = fila + 3
    Range("N" & fila).Formula = "=PROMEDIO(M" & (fila - 3) & ":M" & fila & ")"
    Range("P" & fila).Formula = "=PROMEDIO(O" & (fila - 3) & ":O" & fila & ")"
    
    Range("K" & fila & ":Q" & fila).Select
    Selection.Copy
    Range("K" & (fila + 1) & ":Q" & lastRowSheet).Select
    Selection.PasteSpecial xlPasteAll
    Range("K9:Q" & (filaInicial - 1)).Select
    Selection.ClearContents
    
    Range("K" & filaInicial).Copy
    Range("K10:K" & (filaInicial - 3)).Select
    Selection.PasteSpecial xlPasteAll
    Range("L" & (filaInicial + 1)).Copy
    If (filaInicial - 3 >= 11) Then
        Range("L11:L" & (filaInicial - 3)).Select
        Selection.PasteSpecial xlPasteAll
    End If
End Sub
