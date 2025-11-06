Attribute VB_Name = "e_fuenteAInforme"
Option Explicit
'¿detecta el último día del mes y agrega el primero del siguiente?
'Hay que borrar las variables camiones() comentadas, si la macro sigue funcionando bien

Sub fuenteAInformeMain()

    Dim isNextMonth As Integer
    Dim cantCamiones As Integer
    Dim shortNameTr As String
    Dim longNameTr As String
    Dim camion As Integer
    Dim celdaHora As String
    
    isNextMonth = MsgBox("La fuente corresponde al siguiente mes?", vbYesNo + vbInformation + vbDefaultButton2, "Selección siguiente mes")
    Debug.Print ("siguiente mes = " & isNextMonth)                      '7 no(default), 6 Sí
    cantCamiones = consultarCamion(0)
    
    For camion = 1 To cantCamiones
        Debug.Print camion
                                                                        'buscar último dato de fecha en la hoja final
        shortNameTr = consultarCamion(camion, 2)
        longNameTr = consultarCamion(camion, 1)
        
        datoFinal (shortNameTr)
'y si hay dato en fila 13?
                                                                        'busca fecha y hora del datoFinal, en hoja Fuente
        celdaHora = datoFuente(longNameTr, shortNameTr, isNextMonth)
        If celdaHora = "no" Then
            MsgBox "No se encuentra dato en la fuente camion " & camion
            Sheets(shortNameTr).Select
        Else
            'copia de Fuente a la hoja
            Call haciaHoja(celdaHora, shortNameTr)
            'agrega columnas centrales
            Dim numeroInicio As String
            Dim numeroFin As String
            numeroInicio = numeroInicial()
            numeroFin = numeroFinal()
            Call columnasCentrales(numeroInicio, numeroFin)
            'selecciona los datos que se solicitarán a GPS
            Call datosAlGps(numeroInicio, numeroFin)
        End If
    Next camion
    
End Sub

Function datoFinal(cam As String)
    Sheets(cam).Select
    Range("B200").Select
    Selection.End(xlUp).Select
End Function

Function datoFuente(cam As String, cam2 As String, nextMonth As Integer) As String
    Dim fecha As Date
    Dim hora As Double
    Dim myRange As Range
    Dim c As Range
                                                                                'busca fecha
    fecha = Selection.Value
    Sheets("Fuente").Select
    
    Columns("A").Select
    Debug.Print cam
    Selection.Find(cam, , LookIn:=xlValues, LookAt:=xlPart).Select
    Selection.Offset(0, 1).Select
    If (nextMonth = 7) Then                                                     'Si es siguiente mes(6), se queda en el mismo lugar
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Find(fecha, , LookIn:=xlValues, LookAt:=xlPart).Select
                                                                                            'busca hora
        Sheets(cam2).Select
        Selection.Offset(0, 2).Select
        hora = Selection.Value
        Sheets("Fuente").Select
        Selection.Offset(0, 2).Select
        Range(Selection, Selection.End(xlDown)).Select
        Set myRange = Selection
        datoFuente = "no"
        For Each c In myRange.Cells
            If c.Value = hora Then
                datoFuente = c.Address
                Range(datoFuente).Select
                Exit For
            End If
        Next c
        'verifica fecha y camion correcta
        Selection.Offset(0, -2).Select
        If Not (fecha = Selection.Value) Then
            Debug.Print "fecha no corresponde"
        End If
        Selection.Offset(0, -1).Select
        If Not (cam = Left(Selection.Find(cam, , LookIn:=xlValues, LookAt:=xlPart), 7)) Then
            Debug.Print "camion no corresponde"
        End If
        Selection.Offset(0, 1).Select
        'si celda sin nada, nada que copiar
        Selection.Offset(1, 0).Select
        If (Selection.Value = "") Then
            Debug.Print "nada que copiar o no esta el dato"
            datoFuente = "no"
        End If
    Else
        datoFuente = Selection.Address
        Sheets(cam2).Select
        Selection.Offset(0, 2).Select
        Sheets("Fuente").Select
    End If
End Function

Sub haciaHoja(cellHora, cam2)
    cellHora = Selection.Address
    Selection.End(xlDown).Select
    Selection.Offset(0, 8).Select
    Range(cellHora, Selection).Select
    Selection.Copy
    Sheets(cam2).Select
    Selection.Offset(1, -2).Select
    Selection.PasteSpecial xlPasteValues
    Selection.End(xlToLeft).Select
End Sub

Function numeroInicial()
    numeroInicial = Selection.Address
    numeroInicial = Right(numeroInicial, Len(numeroInicial) - InStrRev(numeroInicial, "$"))
    Selection.Offset(0, 1).Select
    Selection.End(xlDown).Select
End Function

Function numeroFinal()
    numeroFinal = Selection.Address
    numeroFinal = Right(numeroFinal, Len(numeroFinal) - InStrRev(numeroFinal, "$"))
End Function
    
Sub columnasCentrales(n1 As String, n2 As String)
    Range(Range("G" & n1), Range("G" & n2)).Select
    Selection.Value = "=E" & n1 & "-F" & n1
    Range(Range("H" & n1), Range("H" & n2)).Select
    Selection.Value = "=G" & n1 & "+H" & (n1 - 1)
End Sub

Sub datosAlGps(n1 As String, n2 As String)
    Range("F" & n1).End(xlUp).Select
    If (ActiveCell.Row < 13) Then
        Range("F13").Select
    End If
    Selection.Offset(0, -4).Select
    Range(Range(Selection.Address), Range("F" & n2)).Select
End Sub
