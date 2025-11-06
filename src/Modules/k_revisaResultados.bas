Attribute VB_Name = "k_revisaResultados"
Option Explicit

'camionSeleccionado se puede evitar con una consulta más directa a Access

Public carpetaInformes As String
Public ultFila As Integer
Public maximalDifference As Integer

'argumento se utiliza como justificación para terminar el Sub revisaResultados
Public argumento As String
Public condicion1, condicion2, condicion3, condicion4, condicion5 As Boolean
'variables usadas en A.
Public desdeDiaTxt, desdeHoraTxt, desdeMinTxt As Integer
Public desdeDatoExtraStr As String
Public hastaDiaTxt, hastaHoraTxt, hastaMinTxt As Integer
'B
Public diaTitInforme, horaTitInforme As Integer
'C
Public desdeDiaInf, desdeYearInf, desdeMesInf, desdeHoraInf, desdeMinInf As Integer
Public hastaDiaInf, hastaYearInf, hastaMesInf, hastaHoraInf, hastaMinInf As Integer
Public CombustInf, OdometroInf As Double
'D
Public FilaIniDestino As Integer

Public desdeDiferencia, hastaDiferencia As Integer


    
Sub revisaResultados()
    
    carpetaInformes = "C:\Users\retac\Documents\dinero\transportes El Diamante\informes\"
    ultFila = 80
    maximalDifference = 20 'Debe ser un entero equivalente a [0 y 2]
    
    Dim camiones(24, 2) As String
    camiones(1, 1) = "KGRZ-60"
    camiones(2, 1) = "KGRZ-61"
    camiones(3, 1) = "KGRZ-62"
    camiones(4, 1) = "KGRZ-63"
    camiones(5, 1) = "KGRZ-64"
    camiones(6, 1) = "KGRZ-65"
    camiones(7, 1) = "KGRZ-66"
    camiones(8, 1) = "KGRZ-67"
    camiones(9, 1) = "KGRZ-69"
    camiones(10, 1) = "KGRZ-75"
    camiones(11, 1) = "JVJL-93"
    camiones(12, 1) = "JVJL-94"
    camiones(13, 1) = "JVJL-95"
    camiones(14, 1) = "JVJL-96"
    camiones(15, 1) = "JVJL-97"
    camiones(16, 1) = "JVJL-98"
    camiones(17, 1) = "JVJL-99"
    camiones(18, 1) = "JVJR-55"
    camiones(19, 1) = "JVJR-56"
    camiones(20, 1) = "JVJR-57"
    camiones(21, 1) = "JVJR-58"
    camiones(22, 1) = "JVJR-59"
    camiones(23, 1) = "JVJR-60"
    camiones(24, 1) = "JVJR-61"
    
    camiones(1, 2) = "K60"
    camiones(2, 2) = "K61"
    camiones(3, 2) = "K62"
    camiones(4, 2) = "K63"
    camiones(5, 2) = "K64"
    camiones(6, 2) = "K65"
    camiones(7, 2) = "K66"
    camiones(8, 2) = "K67"
    camiones(9, 2) = "K69"
    camiones(10, 2) = "K75"
    camiones(11, 2) = "JL93"
    camiones(12, 2) = "JL94"
    camiones(13, 2) = "JL95"
    camiones(14, 2) = "JL96"
    camiones(15, 2) = "JL97"
    camiones(16, 2) = "JL98"
    camiones(17, 2) = "JL99"
    camiones(18, 2) = "JR55"
    camiones(19, 2) = "JR56"
    camiones(20, 2) = "JR57"
    camiones(21, 2) = "JR58"
    camiones(22, 2) = "JR59"
    camiones(23, 2) = "JR60"
    camiones(24, 2) = "JR61"

    Dim nombreArchivo As String
    Dim wbLibroOrigen As Workbook, _
    wsHojaOrigen As Excel.Worksheet, _
    wbLibroDestino As Workbook, _
    wsHojaDestino As Worksheet
    Dim camionALlenar As Integer
    Dim propInforme() As String
    Dim camionInforme As String
    Set wbLibroDestino = Workbooks(ActiveWorkbook.Name)
    Set wsHojaDestino = wbLibroDestino.Worksheets(ActiveSheet.Name)

    'Buscar primer informe hecho para el camión de la hoja de destino
    nombreArchivo = Dir(carpetaInformes & "informe" & "*.xl*")
    camionALlenar = camionSeleccionado(wsHojaDestino)   'recibe un Integer, correspondiente al camion
    
    Dim condicion As Boolean
    '¿Hay informes en la carpeta?
    condicion = Not (nombreArchivo = "")
    argumento = "No se encontraron informes en la carpeta"
    Application.ScreenUpdating = False
    
    '--------------------------
    ' FILTRO 1 de la macro: Do While (condicion)
    '   busca el primer informe que corresponda al camión de la hoja destino que se encuentra activa
    '--------------------------
    'Si la condicion no se cumple, el FILTRO 2 terminará la macro con el argumento:
    '   No se encontraron informes en la carpeta
    'Si la condición se cumple, FILTRO 1 finalizará con 1 de los 2 "argumento" posibles:
    '   No hay informe para el camión que se busca
    '   El informe corresponde al camión que se busca
    Do While (condicion)
        Set wbLibroOrigen = Workbooks.Open(carpetaInformes & nombreArchivo)
        Set wsHojaOrigen = wbLibroOrigen.Worksheets(1)
        
        'caracteristicas del informe
        'propInforme(0): camion del informe; propInforme(1): mes y día; propInforme(2): hora y minuto de inicio del informe
        propInforme = caracteristicasInforme(wbLibroOrigen.Name)
        
        Debug.Print nombreArchivo
        '¿El informe corresponde al camión que se busca?
        condicion = (propInforme(0) = camionALlenar)
        If (condicion) Then
            argumento = "El informe corresponde al camión que se busca"
            Exit Do
        Else
            Workbooks(wbLibroOrigen.Name).Close
            nombreArchivo = Dir()       'Si hay más informes, condicion vuelve a ser true
            If (nombreArchivo = "") Then
                condicion = False
                argumento = "No hay informe para el camión que se busca"
            Else
                condicion = True
            End If
        End If
    Loop
    Application.ScreenUpdating = True
    
    '--------------------------
    ' FILTRO 2 de la macro: If InStr(argumento, "corresponde")
    '--------------------------
    'Si el FILTRO 2 no se cumple, finaliza la macro entregando el "argumento"
    'Si se cumple(argumento: El informe corresponde al camión que se busca)
    'la macro continúa
    If InStr(argumento, "corresponde") Then
        Debug.Print "La macro continúa"
    Else
        Debug.Print argumento
        Exit Sub
    End If

    '--------------------------
    'Lo siguiente sucede solamente si se cumplieron los "FILTRO" 1 y 2
    '--------------------------
    'leer datos del txt
    Dim rutaDatosOrigen As String
    Dim datosLinea(100) As String   '--los datos de cada línea del txt
    Dim cuentaLineas As Integer     '--cantidad de líneas en el txt, que tienen datos
    Dim txtInicio As Integer
    Dim txtFin As Integer
    rutaDatosOrigen = "C:\Users\retac\Documents\dinero\transportes El Diamante\" & camionALlenar & ".txt"
    cuentaLineas = 0
    Open rutaDatosOrigen For Input As #2
        Do While Not EOF(2)
            Line Input #2, datosLinea(cuentaLineas)
            'Debug.Print datosLinea(cuentaLineas)
            cuentaLineas = cuentaLineas + 1
        Loop
    Close #2
    '¿ está "p01" y "fin"?¿hay más de 1 línea con datos? -> (si)comenzar bucle desde línea 2/(no)error argumento y final
    txtInicio = InStr(datosLinea(0), "p01")
    txtFin = InStr(datosLinea(cuentaLineas - 1), "fin")
    If ((cuentaLineas > 3) And (txtInicio > 0) And (txtFin > 0)) Then
        argumento = "Hay suficientes datos en el archivo txt de origen"
    Else
        argumento = "No hay datos suficientes en el archivo txt de origen"
        Debug.Print argumento
        Exit Sub
    End If
    Dim diferenciaDesde As Integer
    Dim diferenciaHasta As Integer
    Dim infInicialIncorrecto As Boolean
    'Loop desde linea 2 a (cuentaLineas-1)y se revisan siguientes informes con dir()
    'For lineaTxtInicial = 2 To (cuentaLineas - 1)
            'Obtención de datos
            '   A.Se obtienen datos relevantes de la línea desde y la línea hasta
            '       Integer: desdeDiaTxt, desdeHoraTxt, desdeMinTxt, hastaDiaTxt, hastaHoraTxt, hastaMinTxt
            '       String: desdeDatoExtraStr
            '   B.dia y hora del titulo del informe de origen
            '       Integer: diaTitInforme, horaTitInforme
            '   C.datos "Inicio de", "Fin de", "Odómetro al partir" y "Combustible consumido total" del informe de origen
            '       Integer: desdeYearInf, desdeDiaInf, desdeHoraInf, desdeMinInf, hastaDiaInf, hastaHoraInf, hastaMinInf
            '       Double: CombustInf, OdometroInf As Double
            '   D.fila del dato inicial correspondiente, en informe de destino
            
            'A
            Call datosTxt(datosLinea(1), datosLinea(2))
            'B
            diaTitInforme = CInt(Right(propInforme(1), Len(propInforme(1)) - InStr(propInforme(1), "-")))
            horaTitInforme = CInt(Left(propInforme(2), InStr(propInforme(2), "-") - 1))
            'C
            Call datosInf(wsHojaOrigen)
            wbLibroOrigen.Close
            'D
            Call buscaFilaIniDestino(wsHojaDestino)
            If (argumento = "No se encontró la fila inicial") Then
                Debug.Print argumento
                Exit Sub
            End If
            
            ' La siguiente llamada se pueden comentar y descomentar, sirve para ver las variables obtenidas de A,B,C y D
            Call debugObtencionDatos
            
            'Se colocan los datos
            '   condicion1 = fechas informe desde corresponden a fechas txt desde(+-maximalDifference)
            '   ¿ diferenciaDesde?
            desdeDiferencia = calculaDiferencia("desde")
            '   condicion2 = hastaDiaInf, hastaHoraInf, hastaMinInf corresponden a hastaDiaTxt, hastaHoraTxt, hastaMinTxt(+-20minutos)
            '   ¿ diferenciaHasta?
            '   -> (si)Se colocan los datos en el informe de destino, se agregan las diferencias a la derecha ej: -16,3 | 19,1
            '   ->(no)Se selecciona el siguiente informe y se repite el proceso, manteniendo los datos txt
    'next lineaTxtInicial
    Debug.Print argumento
    
End Sub

Function camionSeleccionado(hdestino As Worksheet) As Integer
    Dim camion As Integer
    For camion = 1 To consultaCamion(0)
        If (hdestino.Name = consultaCamion(camion, 2)) Then
            camionSeleccionado = camion
        End If
    Next camion
End Function

Function caracteristicasInforme(nombreLibroOrigen As String) As String()
    Dim propiedadesInforme(3) As String
    Dim pos(3) As Integer
    propiedadesInforme(0) = Right(nombreLibroOrigen, Len(nombreLibroOrigen) - 7)
    pos(0) = InStr(propiedadesInforme(0), "_")
    propiedadesInforme(1) = Right(propiedadesInforme(0), Len(propiedadesInforme(0)) - pos(0))
    pos(1) = InStr(propiedadesInforme(1), "_")
    propiedadesInforme(2) = Right(propiedadesInforme(1), Len(propiedadesInforme(1)) - pos(1))
    propiedadesInforme(1) = Left(propiedadesInforme(1), (pos(1) - 1))
    pos(2) = InStr(propiedadesInforme(2), ".")
    propiedadesInforme(2) = Left(propiedadesInforme(2), (pos(1) - 1))
    propiedadesInforme(0) = Left(propiedadesInforme(0), (pos(0) - 1))
    caracteristicasInforme = propiedadesInforme()
End Function

Sub datosTxt(datosLineaUno As String, datosLineaDos As String)
    Dim restoFrom As Integer
    Dim restoStr As String
    desdeDiaTxt = CInt(Left(datosLineaUno, InStr(datosLineaUno, "-") - 1))
    restoFrom = InStr(datosLineaUno, " ")
    restoStr = Right(datosLineaUno, Len(datosLineaUno) - restoFrom)
    desdeHoraTxt = CInt(Left(restoStr, InStr(restoStr, ":") - 1))
    restoFrom = InStr(restoStr, ":")
    restoStr = Right(restoStr, Len(restoStr) - restoFrom)
    desdeMinTxt = CInt(Left(restoStr, InStr(restoStr, " ") - 1))
    restoFrom = InStr(restoStr, " ")
    restoStr = Right(restoStr, Len(restoStr) - restoFrom)
    restoFrom = InStr(restoStr, " ")
    restoStr = Right(restoStr, Len(restoStr) - restoFrom)
    desdeDatoExtraStr = restoStr
    hastaDiaTxt = CInt(Left(datosLineaDos, InStr(datosLineaDos, "-") - 1))
    restoFrom = InStr(datosLineaDos, " ")
    restoStr = Right(datosLineaDos, Len(datosLineaDos) - restoFrom)
    hastaHoraTxt = CInt(Left(restoStr, InStr(restoStr, ":") - 1))
    restoFrom = InStr(restoStr, ":")
    restoStr = Right(restoStr, Len(restoStr) - restoFrom)
    hastaMinTxt = CInt(Left(restoStr, InStr(restoStr, " ") - 1))
End Sub

Sub datosInf(wsHojaOrigen As Worksheet)
        Dim restoStr As String
        Dim restoFrom As Integer
        restoStr = wsHojaOrigen.Range("E13").Value
        desdeDiaInf = CInt(Left(restoStr, InStr(restoStr, "-") - 1))
        
        restoFrom = InStr(restoStr, "-")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        desdeMesInf = CInt(Left(restoStr, InStr(restoStr, "-") - 1))
        
        restoFrom = InStr(restoStr, "-")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        desdeYearInf = CInt(Left(restoStr, InStr(restoStr, " ") - 1))
        
        restoFrom = InStr(restoStr, " ")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        desdeHoraInf = CInt(Left(restoStr, InStr(restoStr, ":") - 1))
        
        desdeMinInf = CInt(Right(restoStr, Len(restoStr) - InStr(restoStr, ":")))

        restoStr = wsHojaOrigen.Range("AI13").Value
        hastaDiaInf = CInt(Left(restoStr, InStr(restoStr, "-") - 1))
        
        restoFrom = InStr(restoStr, "-")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        hastaMesInf = CInt(Left(restoStr, InStr(restoStr, "-") - 1))
        
        restoFrom = InStr(restoStr, "-")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        hastaYearInf = CInt(Left(restoStr, InStr(restoStr, " ") - 1))
        
        restoFrom = InStr(restoStr, " ")
        restoStr = Right(restoStr, Len(restoStr) - restoFrom)
        hastaHoraInf = CInt(Left(restoStr, InStr(restoStr, ":") - 1))
        
        hastaMinInf = CInt(Right(restoStr, Len(restoStr) - InStr(restoStr, ":")))
        
        restoStr = wsHojaOrigen.Range("BM9").Value
        restoFrom = InStr(restoStr, " ")
        OdometroInf = CDbl(Left(restoStr, restoFrom - 1))
        
        restoStr = wsHojaOrigen.Range("T69").Value
        CombustInf = CDbl(restoStr)
End Sub

Sub buscaFilaIniDestino(wsHojaDestino As Worksheet)
        Dim celdaSeleccionada As String
        Dim diaContCelda, horaContCelda, minContCelda As Integer
        celdaSeleccionada = wsHojaDestino.Range("B" & ultFila).End(xlUp).End(xlUp).Offset(-1, 0).Address
        Do
            celdaSeleccionada = Range(celdaSeleccionada).Offset(1, 0).Address
            diaContCelda = CInt(Format(Range(celdaSeleccionada), "D"))
            horaContCelda = CInt(Format(Range(celdaSeleccionada).Offset(0, 2), "H"))
            minContCelda = CInt(Format(Range(celdaSeleccionada).Offset(0, 2), "N"))
            condicion1 = (diaContCelda = desdeDiaTxt)
            condicion2 = (horaContCelda = desdeHoraTxt)
            condicion3 = (minContCelda = desdeMinTxt)
            condicion4 = (Not (condicion1 And condicion2 And condicion3))   'condición para que se repita el loop
            condicion5 = Range(celdaSeleccionada).Row <= ultFila            'condición para que se repita el loop
        Loop While (condicion4 And condicion5)
        FilaIniDestino = Range(celdaSeleccionada).Row
        If (FilaIniDestino > ultFila) Then
            argumento = "No se encontró la fila inicial"
        Else
            argumento = "Se encontró la fila inicial"
        End If
End Sub
Sub debugObtencionDatos()
    Debug.Print "Obtención de datos"
    Debug.Print "A. desdeDiaTxt, hora, min, desdeDatoExtraStr, hastaDiaTxt, hora, min"
    Debug.Print desdeDiaTxt & " " & desdeHoraTxt & " " & desdeMinTxt & " " & desdeDatoExtraStr & " " & hastaDiaTxt & " " & hastaHoraTxt & " " & hastaMinTxt
    Debug.Print "B. diaTitInforme, horaTitInforme"
    Debug.Print diaTitInforme & " " & horaTitInforme
    Debug.Print "C. desdeDiaInf, desdeMesInf, desdeYearInf, desdeHoraInf, desdeMinInf"
    Debug.Print desdeDiaInf & " " & desdeMesInf & " " & desdeYearInf & " " & desdeHoraInf & " " & desdeMinInf
    Debug.Print "hastaDiaInf, hastaMesInf, hastaYearInf, hastaHoraInf, hastaMinInf"
    Debug.Print hastaDiaInf & " " & hastaMesInf & " " & hastaYearInf & " " & hastaHoraInf & " " & hastaMinInf
    Debug.Print "Double -> combustible: " & CombustInf & " ; Odómetro: " & OdometroInf
    Debug.Print "C. FilaIniDestino"
    Debug.Print FilaIniDestino
End Sub
Function calculaDiferencia(lineaDatosStr As String) As Integer  'dineaDatosStr recibe "desde" o "hasta"
    'Algoritmo descrito en archivo macros-calculaDiferencia.txt
    Dim actYear As Integer, actMonth As Integer, lastDActMonth As Integer, lastDPrevMonth As Integer
    Dim dayTxt As Integer, hourTxt As Integer, minTxt As Integer
    Dim dayInf As Integer, hourInf As Integer, minInf As Integer
    If (lineaDatosStr = "desde") Then
        actYear = desdeYearInf
        actMonth = desdeMesInf
        dayTxt = desdeDiaTxt
        hourTxt = desdeHoraTxt
        minTxt = desdeMinTxt
        dayInf = desdeDiaInf
        hourInf = desdeHoraInf
        minInf = desdeMinInf
    Else 'lineaDatosStr = "hasta"
        actYear = hastaYearInf
        actMonth = hastaMesInf
        dayTxt = hastaDiaTxt
        hourTxt = hastaHoraTxt
        minTxt = hastaMinTxt
        dayInf = hastaDiaInf
        hourInf = hastaHoraInf
        minInf = hastaMinInf
    End If
    lastDActMonth = Format(DateSerial(actualYear, actualMonth + 1, 0), "D")
    lastDPrevMonth = Format(DateSerial(actualYear, actualMonth, 0), "D")
    condicion1 = (maximalDifference <= minTxt)
    condicion2 = (minTxt < (60 - maximalDifference))
    condicion3 = (dayInf = dayTxt)
    condicion4 = (hourInf = hourTxt)
    condicion 5 = (minInf <= minTxt + maximalDifference)    ' nInf <= n+md
    condicion 6 = (minInf >= 60 - maximalDifference)        ' nInf >= 60-md
    If (condicion1 And condicion2) Then             'IF1
        'dayInf y hourInf deben corresponder a dayTxt y hourTxt
        If (condicion3 And condicion4) Then
            calculaDiferencia = minInf - minTxt
        Else
            calculaDiferencia = 100
        End If
    ElseIf (Not condicion1) Then                    'IF2 n < md
    
    ElseIf (Not condicion2) Then                    'IF3
        condicion 5 = (algo)    ' algo
        condicion 6 = (algo)        ' algo
    End If
End Function
Sub pruebas()
    'propInforme = "245-06"
    'otraPropInforme = "5-020"
    'diaTitInforme = CInt(Left(propInforme, InStr(propInforme, "-") - 1))
    'mesTitInforme = CInt(Right(propInforme, Len(propInforme) - InStr(propInforme, "-")))
    'horaTitInforme = CInt(Left(otraPropInforme, InStr(otraPropInforme, "-") - 1))
    'minTitInforme = CInt(Right(otraPropInforme, Len(otraPropInforme) - InStr(otraPropInforme, "-")))
    ' Visual Basic expands the 4 in the statement Dim dub As Double = 4R to 4.0:
    Debug.Print Format(DateSerial(2018, 7 + 1, 0), "D")
End Sub
