Attribute VB_Name = "yc_informeDeCopec"
Option Explicit
Dim esteRegistro(24, 100) As InformeCopec
Dim registrosCamion(24) As Integer              'cantidad de registros para el camion (n)
Dim fuenteRegistro(24, 100) As String
Dim datosFuenteCant(24) As Integer          'cantidad de registros Fuente para el camion (n)

Sub registrarInfxCamion()
Attribute registrarInfxCamion.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim desdeCamion, hastaCamion, camion As Integer
    desdeCamion = Application.InputBox(prompt:="Ingresar DESDE qué camión extraer información hacia informe final.", Title:="Desde", Default:=1, Type:=1)
    hastaCamion = Application.InputBox(prompt:="Ingresar HASTA qué camión.", Title:="Hasta", Default:=24, Type:=1)
    Application.ScreenUpdating = False
    For camion = desdeCamion To hastaCamion
        Debug.Print (Chr(10) & "camion: " & camion);
        Call revisarInfxCamion(camion)
        Call revisaValidezDatos(camion)
        Call lapsosVaciosMain
    Next camion
    Application.ScreenUpdating = True
End Sub

Sub revisarInfxCamion(ByVal camion As Integer)
    
    Dim rutaInfCOPEC As String, registro As Integer, nombreArchivo As String _
        , cuentaLoop As Integer, startStringPos As Integer, informeARevisar As Excel.Workbook
    
    rutaInfCOPEC = "D:\Mis Documentos HD\dinero\transportes El Diamante\informes"
    registro = 1
    nombreArchivo = Dir(rutaInfCOPEC & "\informe" & camion & "_*.xls*")
    
    cuentaLoop = 1
    startStringPos = 10
    Do While nombreArchivo <> ""
        If camion >= 10 Then startStringPos = 11
        
        If (cuentaLoop = 1) Then
            Debug.Print (Chr(10) & " |---informe: " & Mid(nombreArchivo, startStringPos, 2));
            cuentaLoop = -1
        Else
            Debug.Print (", " & Mid(nombreArchivo, startStringPos, 2));
        End If
        
        Set informeARevisar = Workbooks.Open(rutaInfCOPEC & "\" & nombreArchivo)
        Call buscarDatosRelevantes(camion, registro)
        Workbooks(nombreArchivo).Close
        'Debug.Print esteRegistro(camion, registro).allTheData
        registro = registro + 1
        nombreArchivo = Dir
    Loop
    
    registrosCamion(camion) = registro - 1
End Sub

Sub buscarDatosRelevantes(ByVal cam As Integer, ByVal reg As Integer)         'camión, inicio, fin, Odómetro, litros consumidos
    Dim vehiculo As String, vehiculoFrom As Integer, vehiculoLength As String
    Dim inicio As Date
    Dim fin As Date
    Dim odometro As Single
        Dim odStr As String         'datos secundarios para obtener odometro.
        Dim dotPos As Integer
    Dim combustible As Single
        Dim combStr As String
    vehiculo = "no"
    inicio = #10/10/2010#
    fin = #10/10/2020#
    odometro = 0
    combustible = 0
    On Error Resume Next
    vehiculo = Cells.Find(What:="Vehiculo:", After:=Range("A1")).Offset(0, 1).Value
    If vehiculo <> "no" Then
            Dim contVehiculoFirstReview As Boolean
            contVehiculoFirstReview = True
            Do
                vehiculoLength = Len(vehiculo)
                vehiculoFrom = InStrRev(vehiculo, " ")
                If vehiculoLength = vehiculoFrom Then
                    vehiculo = Mid(vehiculo, 1, vehiculoLength - 1)
                Else
                    contVehiculoFirstReview = False
                End If
            Loop While (contVehiculoFirstReview)
            vehiculoFrom = InStrRev(vehiculo, " ") + 1
            vehiculoLength = Len(vehiculo) - vehiculoFrom + 1
            vehiculo = Mid(vehiculo, vehiculoFrom, vehiculoLength)
            If Mid(vehiculo, 1, 1) = "K" Then
                vehiculo = "K" & Mid(vehiculo, Len(vehiculo) - 1, Len(vehiculo))
            ElseIf Mid(vehiculo, 1, 4) = "JVJL" Then
                vehiculo = "JL" & Mid(vehiculo, Len(vehiculo) - 1, Len(vehiculo))
            ElseIf Mid(vehiculo, 1, 4) = "JVVR" Then
                vehiculo = "JR" & Mid(vehiculo, Len(vehiculo) - 1, Len(vehiculo))
            End If
    
        inicio = Cells.Find(What:="Inicio de viaje", After:=Range("A1")).Offset(0, 1).Value
        fin = Cells.Find(What:="Fin de Viaje", After:=Range("A1")).Offset(0, 1).Value
        odStr = Cells.Find(What:="Odómetro al Partir", After:=Range("A1")).Offset(0, 2).Value   'formato "123.567,9 km"
            odStr = Mid(odStr, 1, Len(odStr) - 3)       ' se quita " km"
            dotPos = InStr(1, odStr, ".")               ' posición del punto
            If dotPos > 0 Then odStr = Mid(odStr, 1, dotPos - 1) & Mid(odStr, dotPos + 1, Len(odStr) - dotPos)
            odometro = CSng(odStr)
        combStr = Cells.Find(What:="Total:", After:=Range("A1")).Offset(0, 1).Value
            dotPos = InStr(1, combStr, ".")             ' posición del punto
            If dotPos > 0 Then combStr = Mid(combStr, 1, dotPos - 1) & Mid(combStr, dotPos + 1, Len(combStr) - dotPos)
            combustible = CSng(combStr)
    End If
    Set esteRegistro(cam, reg) = New InformeCopec
    esteRegistro(cam, reg).let_truckId = vehiculo
    esteRegistro(cam, reg).let_startReg = inicio
    esteRegistro(cam, reg).let_endReg = fin
    esteRegistro(cam, reg).let_odReg = odometro
    esteRegistro(cam, reg).let_consReg = combustible
End Sub

Sub revisaValidezDatos(ByVal vehiculoNum As Integer)
    Dim vehiculo As String
    Dim filReg As Integer
    Dim asocTxtToInf() As Integer   'asociación entre cada (línea del txt fuente) con el número de fila en el informe final EMS a la que corresponde.
    ' variables para comprobación y colocación
    Dim aRegDate, aRegTime, bRegDate, bRegTime As String
    Dim aSource, aSourceDate, aSourceTime As String
    Dim bSource, bSourceDate, bSourceTime As String
    Dim auxStrFinder As Integer
    Dim mes As String
    Dim hora1, hora2, hora3 As Date
    Dim cond1, cond2 As Boolean
    Dim diff1, diff2 As Double
    
    '¿ hay la misma cantidad de datos en el txt correspondiente que los registros?
    '   poner cada línea del txt fuente en un fuenteRegistro(camion, línea).
    '   Cada línea está relacionada con cada inicio de registro de los informes EMS,excepto la última: esteRegistro(vehiculoNum, registro).
    '   cantidad de registros: registrosCamion(vehiculoNum) ; cantidad de lineas fuente: datosFuenteCant(vehiculoNum) ( 1 más que los registros !!)
    Call revisaRegistroFuente(vehiculoNum)
    If (datosFuenteCant(vehiculoNum) - 1) <> registrosCamion(vehiculoNum) Then
        Debug.Print ("_ Incongruencia entre cantidad de valores fuenteTXT y registros de informes EMS para el camion " & vehiculoNum & ": " & vehiculo)
        Exit Sub
    End If

    ' Activar libro de informe EMS final
    activaLibro (vehiculoNum)
            
    'Buscar correspondencia de líneas txt fuente con filas en informe final.
    asocTxtToInf = buscaCorrespondencia(vehiculoNum)
    
    For filReg = 1 To registrosCamion(vehiculoNum)
        vehiculo = esteRegistro(vehiculoNum, filReg).get_truckId
        If vehiculo <> "no" Then

            'comprobación y colocación. recolección de variables.
            aRegDate = Application.Text(DateValue(esteRegistro(vehiculoNum, filReg).get_startReg), "dd-mmm")
                mes = mesEngToSpa(Mid(aRegDate, 4, 3))
                aRegDate = Mid(aRegDate, 1, 3) & mes
            aRegTime = Application.Text(TimeValue(esteRegistro(vehiculoNum, filReg).get_startReg), "hh:mm")
            bRegDate = Application.Text(DateValue(esteRegistro(vehiculoNum, filReg).get_endReg), "dd-mmm")
                mes = mesEngToSpa(Mid(bRegDate, 4, 3))
                bRegDate = Mid(bRegDate, 1, 3) & mes
            bRegTime = Application.Text(TimeValue(esteRegistro(vehiculoNum, filReg).get_endReg), "hh:mm")
            aSource = fuenteRegistro(vehiculoNum, filReg)
                auxStrFinder = InStr(1, aSource, " ")
                mes = mesEngToSpa(Mid(aSource, 4, 3))
                aSourceDate = Mid(aSource, 1, 3) & mes
                aSourceTime = Mid(aSource, auxStrFinder + 1, 5)
            bSource = fuenteRegistro(vehiculoNum, filReg + 1)
                auxStrFinder = InStr(1, bSource, " ")
                mes = mesEngToSpa(Mid(bSource, 4, 3))
                bSourceDate = Mid(bSource, 1, 3) & mes
                bSourceTime = Mid(bSource, auxStrFinder + 1, 5)
        
            'comprobación y colocación. Comprobación.
            'Sea fecha-hora inicial [a], y final [b]
            'Comprobar si ( |a txt - a registro| <= 30 minutos ) Y ( |b txt - b registro| <= 30 )
            '               sí -> insertar en excel, agregar diferencia de minutos. Si no aparece consumo GPS en fila inicio, escribir "ok"
            '               no -> ¿es el último registro?
            '                       si -> agregar "nada hasta aquí"
            '                       no -> no colocar nada en esa fila
            
            'cond1 = (aSourceTime - aRegTime) <= 30
                hora1 = DateValue(aSourceDate) + TimeValue(aSourceTime)
                hora2 = DateValue(aRegDate) + TimeValue(aRegTime)
                hora3 = CDate(hora1 - hora2)
                diff1 = Abs(60 * 24 * CDbl(hora3))
                'Debug.Print (hora1 & " // " & hora2 & " // " & diff1);
            cond1 = (Abs(60 * 24 * CDbl(hora3)) <= 30)
                hora1 = DateValue(bSourceDate) + TimeValue(bSourceTime)
                hora2 = DateValue(bRegDate) + TimeValue(bRegTime)
                hora3 = CDate(hora1 - hora2)
                diff2 = Abs(60 * 24 * CDbl(hora3))
                'Debug.Print (" | " & hora1 & " // " & hora2 & " // " & diff2)
            cond2 = (Abs(60 * 24 * CDbl(hora3)) <= 30)
        '    Debug.Print (cond1 & " and " & cond2 & " = ");
        '    Debug.Print (cond1 And cond2)
            If (cond1 And cond2) Then
                'insertar en excel, agregar diferencia de minutos
                Range("I" & asocTxtToInf(filReg)).Value = esteRegistro(vehiculoNum, filReg).get_odReg
                Range("F" & asocTxtToInf(filReg + 1)).Value = esteRegistro(vehiculoNum, filReg).get_consReg
                'Debug.Print (esteRegistro(vehiculoNum, filReg).allTheData)
                'si no aparece consumo GPS en fila inicio o aparece "nada*", además escribir "ok"
                If ((Range("F" & asocTxtToInf(filReg)).Value = 0) Or (Range("F" & asocTxtToInf(filReg)).Value Like "?ada*")) Then
                    Range("F" & asocTxtToInf(filReg)).Value = "ok"
                End If
            ElseIf filReg = registrosCamion(vehiculoNum) Then
                Dim putNothingUntilHere As Boolean
                Dim filaHaciaArriba As Integer
                filaHaciaArriba = Range("F80").End(xlUp).Row
                putNothingUntilHere = Not (asocTxtToInf(filReg + 1) < filaHaciaArriba)
                If putNothingUntilHere Then
                    Range("F" & asocTxtToInf(filReg + 1)).Value = "nada hasta aquí"
                End If
            Else
                Range("F" & asocTxtToInf(filReg + 1)).Value = ""
                If Range("F" & asocTxtToInf(filReg)).Value Like "?ada*" Then
                    Range("F" & asocTxtToInf(filReg)).Value = ""
                End If
            End If
            If diff1 > 30 Then
                Range("F" & asocTxtToInf(filReg)).Value = ""
                'Debug.Print ("I" & asocTxtToInf(filReg - 1))
                Range("I" & asocTxtToInf(filReg) - 1).Value = ""
            ElseIf (((Range("R" & asocTxtToInf(filReg)).Value <= 30) Or (Range("R" & asocTxtToInf(filReg)).Value = "no")) _
                        And (Range("F" & asocTxtToInf(filReg)).Value = 0)) Then
                Range("F" & asocTxtToInf(filReg)).Value = "ok"
            End If
            Range("S" & asocTxtToInf(filReg)).Value = Round(diff1)
            Range("R" & asocTxtToInf(filReg + 1)).Value = Round(diff2)
        Else
            Range("S" & asocTxtToInf(filReg)).Value = vehiculo
            Range("R" & asocTxtToInf(filReg + 1)).Value = vehiculo
            If Range("F" & asocTxtToInf(filReg)).Value Like "?ada*" Then
                Range("F" & asocTxtToInf(filReg)).Value = ""
            End If
        End If
    Next filReg
    Debug.Print ("_ Done!")
End Sub
Sub probando()
        If (Range("F33").Value > 0) And (Range("F33").Value <> "ok") Then
            Debug.Print ("eureka!")
        End If
End Sub


Function buscaCorrespondencia(cam) As Integer()
    Dim endDateText As Integer
    Dim fReg As String
    Dim fRegRow As Integer
    Dim fRegDate As String
    Dim fRegHour As String
    Dim asocBRow As Integer
    Dim asocRowVal As Date
    Dim asocRowToTxtLine(100) As Integer   'datosFuenteCant(cam) + 1
    'fRegRow = 1
    For fRegRow = 1 To datosFuenteCant(cam)
        fReg = fuenteRegistro(cam, fRegRow)
        endDateText = InStr(1, fReg, " ")       'Puede ser Chr(9) en vez de un espacio en blanco y no funciona.
                                                'En tal caso, probablemente no se guardaron los txt con la macro
                                                'y no están correctamente formateados para esta macro.
        Dim letras As Integer
        fRegDate = Mid(fReg, 1, endDateText - 1)    'endDateText siempre vale 7
        fRegHour = Mid(fReg, endDateText + 1, 5)
        For asocBRow = 10 To 100                                                 'en busca de siguiente fecha
            asocRowVal = Range("b" & asocBRow).Value
            If Application.Text(asocRowVal, "[$-409]dd-mmm") = fRegDate Then
                If Application.Text(Range("D" & asocBRow).Value, "hh:mm") = fRegHour Then
                    asocRowToTxtLine(fRegRow) = asocBRow
                    'Debug.Print (fRegRow & ": " & asocRowToTxtLine(fRegRow))
                    Exit For
                End If
            End If
        Next asocBRow
    Next fRegRow
    buscaCorrespondencia = asocRowToTxtLine
End Function

Sub revisaRegistroFuente(ByVal textoNumero As Integer)
    Dim ruta As String
    Dim iFile As Integer
    datosFuenteCant(0) = 0
    datosFuenteCant(textoNumero) = 0
    ruta = "D:\Mis Documentos HD\dinero\transportes El Diamante\" & textoNumero & ".txt"
    Debug.Print (Chr(10) & " |- verificando: " & ruta)

    iFile = FreeFile
    Open ruta For Input As iFile
    Do While Not EOF(iFile)
        Line Input #iFile, fuenteRegistro(textoNumero, datosFuenteCant(textoNumero))
        datosFuenteCant(textoNumero) = datosFuenteCant(textoNumero) + 1
    Loop
    datosFuenteCant(textoNumero) = datosFuenteCant(textoNumero) - 2
    Close iFile
End Sub

Function mesEngToSpa(ByVal month As String) As String
    If month = "Jan" Then
        mesEngToSpa = "Ene"
    ElseIf month = "Feb" Then
        mesEngToSpa = "Feb"
    ElseIf month = "Mar" Then
        mesEngToSpa = "Mar"
    ElseIf month = "Apr" Then
        mesEngToSpa = "Abr"
    ElseIf month = "May" Then
        mesEngToSpa = "May"
    ElseIf month = "Jun" Then
        mesEngToSpa = "Jun"
    ElseIf month = "Jul" Then
        mesEngToSpa = "Jul"
    ElseIf month = "Aug" Then
        mesEngToSpa = "Ago"
    ElseIf month = "Sep" Then
        mesEngToSpa = "Sep"
    ElseIf month = "Oct" Then
        mesEngToSpa = "Oct"
    ElseIf month = "Nov" Then
        mesEngToSpa = "Nov"
    ElseIf month = "Dec" Then
        mesEngToSpa = "Dic"
    End If
End Function

Function activaLibro(ByVal cam As Integer) As Boolean
    Dim libroEncontrado As Boolean
    Dim wkbk As String
    activaLibro = False
    Dim libro As Integer
    For libro = 1 To Workbooks.Count
        If Mid(Workbooks(libro).Name, 1, 1) = "(" Then
            wkbk = Workbooks(libro).Name
            activaLibro = True
        End If
    Next libro
    If activaLibro = False Then
        MsgBox ("No se encontró un libro que empiece con '('")
        End
        Exit Function
    End If
    Workbooks(wkbk).Sheets(cam).Activate
End Function









'   ***** Fragmentos de código ******
'
'    Windows("informe1_12-15_03-4.xls").Activate
'
'    Cells.Find(What:="Odómetro al Partir", After:=ActiveCell, LookIn:= _
'        xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext _
'        , MatchCase:=False, SearchFormat:=False).Activate
'
'
    'método 1
'    Dim fileSys, carpeta, archivo As Object
'    Dim rutaInfCOPEC, camion As String
'    rutaInfCOPEC = "D:\Mis Documentos HD\dinero\transportes El Diamante\informes"
'    camion = "1"
'    Set fileSys = CreateObject("Scripting.FileSystemObject")
'    Set carpeta = fileSys.getfolder(rutaInfCOPEC)
'    For Each archivo In carpeta.Files
'        If Left(archivo.Name, 9) = "informe" & camion & "_" Then
'            Debug.Print (archivo.Name)
'        End If
'    Next

    'método 2
'    Dim rutaInfCOPEC, camion As String
'    rutaInfCOPEC = "D:\Mis Documentos HD\dinero\transportes El Diamante\informes"
'    camion = "1"
'    Dim nombreArchivo As String
'    nombreArchivo = Dir(rutaInfCOPEC & "\informe" & camion & "*.xls*")
'    Do While nombreArchivo <> ""
'        Debug.Print nombreArchivo
'        nombreArchivo = Dir
'    Loop

'Range("A:A, C:D, G:J, L:L, N:O").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("A:E").AutoFit

'    Dim wb As Workbook
'    Dim goalBook As String
'    Dim goalBookSheets As Integer
'    For Each wb In Application.Workbooks
'        If wb.Name Like "(*)*.xlsx" Then
'            goalBook = wb.Name
'        End If
'    Next
'    goalBookSheets = Workbooks(goalBook).Worksheets.Count
'    ActiveSheet.Move After:=Workbooks(goalBook).Worksheets(goalBookSheets)

'    Range("G2").Select
'    Selection.FormulaLocal = "=AÑO(B2)"
'    Selection.Offset(0, 1).Select

'Sub funciones_para_variables_tipo_dato()
'    Debug.Print (Date)
'    Debug.Print (Now)
'    Debug.Print (DateValue("15-Dic 12:34"))     'transforma cadena a solo fecha. No reconoce "Dec"
'    Debug.Print (CDate("15-Dic 12:34"))         'transforma cadena a fecha y hora. No reconoce "Dec"
'    Debug.Print Format("15-Dic 12:34", "mmm/yyyy hh")
'    Debug.Print Application.Text(#12/15/2018 12:34:00 PM#, "[$-409]dd-mmm")
'    Debug.Print (TimeValue("15-Dic 12:34"))
'    Debug.Print (CDate(41035))
'    Debug.Print (DateSerial(2012, 5, 16))
'    https://www.science.co.il/language/Locale-codes.php
'End Sub
