Attribute VB_Name = "i_paso1txt"
Option Explicit

'los datos que se recogerán ya deben estar seleccionados
'antes de ejecutar esta macro

Sub pasoUnoTxtMain()
    Dim camion As Integer
    Dim lastTruck As Integer
    lastTruck = consultarCamion(0)  'función de módulo zTrabajoConAccess

    'Preguntar opcion conLapsos al usuario
    Dim preguntaLapsos As Integer
    Dim fromTruck As Integer
    Dim untilTruck As Integer
    Dim conLapsos As Boolean
    preguntaLapsos = MsgBox("Desea recoger los datos eliminando lapsos vacíos?", vbYesNo + vbInformation + vbDefaultButton2, "Eliminación de lapsos")
    fromTruck = Application.InputBox(prompt:="Ingresar número de camión desde el cual iniciar( Default = 1)", Title:="desde", Default:=1, Type:=1)
    untilTruck = Application.InputBox(prompt:="Ingresar número de camión hasta el cual ejecutar( Default = último camión)", Title:="hasta", Default:=lastTruck, Type:=1)
    
    conLapsos = True
    If (preguntaLapsos = 7) Then
        conLapsos = False
    End If
    
    'Guardar los documentos de texto
    For camion = fromTruck To untilTruck
        Debug.Print camion
        Sheets(consultarCamion(camion, 2)).Select
        Call guardarTxt(camion, conLapsos)
    Next camion
    Sheets(consultarCamion(1, 2)).Select
End Sub

Sub guardarTxt(textoNumero As Integer, conLapso As Boolean)
    Dim ruta As String
    Dim firstRow As Integer
    Dim cantRows As Integer
    Dim cierreMes As Integer
    
    ruta = "D:\Mis Documentos HD\dinero\transportes El Diamante\" & textoNumero & ".txt"
    
    firstRow = (Selection.Row) - 1
    cantRows = Selection.Rows.Count
    'Si cierreMes > 0 hay selección de 2 filas separadas
    cierreMes = filaExtra()     'Número de segunda fila seleccionada
    If (cierreMes > 0) Then
        cantRows = 2
    End If
    
    Open ruta For Output As #1
    Print #1, "p01"
    Dim i As Integer
    Dim filaEnUso As Integer
    Dim finLapso As Boolean
    For i = 1 To cantRows
        filaEnUso = firstRow + i
        If (i = 2) And (cierreMes > 0) Then
            filaEnUso = cierreMes
        End If
        If ((i = 1) Or (finLapso = True) Or (conLapso = False)) Then
            Print #1, WorksheetFunction.Text(Cells(filaEnUso, 2), "DD-MMM") & " " _
                & WorksheetFunction.Text(Cells(filaEnUso, 4), "HH:MM") & " " _
                & Cells(filaEnUso, 5).Value & " " _
                & Cells(filaEnUso, 6).Value
            finLapso = False
        Else
            ' Verificar si celdas F actual y siguiente tienen datos
            Dim celActual As Boolean
            Dim celSiguiente As Boolean
            celActual = Not (Cells(filaEnUso, 6).Value = "")
            celSiguiente = Not (Cells(filaEnUso + 1, 6).Value = "")
            ' ¿ se incluye actual linea o no?
            If (celActual = True) And (celSiguiente = False) Then
                'SÍ se incluye actual si:
                Print #1, WorksheetFunction.Text(Cells(filaEnUso, 2), "DD-MMM") & " " _
                    & WorksheetFunction.Text(Cells(filaEnUso, 4), "HH:MM") & " " _
                    & Cells(filaEnUso, 5).Value & " " _
                    & Cells(filaEnUso, 6).Value
            Else
                'NO se incluye actual
                If (celActual = False) And (celSiguiente = True) Then
                    finLapso = True
                End If
            End If
        End If
    Next i
    Print #1, "fin"
    Close #1
End Sub

'Al seleccionar 2 rangos en filas separadas
'devuelve el número de fila de la segunda
'sino, devuelve 0
Function filaExtra() As Integer
    Dim rangSt As String
    Dim hayFila As Integer
    rangSt = Selection.Address
    hayFila = InStr(rangSt, ",")
    If (hayFila > 0) Then
        filaExtra = Right(rangSt, Len(rangSt) - InStrRev(rangSt, "$"))
    Else
        filaExtra = 0
    End If
End Function
