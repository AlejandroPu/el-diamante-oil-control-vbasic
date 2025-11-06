Attribute VB_Name = "w_archivosCarpetas"
Option Explicit

'es necesario cerrar la columna F para el mes siguiente
'si no hay datos, tal vez no tiene sentido cerrar el mes

Public estaVariableCIg As Integer

Sub cuentaInformesGenerados()
    Dim xFolder As String
    Dim xPath As String
    Dim xCount As Long
    Dim xFile As String
    Dim nombre, nombreAnterior As String
    xPath = "D:\Mis Documentos HD\dinero\transportes El Diamante\informes\*.xls"
    xFile = Dir(xPath)
    
    xCount = 0
    nombreAnterior = Left(xFile, InStr(xFile, "_") - 1)
    Do While xFile <> ""
        nombre = Left(xFile, InStr(xFile, "_") - 1)
        If nombre <> nombreAnterior Then
            Debug.Print (nombreAnterior & " = " & xCount)
            xCount = 0
        End If
        nombreAnterior = nombre
        xCount = xCount + 1
        xFile = Dir()
    Loop
    Debug.Print (nombreAnterior & " = " & xCount)
    MsgBox "Revisar registro consola vba"
End Sub
'Function estaFuncion() As Integer
'    Dim rowReview As Integer
'    mesActual = month(Range("B" & ultimaFilaSigMes).End(xlUp).Value)
'    mesSiguiente = month(Range("B" & lastRowSheet).End(xlUp).Value)
'    Range("B" & lastRowSheet).End(xlUp).Select
'    For rowReview = ultimaFilaSigMes To 9 Step -1
'        Selection.Offset(-1, 0).Select
'        'Debug.Print (Month(Selection) & " = " & mesActual & " ?")
'        If (month(Selection) = mesActual) Then
'            buscaUltFilMesAct = Selection.Row
'            Exit For
'        End If
'    Next rowReview
'End Function
