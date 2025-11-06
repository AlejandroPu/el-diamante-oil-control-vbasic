Attribute VB_Name = "c_arreglaFuente"
Option Explicit

'Esta macro, arregla los datos fuente descargados
'de la página de copec y la nombre "Fuente"

Sub arreglaFuenteMain()
    'elimina columnas sobrantes y ajusta los anchos
    'Range("A:A, C:D, G:J, L:L, N:O").Select
    Range("A:C, E:K, N:P, R:R, T:U").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:E").AutoFit

    'pega formato deseado desde celda G2
    Call formatoDeseado
    
    'agrega columnas vacias y pega valores correctos
    Call arreglaDatos
    
    'agrega filas vacias entre datos de camiones
    Call agregaFilas
    
    Call nombreHoja
    
    Dim wb As Workbook
    Dim goalBook As String
    Dim goalBookSheets As Integer
    For Each wb In Application.Workbooks
        If wb.Name Like "(*)*.xlsx" Then
            goalBook = wb.Name
        End If
    Next
    goalBookSheets = Workbooks(goalBook).Worksheets.Count
    ActiveSheet.Move After:=Workbooks(goalBook).Worksheets(goalBookSheets)
End Sub

Sub formatoDeseado()
    Range("G2").Select
    Selection.FormulaLocal = "=AÑO(B2)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=MES(B2)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=DIA(B2)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=HORA(C2)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=MINUTO(C2)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=FECHA(G2;H2;I2)"
    Selection.Offset(0, 2).Select
    Selection.FormulaLocal = "=NSHORA(J2;K2;0)"
    Selection.Offset(0, 1).Select
    Selection.FormulaLocal = "=D2+0"
    Selection.Offset(0, 5).Select
    Selection.FormulaLocal = "=E2+0"
    
    Columns("G:T").ColumnWidth = 1
    Columns("G:T").AutoFit
    Range("G2:T2").Copy
    Range("E2").End(xlDown).Select
    Selection.Offset(0, 2).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
End Sub
Sub arreglaDatos()
    ActiveCell.FormulaR1C1 = ""
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("F:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Dim finalRange As String
    finalRange = Range("Y2").End(xlDown).Address
    Range(Range("Q2"), finalRange).Copy
    Range("B2").PasteSpecial xlPasteValues
    Range("B2").PasteSpecial xlPasteFormats
    Columns("B:J").ColumnWidth = 1
    Columns("B:J").AutoFit
    
End Sub

Sub agregaFilas()
    Dim camion As String
    Dim antiFreeze As Integer
    Dim valorFreeze As Integer
    Range("A2").Select
    valorFreeze = 300
    Do
        camion = Selection.Value
        antiFreeze = 0
        Do
            Selection.Offset(1, 0).Select
            antiFreeze = antiFreeze + 1
            If (antiFreeze > valorFreeze) Then
                Exit Do
            End If
        Loop While (Selection.Value = camion)
        Selection.EntireRow.Select
        Selection.Insert
        Selection.End(xlToLeft).Select
        Selection.Offset(1, 0).Select
    Loop While (antiFreeze < valorFreeze)
    Selection.End(xlUp).Select
End Sub

Sub nombreHoja()
    ActiveSheet.Name = "Fuente"
End Sub
