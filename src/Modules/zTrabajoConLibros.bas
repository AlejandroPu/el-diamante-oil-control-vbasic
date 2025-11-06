Attribute VB_Name = "zTrabajoConLibros"
Option Explicit

Public Sub trabajoLibros()

    Const celdaOrigen = "B6"
'    Call abrirLibro 'abrelibro y pega
    Call contarArchivos 'Cuenta archivos en una carpeta
    
'    Application.ScreenUpdating = False  'Evita el parpadeo en pantalla producto del proceso de la macro
'    Application.StatusBar = "Ejecutando"
'    Ejecución de la macro
'    Application.StatusBar = False
'    Application.ScreenUpdating = True   'colocar al final de la macro
    
    'like
'    palabra = "*" & Selection & "*"
'    If Range("B10") Like "palabra" Then
'        Debug.Print ActiveCell.Address
'    End If
    
    
'    Dim wbLibroActual As Workbook
'    Dim wsHojaActual As Worksheet
'    Dim wbLibroNuevo As Workbook
    
'    Set wbLibroActual = Workbooks(ThisWorkbook.Name)
'    Set wsHojaActual = wbLibroActual.ActiveSheet
    'Set wsHojaActual = wbLibroActual.Worksheets("Hoja1")   'Otra forma de hacer este set
    'Set wsHojaActual = wbLibroActual.Worksheets(1)   'Otra forma de hacer este set
    
    'MsgBox wsHojaActual.UsedRange.Address
    'wsHojaActual.UsedRange.AutoFilter field:=5, criteria:="Maule" ' field señala el número de la columna para aplicar el filtro
    'wsHojaActual.UsedRange.AutoFilter 'Al colocar de nuevo, quita el autofiltro
    
    'Buscar última fila con datos en col A
    'algo = wsHojaActual.Range("A" & Rows.Count).End(xlUp).Row
    
    'crear libro nuevo
'    Set wbLibroNuevo = Workbooks.Add
'    wbLibroNuevo.ActiveSheet.Paste
    'Application.CutCopyMode = False     'quita la selección que queda cuando se copia algo
    
'    Windows(wbLibroActual.Name).Activate
    
End Sub

Sub abrirLibro()
    Dim ruta As String
    ruta = "C:\Users\retac\Documents\dinero\transportes El Diamante\(2206) Test Mayo 2018 copia1.xlsx"
    
    Dim wbLibroOrigen As Workbook, _
    wsHojaOrigen As Excel.Worksheet, _
    wbLibroDestino As Workbook, _
    wsHojaDestino As Worksheet, _
    uFila As Integer
    
    Set wbLibroDestino = Workbook(ThisWorkbook.Name)
    Set wbhojadestino = wbLibroDestino.Worksheets("HojaProducto")
    
    Set wbLibroOrigen = Workbooks.Open(ruta)
    Set wsHojaOrigen = wbLibroOrigen.Worksheets("HojaOriginal")
    
    uFila = wsHojaOrigen.Range("A" & Rows.Count).End(xlUp).Row
    wsHojaOrigen.Range("A3:H" & uFila).Copy Destination:=wsHojaDestino.Range("A1")
    
    Workbooks(wslibroorigen.Name).Close savechanges:=False  'No guarda la hoja
    
End Sub

Public Sub contarArchivos()
    Dim cNombreArchivo, cCarpeta As String
    Dim conteo As Integer
    
    cCarpeta = ActiveWorkbook.Path & "\"
    cNombreArchivo = Dir(cCarpeta & "(18" & "*.xl*")
    conteo = 0
    Do While Not (cNombreArchivo = "")
        conteo = conteo + 1
        cNombreArchivo = Dir()  'Al no colocar nada entre parentesis itera usando criterio definido previamente
    Loop

End Sub
