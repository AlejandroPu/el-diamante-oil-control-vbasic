Attribute VB_Name = "x_datosDeAccess"
Option Explicit

' "camion" corresponde a la columna Id de la base de datos
' Excepto Si "camion" = 0, en que se realiza la consulta CantDeCamiones
' "campo" corresponde a la columna que se consulta
'   columna 1, formato KGRZ-60
'   columna 2, formato K60

Function consultarCamion(camion As Integer, Optional campo As Integer) As String
    Dim cn As Object
    Dim datos As Object
    Dim consultaSQL As String
    Dim conexion As String
    
    Set cn = CreateObject("ADODB.Connection")
    conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=D:\Mis Documentos HD\dinero\transportes El Diamante\DatosInformes.accdb"
    If (camion > 0) Then
        consultaSQL = "Select * from Camiones WHERE id=" & camion & ";"
        cn.Open conexion
        
        Set datos = cn.Execute(consultaSQL)
        consultarCamion = datos.Fields(campo)
    Else
        consultaSQL = "Select * from CantDeCamiones;"
        cn.Open conexion
        
        Set datos = cn.Execute(consultaSQL)
        consultarCamion = datos.Fields(0)
    End If
    
    datos.Close
    Set datos = Nothing
    cn.Close
    Set cn = Nothing
End Function


'   antes usaba este arreglo
'    Dim camiones(24, 2) As String
'    camiones(1, 1) = "KGRZ-60"
'    camiones(2, 1) = "KGRZ-61"
'    camiones(3, 1) = "KGRZ-62"
'    camiones(4, 1) = "KGRZ-63"
'    camiones(5, 1) = "KGRZ-64"
'    camiones(6, 1) = "KGRZ-65"
'    camiones(7, 1) = "KGRZ-66"
'    camiones(8, 1) = "KGRZ-67"
'    camiones(9, 1) = "KGRZ-69"
'    camiones(10, 1) = "KGRZ-75"
'    camiones(11, 1) = "JVJL-93"
'    camiones(12, 1) = "JVJL-94"
'    camiones(13, 1) = "JVJL-95"
'    camiones(14, 1) = "JVJL-96"
'    camiones(15, 1) = "JVJL-97"
'    camiones(16, 1) = "JVJL-98"
'    camiones(17, 1) = "JVJL-99"
'    camiones(18, 1) = "JVJR-55"
'    camiones(19, 1) = "JVJR-56"
'    camiones(20, 1) = "JVJR-57"
'    camiones(21, 1) = "JVJR-58"
'    camiones(22, 1) = "JVJR-59"
'    camiones(23, 1) = "JVJR-60"
'    camiones(24, 1) = "JVJR-61"
'
'    camiones(1, 2) = "K60"
'    camiones(2, 2) = "K61"
'    camiones(3, 2) = "K62"
'    camiones(4, 2) = "K63"
'    camiones(5, 2) = "K64"
'    camiones(6, 2) = "K65"
'    camiones(7, 2) = "K66"
'    camiones(8, 2) = "K67"
'    camiones(9, 2) = "K69"
'    camiones(10, 2) = "K75"
'    camiones(11, 2) = "JL93"
'    camiones(12, 2) = "JL94"
'    camiones(13, 2) = "JL95"
'    camiones(14, 2) = "JL96"
'    camiones(15, 2) = "JL97"
'    camiones(16, 2) = "JL98"
'    camiones(17, 2) = "JL99"
'    camiones(18, 2) = "JR55"
'    camiones(19, 2) = "JR56"
'    camiones(20, 2) = "JR57"
'    camiones(21, 2) = "JR58"
'    camiones(22, 2) = "JR59"
'    camiones(23, 2) = "JR60"
'    camiones(24, 2) = "JR61"
    
