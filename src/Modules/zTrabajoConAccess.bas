Attribute VB_Name = "zTrabajoConAccess"
Option Explicit

Sub arrayCamiones()
    Dim miConexion As Connection
    Dim Rs As Recordset
    Set miConexion = New ADODB.Connection
    With miConexion
        .Provider = "Microsoft.ACE.OLEDB.16.0"
        .ConnectionString = "Data Source=C:\Users\retac\Documents\dinero\transportes El Diamante\DatosInformes.accdb"
    End With
    
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT * FROM Camiones", miConexion, adOpenKeyset, adLockOptimistic, adCmdText
    
        Debug.Print Rs.Fields("Patente")
        
    'cmd_Guardar.Enabled = False
End Sub

Sub arrayCamiones2()
    Dim cn As Object
    Dim datos As Object
    Dim consultaSQL As String
    Dim conexion As String
    Dim cont As Long
    
    Set cn = CreateObject("ADODB.Connection")
    conexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=C:\Users\retac\Documents\dinero\transportes El Diamante\DatosInformes.accdb"
    consultaSQL = "Select * from Camiones;"
    cn.Open conexion
    
    Set datos = cn.Execute(consultaSQL)
    
    cont = 1

    Do While Not datos.EOF
        Debug.Print datos.Fields(2)
        cont = cont + 1
        datos.MoveNext
    Loop
    
    datos.Close
    Set datos = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Sub obtenDato()
    Dim dato As String
    dato = consultarCamion(5, 1)  'camion,(0=id/1=patente/2=abreviacion) ; camion=0 para recibir la cantidad de camiones
    Debug.Print dato
End Sub
