Attribute VB_Name = "a_menuPrincipal"
Option Explicit

' método abreviado: control+p

Sub menuPrincipalMain()
Attribute menuPrincipalMain.VB_Description = "Menú Principal"
Attribute menuPrincipalMain.VB_ProcData.VB_Invoke_Func = "p\n14"
    menuPrincipal.Show
    
'    Dim opcionPrincipal As Variant
'    Dim menu As String
'
'    menu = "Ingrese número de macro a ejecutar:" & Chr(13) & Chr(13) & _
'        "0. abortar ejecucion" & Chr(10) & "1. arreglaFuente" & Chr(10) & _
'        "2. fuenteAInforme" & Chr(10) & "3. lapsosVacios" & Chr(10) & _
'        "4. paso1txt" & Chr(10) & "5. cerrarInforme" & Chr(10) & _
'        "6. FormateaSiguienteMes" & Chr(10) & "der/iz para navegar hojas"
'
'    Do
'        opcionPrincipal = Application.InputBox(prompt:=menu, Title:="Menú Principal", Type:=1)
'
'        Application.Wait (Now + TimeValue("0:00:01"))
'
'        Select Case opcionPrincipal
'            Case 0
'                Exit Do
'            Case 1
'                Call arreglaFuenteMain
'            Case 2
'                Call fuenteAInformeMain
'            Case 3
'                Call lapsosVaciosMain
'            Case 4
'                Call pasoUnoTxtMain
'            Case 5
'                Call cerrarInformeMain
'            Case 6
'                Call FormateaSiguienteMesMain
'            Case 7
'                Exit Sub    'Emergency exit
'        End Select
'    Loop
End Sub

Sub navegaHojas()
    
End Sub
