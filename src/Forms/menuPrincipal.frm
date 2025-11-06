VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menuPrincipal 
   Caption         =   "Meú Principal de macros - Informe final EMS"
   ClientHeight    =   4812
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6252
   OleObjectBlob   =   "menuPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "menuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    CommandButton.SetFocus
End Sub

Private Sub UserForm_Click()
    CommandButton.SetFocus
End Sub

'HOJAS
Private Sub Hoja1_Click()
    Worksheets("Hoja1").Select
End Sub
Private Sub JL93_Click()
    Worksheets("JL93").Select
End Sub

Private Sub JL94_Click()
    Worksheets("JL94").Select
End Sub

Private Sub JL95_Click()
    Worksheets("JL95").Select
End Sub

Private Sub JL96_Click()
    Worksheets("JL96").Select
End Sub

Private Sub JL97_Click()
    Worksheets("JL97").Select
End Sub

Private Sub JL98_Click()
    Worksheets("JL98").Select
End Sub

Private Sub JL99_Click()
    Worksheets("JL99").Select
End Sub

Private Sub JR55_Click()
    Worksheets("JR55").Select
End Sub

Private Sub JR56_Click()
    Worksheets("JR56").Select
End Sub

Private Sub JR57_Click()
    Worksheets("JR57").Select
End Sub

Private Sub JR58_Click()
    Worksheets("JR58").Select
End Sub

Private Sub JR59_Click()
    Worksheets("JR59").Select
End Sub

Private Sub JR60_Click()
    Worksheets("JR60").Select
End Sub

Private Sub JR61_Click()
    Worksheets("JR61").Select
End Sub

Private Sub K60_Click()
    Worksheets("k60").Select
    CommandButton.SetFocus
End Sub

Private Sub K61_Click()
    Worksheets("k61").Select
End Sub

Private Sub K62_Click()
    Worksheets("k62").Select
End Sub

Private Sub K63_Click()
    Worksheets("k63").Select
End Sub

Private Sub K64_Click()
    Worksheets("k64").Select
End Sub

Private Sub K65_Click()
    Worksheets("k65").Select
End Sub

Private Sub K66_Click()
    Worksheets("k66").Select
End Sub

Private Sub K67_Click()
    Worksheets("k67").Select
End Sub

Private Sub K69_Click()
    Worksheets("k69").Select
End Sub

Private Sub K75_Click()
    Worksheets("k75").Select
End Sub

'MACROS
Private Sub Macro0Abort_Click()
    Unload Me
End Sub

Private Sub Macro1SourceArrange_Click()
    Unload Me
    Call arreglaFuenteMain
    menuPrincipal.Show
End Sub

Private Sub Macro2SourceToReport_Click()
    Unload Me
    Call fuenteAInformeMain
    menuPrincipal.Show
End Sub

Private Sub Macro3EmptyGaps_Click()
    Unload Me
    Call lapsosVaciosMain
    menuPrincipal.Show
End Sub

Private Sub Macro4Step1Txt_Click()
    Unload Me
    Call pasoUnoTxtMain
    menuPrincipal.Show
End Sub

Private Sub Macro5CloseReport_Click()
    Unload Me
    Call cerrarInformeMain
    menuPrincipal.Show
End Sub

Private Sub Macro6MonthFormat_Click()
    Unload Me
    Call FormateaSiguienteMesMain
    menuPrincipal.Show
End Sub

Private Sub macro7HazInforme_Click()
    Unload Me
    Call registrarInfxCamion    ' módulo yc_informeDeCopec
    menuPrincipal.Show
End Sub

Private Sub macro8CuentaArchivos_Click()
    ' Cuenta archivos de la carpeta con informes EMS guardados
    Unload Me
    Call cuentaInformesGenerados
    menuPrincipal.Show
End Sub

Private Sub PrevSheet_Click()
    Worksheets(ActiveSheet.Index - 1).Select
End Sub

Private Sub PrevSheet_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub PrevSheet_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub NextSheet_Click()
    Worksheets(ActiveSheet.Index + 1).Select
End Sub

Private Sub NextSheet_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub NextSheet_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub upButton_Click()
    ActiveWindow.SmallScroll Up:=3
End Sub

Private Sub upButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub upButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub DownButton_Click()
    ActiveWindow.SmallScroll Down:=3
End Sub

Private Sub DownButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub DownButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CommandButton.SetFocus
End Sub

Private Sub CommandButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            ActiveWindow.SmallScroll Down:=3
        Case vbKeyUp
            ActiveWindow.SmallScroll Up:=3
        Case vbKeyLeft
            'pruebasBox.Value = ActiveSheet.Index
            Worksheets(ActiveSheet.Index - 1).Select
        Case vbKeyRight
            Worksheets(ActiveSheet.Index + 1).Select
        Case vbKey0
            Unload Me
        Case vbKey1
            'Call arreglaFuenteMain
            Call Macro1SourceArrange_Click
        Case vbKey2
            'Call fuenteAInformeMain
            Call Macro2SourceToReport_Click
        Case vbKey3
            'Call lapsosVaciosMain
            Call Macro3EmptyGaps_Click
        Case vbKey4
            Call Macro4Step1Txt_Click
        Case vbKey5
            'Call cerrarInformeMain
            Call Macro5CloseReport_Click
        Case vbKey6
            'Call FormateaSiguienteMesMain
            Call Macro6MonthFormat_Click
        Case vbKey7
            Call macro7HazInforme_Click
        Case vbKey8
            Call macro8CuentaArchivos_Click
    End Select
End Sub
