VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InformeCopec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private truckId As String
Private startReg As Date
Private endReg As Date
Private odReg As Single
Private consReg As Single

Public Property Get get_truckId() As Variant
    get_truckId = truckId
End Property

Public Property Let let_truckId(ByVal givenTruck As String)
    truckId = givenTruck
End Property

Public Property Get get_startReg() As Variant
    get_startReg = startReg
End Property

Public Property Let let_startReg(ByVal givenStartReg As Variant)
    startReg = givenStartReg
End Property

Public Property Get get_endReg() As Variant
    get_endReg = endReg
End Property

Public Property Let let_endReg(ByVal givenEndReg As Variant)
    endReg = givenEndReg
End Property

Public Property Get get_odReg() As Variant
    get_odReg = odReg
End Property

Public Property Let let_odReg(ByVal givenOdReg As Variant)
    odReg = givenOdReg
End Property

Public Property Get get_consReg() As Variant
    get_consReg = consReg
End Property

Public Property Let let_consReg(ByVal givenConsReg As Variant)
    consReg = givenConsReg
End Property

Public Function allTheData() As String
    allTheData = truckId & " | " & startReg & " | " & endReg & " | " & odReg & " | " & consReg
End Function
