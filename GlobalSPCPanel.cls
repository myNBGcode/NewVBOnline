VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSPCPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private apanel As SPCPanelX
 
Private Sub Class_Initialize()
    Set apanel = CreateObject("SPCPanelXControl.SPCPanelX")
    apanel.host = cPRINTERSERVER
    apanel.Port = cPrinterPort
End Sub

Public Sub LockPrinter()
    apanel.LockPrinter
End Sub

Public Sub UnLockPrinter()
    apanel.UnLockPrinter
End Sub

Public Function PrintText(Data As String) As Boolean
    PrintText = apanel.PrintText(Data)
End Function

Public Sub ShowServer()
    apanel.ShowServer
End Sub

Public Sub HideServer()
    apanel.HideServer
End Sub
Public Sub StartPrinter()
    apanel.StartPrinter
End Sub
Public Sub StopPrinter()
    apanel.StopPrinter
End Sub

Private Sub Class_Terminate()
    Set apanel = Nothing
End Sub
