VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Dim aListener As NPipeListener

Public Sub start_work()
    
    Set aListener = CreateObject("RAuthX.NPipeListener")
    aListener.Start_Listener
    
'    Unload aListener
End Sub

Public Sub stop_work()
    Unload aListener
End Sub

