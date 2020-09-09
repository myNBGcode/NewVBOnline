VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlESignatureSocketHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function TrySendData(reportDoc As MSXML2.DOMDocument30) As Boolean

    Dim eform As New eSignatureForm

    eform.ReceivePending = True
    eform.Canceled = False

    eform.SendData (reportDoc.XML)

    eform.Show vbModal
    While (eform.ReceivePending)
        DoEvents
    Wend

    If (eform.gError > 0 Or eform.Canceled = True) Then
        TrySendData = False
        eform.gError = 0
        eform.Canceled = False
    Else
        TrySendData = True
    End If

    Unload eform
    Set eform = Nothing

End Function

