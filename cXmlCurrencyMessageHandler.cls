VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlCurrencyMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public workDocument As MSXML2.DOMDocument30
Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement

'Dim msglist As IXMLDOMNodeList

Dim doc As New MSXML2.DOMDocument30

Public Function LoadXml(inDocument As String) As String

    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXml inDocument
    
    If Trim(workDocument.selectSingleNode("//RESP/RC").Text) <> "0" Then
        LoadXml = inDocument
        Exit Function
    End If
    
    'msglist = workDocument.SelectNodes("//ODATA/STRMSG")
    'If (msglist.length > 0) Then
    If (workDocument.SelectNodes("//ODATA/STRMSG").length > 0) Then
        'Load XMLMessageForm
        Set CurrencyMessageForm.MessageDocument = workDocument
        CurrencyMessageForm.Show vbModal, ActiveL2TrnHandler.activeform
    
'        Set XMLMessageForm.MessageDocument = workDocument
'        XMLMessageForm.Show vbModal, ActiveL2TrnHandler.activeform
        LoadXml = workDocument.XML
    Else
      LoadXml = inDocument
      Exit Function
    End If
    
End Function
