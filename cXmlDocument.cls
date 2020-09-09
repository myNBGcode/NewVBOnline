VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement

Public Function LoadXML(data As String) As String
    Dim doc As MSXML2.DOMDocument30
    doc.LoadXML data
    Set content = doc.documentElement
End Function

Public Function Xml() As String
    If content Is Nothing Then
        Xml = ""
    Else
        Xml = content.Xml
    End If
End Function
