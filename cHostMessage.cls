VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHostMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WorkDocument As MSXML2.DOMDocument30

Public Sub Prepare(name As String)
    If WorkDocument Is Nothing Then
        Set WorkDocument = New MSXML2.DOMDocument30
        WorkDocument.appendChild WorkDocument.createElement(name) '"initialmessage"
    End If
End Sub

Public Function Field(name As String) As cHostMessageField
    Dim anode As IXMLDOMElement
    Set anode = WorkDocument.documentElement.selectSingleNode("./field[@name='" & name & "']")
    
    If anode Is Nothing Then
        Set anode = WorkDocument.createElement("field")
        Dim aattr As IXMLDOMAttribute
        Set aattr = WorkDocument.createAttribute("name")
        aattr.Value = name
        anode.Attributes.setNamedItem aattr
        WorkDocument.documentElement.appendChild anode
    Else
    
    End If
    Set Field = New cHostMessageField
    Set Field.owner = Me
    Set Field.workelement = anode
End Function

Public Function SelectNodes(apath) As IXMLDOMNodeList
    Set SelectNodes = WorkDocument.documentElement.SelectNodes(apath)
End Function
