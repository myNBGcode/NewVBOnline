VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlReportTDObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public owner As cHtmlReportTRObject
Public workelement As IXMLDOMElement

Public Property Get WorkDocument() As MSXML2.DOMDocument30
    Set WorkDocument = owner.WorkDocument
End Property

Public Property Let Value(avalue)
    workelement.Text = avalue
End Property

Public Property Get Table(tablename As String)
    Dim anode As IXMLDOMElement
    Set anode = workelement.selectSingleNode("./table[@name='" & tablename & "']")
    If anode Is Nothing Then
        
        Set anode = WorkDocument.createElement("table")
        Dim nameattr As IXMLDOMAttribute
        Set nameattr = WorkDocument.createAttribute("name")
        nameattr.Value = tablename
        workelement.appendChild anode
        anode.Attributes.setNamedItem nameattr
    End If
        
    Dim atable As New cHtmlReportTableObject
    Set atable.owner = Me
    Set atable.workelement = anode
    Set Table = atable
End Property


