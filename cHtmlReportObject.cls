VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlReportObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WorkDocument As MSXML2.DOMDocument30
Public workelement As IXMLDOMElement

Private Sub Class_Initialize()
    Set WorkDocument = New MSXML2.DOMDocument30
    WorkDocument.appendChild WorkDocument.createElement("report")
    Set workelement = WorkDocument.documentElement
End Sub

Public Property Get Table(tablename As String)
    Dim anode As IXMLDOMElement
    Set anode = WorkDocument.documentElement.selectSingleNode("./table[@name='" & tablename & "']")
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

