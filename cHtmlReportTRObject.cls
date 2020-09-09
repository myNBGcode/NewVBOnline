VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlReportTRObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public owner As cHtmlReportTableObject
Public workElement As IXMLDOMElement

Public pos

Public Property Get workDocument() As MSXML2.DOMDocument30
    Set workDocument = owner.workDocument
End Property

Private Sub createxml(apos, aaction)
    Dim anode As IXMLDOMElement
    Set anode = owner.workDocument.createElement("tr")
    Dim posattr As IXMLDOMAttribute
    Dim actionattr As IXMLDOMAttribute
    Set posattr = owner.workDocument.createAttribute("pos")
    posattr.value = apos
    anode.Attributes.setNamedItem posattr
    If aaction <> "" Then
        Set actionattr = owner.workDocument.createAttribute("action")
        actionattr.value = aaction
        anode.Attributes.setNamedItem actionattr
    End If
    owner.workElement.appendChild anode
    Set workElement = anode
End Sub

Public Sub InsertLast()
    createxml pos, "insertlast"
End Sub

Public Sub InsertFirst()
    createxml pos, "insertfirst"
End Sub

Public Sub Remove()
    createxml pos, "remove"
End Sub

Public Property Get TD(tdpos)
    'If workElement Is Nothing Then
        
    'End If
    
    createxml pos, ""
    
    Dim anode As IXMLDOMElement
    Set anode = workDocument.createElement("td")
    Dim posattr As IXMLDOMAttribute
    Set posattr = workDocument.createAttribute("pos")
    posattr.value = tdpos
    workElement.appendChild anode
    anode.Attributes.setNamedItem posattr
    
    Dim atd As New cHtmlReportTDObject
    Set atd.owner = Me
    Set atd.workElement = anode
    
    Set TD = atd
End Property
