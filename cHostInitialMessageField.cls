VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHostMessageField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public owner As cHostInitialMessage
Public workelement As IXMLDOMElement

Public Property Get Value()
    Value = workelement.Text
End Property

Public Property Let Value(avalue)
    workelement.Text = avalue
End Property

Public Property Get Text() As String
    Text = workelement.Text
End Property

Public Property Let Text(avalue As String)
    workelement.Text = avalue
End Property

Public Property Get AsInteger() As Long
    AsInteger = CLng("0" & Trim(workelement.Text))
End Property

Public Property Let AsInteger(avalue As Long)
    workelement.Text = CStr(avalue)
End Property


Public Property Get AsDouble() As Double
    AsDouble = CDbl(workelement.Text)
End Property

Public Property Let AsDouble(avalue As Double)
    workelement.Text = CStr(avalue)
End Property


Public Function Field(name As String)
    Dim anode As IXMLDOMElement
    Set anode = workelement.selectSingleNode("./field[@name='" & name & "']")
    
    If anode Is Nothing Then
        Set anode = workelement.ownerDocument.createElement("field")
        Dim aattr As IXMLDOMAttribute
        Set aattr = workelement.ownerDocument.createAttribute("name")
        aattr.Value = name
        anode.Attributes.setNamedItem aattr
        workelement.appendChild anode
    Else
    
    End If
    Dim aField As New cHostInitialMessageField
    Set aField.owner = owner
    Set aField.workelement = anode
End Function
