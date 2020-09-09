VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public bNode As IXMLDOMElement
'Private fName As String
'Private fValue

'Public Property Get Name() As String
'    Name = fName
'End Property

'Public Property Let Name(inValue As String)
'    If fName <> "" Then
'        MsgBox "Το Πεδίο " & inValue & " έχει ήδη όνομα."
'    Else
'        fName = inValue
'    End If
'End Property

Public Property Let Value(inValue)
    bNode.Text = inValue
End Property

Public Property Get Value()
    Value = bNode.Text
End Property

'Friend Property Let Node(inValue As Variant)
'    Set bNode = inValue
'End Property

'Friend Property Get Node() As Variant
'    Set Node = bNode
'End Property

'Private Sub Class_Initialize()
'    fName = ""
'End Sub
