VERSION 5.00
Begin VB.UserControl L2CheckBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox Control 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "L2CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form
Public name As String
Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long
Public TTabIndex As Integer, tTabStop As Boolean

Private Sub Control_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
Dim apos As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: .height = height
    End With
End Sub

Public Property Get value() As Integer
    value = Control.value
End Property

Public Property Let value(invalue As Integer)
    Control.value = invalue
    PropertyChanged "Value"
End Property

Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("checkbox")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Control.Caption
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = Me.tVisible
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("tabstop")
    attr.nodeValue = tTabStop
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("tabindex")
    attr.nodeValue = TTabIndex
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("value")
    attr.nodeValue = Control.value
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("readonly")
    attr.nodeValue = Not Control.Enabled
    elm.setAttributeNode attr
                    
    Set IXMLDOMElementView = elm
End Function

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    For Each aattr In elm.Attributes
        Select Case UCase(aattr.baseName)
            Case "LEFT"
                tLeft = aattr.value
            Case "TOP"
                tTop = aattr.value
            Case "WIDTH"
                tWidth = aattr.value
            Case "HEIGHT"
                tHeight = aattr.value
            Case "CAPTION"
                Control.Caption = aattr.value
            Case "ENABLED"
                If aattr.value = bvTrue Then
                    Enabled = True
                    Control.Enabled = True
                ElseIf aattr.value = bvFalse Then
                    Enabled = False
                    Control.Enabled = False
                End If
            Case "VISIBLE"
                If aattr.value = bvFalse Then
                    tVisible = False
                ElseIf aattr.value = bvTrue Then
                    tVisible = True
                End If
            Case "TABSTOP"
                If aattr.value = bvFalse Then
                    tTabStop = False
                ElseIf aattr.value = bvTrue Then
                    tTabStop = True
                End If
            Case "TABINDEX"
                TTabIndex = aattr.value
            Case "VALUE"
                Control.value = aattr.value
        End Select
    Next aattr
    
End Sub

Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)

    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.value
    
    LoadFromIXMLDOMElement inNode
        
End Sub



Public Sub CleanUp()
    Set owner = Nothing
End Sub

