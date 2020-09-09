VERSION 5.00
Begin VB.UserControl L2Label 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Enabled         =   0   'False
   FillStyle       =   0  'Solid
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Label Control 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "L2Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form
Public name As String
Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long

Private Sub UserControl_GotFocus()
    MsgBox "labelFocus"
End Sub

Private Sub UserControl_Resize()
    Control.Left = 0
    Control.Top = 0
    Control.width = width
    Control.height = height
End Sub

Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)

    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.value
    
    LoadFromIXMLDOMElement inNode
End Sub

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    Enabled = False
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
            Case "VISIBLE"
                If UCase(aattr.value) = bvFalse Then
                    Me.tVisible = False
                ElseIf UCase(aattr.value) = bvTrue Then
                    Me.tVisible = True
                End If
            
            Case "CAPTION"
                Control.Caption = aattr.value
        End Select
    Next aattr
End Sub

Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("label")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = tVisible
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Control.Caption
    elm.setAttributeNode attr
                    
    Set IXMLDOMElementView = elm
End Function



Public Sub CleanUp()
    Set owner = Nothing
End Sub

