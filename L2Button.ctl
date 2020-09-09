VERSION 5.00
Begin VB.UserControl L2Button 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   DefaultCancel   =   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   1995
   Begin VB.CommandButton Control 
      BackColor       =   &H00C0C0C0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "L2Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As L2Form

Public name As String
Public tEnabled As Boolean, tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long, tCancel
Public TTabIndex As Integer, tTabStop As Boolean

Private onClick As String, HotKey As String, HookHidden As Boolean, Cancel As String
Private aonClick As String

Private Sub Control_Click()
   
    If Control.Cancel Then
        Unload owner
    End If
   If onClick = "." Then Exit Sub
   If onClick <> "" Then
        owner.Enabled = False
        owner.KeyPreview = False
        

        Dim aname As String
        aname = onClick
        onClick = "."
        Dim ajob As cXMLDocumentJob
        Set ajob = owner.owner.DocumentManager.XmlObjectList.Item(aname)
        If ajob Is Nothing Then
            MsgBox "Δεν βρέθηκε το job: " & aname, vbCritical, "Λάθος..."
            onClick = aname
            Exit Sub
        End If
        ajob.XML
        
        On Error Resume Next
        If onClick = "." Then
            onClick = aname
        End If
        If Not ajob.exitformflag Then owner.Enabled = True: owner.KeyPreview = True
      
    ElseIf onClick = "" And (HotKey = "ESC" Or Cancel = "-1") Then
    
        Unload owner
    End If
    
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: .height = height
    End With
End Sub

Public Property Get HotKeyValue() As Integer
    Select Case UCase(HotKey)
        Case "ESC": HotKeyValue = vbKeyEscape
        Case "ENTER": HotKeyValue = vbKeyReturn
        Case "F1": HotKeyValue = vbKeyF1
        Case "F2": HotKeyValue = vbKeyF2
        Case "F3": HotKeyValue = vbKeyF3
        Case "F4": HotKeyValue = vbKeyF4
        Case "F5": HotKeyValue = vbKeyF5
        Case "F6": HotKeyValue = vbKeyF6
        Case "F7": HotKeyValue = vbKeyF7
        Case "F8": HotKeyValue = vbKeyF8
        Case "F9": HotKeyValue = vbKeyF9
        Case "F10": HotKeyValue = vbKeyF10
        Case "F11": HotKeyValue = vbKeyF11
        Case "F12": HotKeyValue = vbKeyF12
        Case "CTRL-M": HotKeyValue = 77
        Case "ALT-M": HotKeyValue = 77
        Case "CTRL-C": HotKeyValue = 67
        Case "ALT-C": HotKeyValue = 67
    End Select
End Property

Public Sub Click()
    Control_Click
End Sub

Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)

    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.Value

    LoadFromIXMLDOMElement inNode
End Sub

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    For Each aattr In elm.Attributes
        Select Case UCase(aattr.baseName)
            Case "LEFT"
                tLeft = aattr.Value
            Case "TOP"
                tTop = aattr.Value
            Case "WIDTH"
                tWidth = aattr.Value
            Case "HEIGHT"
                tHeight = aattr.Value
            Case "TABSTOP"
                If UCase(aattr.Value) = bvTrue Then
                    tTabStop = True
                ElseIf UCase(aattr.Value) = bvFalse Then
                    tTabStop = False
                End If
            Case "TABINDEX"
                TTabIndex = aattr.Value
            Case "VISIBLE"
                If UCase(aattr.Value) = bvFalse Then
                    Me.tVisible = False
                ElseIf UCase(aattr.Value) = bvTrue Then
                    Me.tVisible = True
                End If
            Case "CAPTION"
                Control.Caption = aattr.Value
            Case "ENABLED"
                If aattr.Value = bvFalse Then
                    Enabled = False
                    Control.Enabled = False
                    tEnabled = False
                ElseIf aattr.Value = bvTrue Then
                    Enabled = True
                    Control.Enabled = True
                    tEnabled = True
                End If
            Case "ONCLICK"
                onClick = aattr.Value
            Case "HOTKEY"
                HotKey = aattr.Value
            Case "HOOKHIDDEN"
                If aattr.Value = bvFalse Then
                    HookHidden = False
                ElseIf aattr.Value = bvTrue Then
                    HookHidden = True
                End If
            Case "CANCEL"
                HotKey = "ESC"
                Cancel = aattr.Value
'                If aattr.value = bvFalse Then
'                    Control.Cancel = False
'                    tCancel = False
'                ElseIf aattr.value = bvTrue Then
'                    Control.Cancel = True
'                    tCancel = True
'                End If
        End Select
    Next aattr
End Sub

Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("button")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("enabled")
    attr.nodeValue = tEnabled
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = tVisible
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Control.Caption
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("hotkey")
    attr.nodeValue = HotKey
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("onclick")
    attr.nodeValue = onClick
    elm.setAttributeNode attr
                    
    Set IXMLDOMElementView = elm
End Function


Public Sub CleanUp()
    Set owner = Nothing
End Sub

