VERSION 5.00
Begin VB.UserControl L2ComboBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox Control 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "L2ComboBox.ctx":0000
      Left            =   0
      List            =   "L2ComboBox.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "L2ComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As L2Form
Public name As String
Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long
Public TTabIndex As Integer, tTabStop As Boolean
Private keylist() As String
Private onClick As String
Private DisableEvents As Boolean
Public Caption As String, Label As String

Private Sub Control_Click()
    If Not DisableEvents Then
        If onClick <> "" Then
            owner.Enabled = False
            owner.owner.DocumentManager.XmlObjectList.item(onClick).XML
            owner.Enabled = True
        End If
    End If
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    Else
    End If
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: height = .height
    End With
End Sub

Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    
    Set elm = XML.createElement("combobox")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("listindex")
    attr.nodeValue = Control.ListIndex
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("text")
    If Control.ListIndex = -1 Then
        attr.nodeValue = ""
    Else
        attr.nodeValue = Control.list(Control.ListIndex)
    End If
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("value")
    If Control.ListIndex = -1 Then
        attr.nodeValue = ""
    Else
        attr.nodeValue = keylist(Control.ListIndex)
    End If
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("listcount")
    attr.nodeValue = Control.ListCount
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("enabled")
    attr.nodeValue = Control.Enabled
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = tVisible
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("tabstop")
    attr.nodeValue = tTabStop
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("tabindex")
    attr.nodeValue = TTabIndex
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Me.Caption
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("label")
    attr.nodeValue = Me.Label
    elm.setAttributeNode attr
    
    Dim i As Long, astr As String, rowelm As IXMLDOMElement
    For i = 0 To Control.ListCount - 1
        astr = keylist(i) & "=" & Control.list(i)
        Set rowelm = XML.createElement("keyvalue")
        rowelm.Text = astr
        elm.appendChild rowelm
    Next i
                    
    Set IXMLDOMElementView = elm
End Function

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    Dim i As Long, oldindex As Long, pos As Long, rowelm As IXMLDOMElement, astr As String
    i = elm.SelectNodes("./keyvalue").length
    oldindex = Control.ListIndex
    DisableEvents = True
    If i > 0 Then
        ReDim keylist(i)
        Control.Clear
        For Each rowelm In elm.SelectNodes("./keyvalue")
            astr = rowelm.Text
            pos = InStr(1, astr, "=", vbTextCompare)
            If pos > 0 Then
                If pos > 1 Then
                    If pos < Len(astr) Then
                        Control.AddItem Right(astr, Len(astr) - pos)
                        keylist(Control.NewIndex) = Left(astr, pos - 1)
                        'Control.ItemData(Control.NewIndex) = Left(astr, pos - 1)
                    ElseIf pos = Len(astr) Then
                        Control.AddItem ""
                        keylist(Control.NewIndex) = Left(astr, pos - 1)
                        'Control.ItemData(Control.NewIndex) = Left(astr, pos - 1)
                    End If
                ElseIf pos = 1 Then
                    If pos < Len(astr) Then
                        Control.AddItem Right(astr, Len(astr) - pos)
                    ElseIf pos = Len(astr) Then
                    End If
                End If
            End If
        Next rowelm
        If oldindex >= -1 And oldindex < Control.ListCount Then
            Control.ListIndex = oldindex
        Else
            Control.ListIndex = 0
        End If
    End If
    
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
            Case "ENABLED"
                If aattr.value = bvFalse Then
                    'Control.Locked = False
                    Enabled = False
                    Control.BackColor = &H8000000F
                ElseIf aattr.value = bvTrue Then
                    'Control.Locked = True
                    Enabled = True
                    Control.BackColor = &HFFFFFF
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
            Case "LISTINDEX"
                Control.ListIndex = aattr.value
            Case "VALUE"
                For i = 0 To Control.ListCount - 1
                    If keylist(i) = aattr.value Then
                        Control.ListIndex = i: Exit For
                    End If
                Next i
            Case "TEXT"
                For i = 0 To Control.ListCount - 1
                    If Control.list(i) = aattr.value Then
                        Control.ListIndex = i: Exit For
                    End If
                Next i
            Case "ONCLICK"
                onClick = aattr.value
            Case "CAPTION"
                Caption = aattr.value
            Case "LABEL"
                Label = aattr.value
        End Select
    Next aattr
    DisableEvents = False
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

