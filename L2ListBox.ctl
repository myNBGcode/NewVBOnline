VERSION 5.00
Begin VB.UserControl L2ListBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox Control 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "L2ListBox.ctx":0000
      Left            =   480
      List            =   "L2ListBox.ctx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "L2ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form

Public name As String
Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long
Public TTabIndex As Integer, tTabStop As Boolean
Public classes As New Collection
Public Caption As String


Private Sub Control_KeyPress(KeyAscii As Integer)
Dim apos As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub UserControl_Resize()
   Control.Left = 0
   Control.Top = 0
   Control.width = width
   Control.height = height
End Sub


Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("listbox")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("listcount")
    attr.nodeValue = Control.ListCount
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("listindex")
    attr.nodeValue = Control.ListIndex
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
    Set attr = XML.createAttribute("value")
    attr.nodeValue = Control.Text
    elm.Attributes.setNamedItem attr
    
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Me.Caption
    elm.setAttributeNode attr


    Dim i As Integer
    Dim child As IXMLDOMElement
    Dim aclass As L2ListBoxClass
    
    For Each aclass In classes
        Set child = XML.createElement("class")
        elm.appendChild child
        Set attr = XML.createAttribute("name")
        attr.nodeValue = aclass.ClassName
        child.Attributes.setNamedItem attr
        Set attr = XML.createAttribute("formatstring")
        attr.nodeValue = aclass.FormatString
        child.Attributes.setNamedItem attr
        
    Next aclass
    For i = 1 To Control.ListCount
        Set child = XML.createElement("item")
        elm.appendChild child
        Set attr = XML.createAttribute("value")
        attr.nodeValue = Control.list(i - 1)
        child.Attributes.setNamedItem attr
    Next
    
    
    Set IXMLDOMElementView = elm
End Function


Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    Dim classAttr As IXMLDOMAttribute
    Dim formatstrAttr As IXMLDOMAttribute
    Dim aItem As IXMLDOMElement
    Dim aclass As L2ListBoxClass
    
    If elm.SelectNodes("./class").length > 0 Then
        For Each aItem In elm.SelectNodes("./class")
            
            
            Dim Key As String
            Dim FormatString As String
            Key = aItem.Attributes.getNamedItem("name").nodeValue
            FormatString = aItem.Attributes.getNamedItem("formatstring").nodeValue
            
            Set aclass = New L2ListBoxClass
            aclass.ClassName = Key
            aclass.FormatString = FormatString
    
            On Error Resume Next
            classes.Remove (Key)
    
            classes.add aclass, Key
        Next aItem
    End If
    
    If elm.SelectNodes("./item").length > 0 Then
        Dim aIndex As Integer
        aIndex = Control.ListIndex
        Control.Clear
        For Each aItem In elm.SelectNodes("./item")
            Set classAttr = aItem.Attributes.getNamedItem("class")
            Set formatstrAttr = aItem.Attributes.getNamedItem("formatstring")
            Dim countattr As IXMLDOMAttribute, count As String
            Set countattr = aItem.Attributes.getNamedItem("count")
            count = "1":
            If Not (countattr Is Nothing) Then count = countattr.value
            If (Not classAttr Is Nothing And classes.count > 0) Or (Not formatstrAttr Is Nothing) Then
                Dim format As String
                If Not formatstrAttr Is Nothing Then
                    format = formatstrAttr.value
                Else
                    Set aclass = classes.item(classAttr.nodeValue)
                    format = aclass.FormatString
                End If
                
                Dim Row As IXMLDOMElement
                Dim col As IXMLDOMElement
                Dim colitems() As String
                Dim Counter As Long, i As Long
                Counter = aItem.SelectNodes("./part").length
                ReDim colitems(Counter): i = 1
                
                If Counter > 0 Then
                    For Each col In aItem.SelectNodes("./part")
                        colitems(i - 1) = col.Text: i = i + 1
                    Next col
                End If
                
                Dim aValue As String
                aValue = gFormat_(format, colitems)
                If count = "" Then count = "1"
                Dim j As Integer
                For j = 1 To CInt(count)
                    Control.AddItem aValue
                Next j
            
            Else
                
                Set aattr = aItem.Attributes.getNamedItem("value")
                If aattr Is Nothing Then
                    If count = "" Then count = "1"
                    For j = 1 To CInt(count)
                        Control.AddItem ""
                    Next j
                Else
                    If count = "" Then count = "1"
                    For j = 1 To CInt(count)
                        Control.AddItem aattr.value
                    Next j
                End If
            End If
        Next aItem
        If aIndex <= Control.ListCount Then
            Control.ListIndex = aIndex
        Else
            Control.ListIndex = Control.ListCount
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
            Case "CAPTION"
                Caption = aattr.value
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

