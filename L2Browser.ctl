VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.UserControl L2Browser 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin SHDocVwCtl.WebBrowser Control 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "L2Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form

Public name As String
Public tEnabled As Boolean, tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long, tCancel
Public urlHeader As String, urlParams As String

Public PageHeaderFlag As Boolean, PageFooterFlag As Boolean
Public PageHeader As String, PageFooter As String, Pagemargin_left As String, Pagemargin_right As String, Pagemargin_top As String, Pagemargin_bottom As String
Public Pageorientation As String

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: .height = height
    End With
End Sub

Private Sub PrintPage()
    
    Dim aregistry As New cRegistry
    Dim tmpval As String
    If PageHeaderFlag Or PageFooterFlag Or Pagemargin_left <> "" Or Pagemargin_right <> "" Or Pagemargin_top <> "" Or Pagemargin_bottom <> "" Then
        aregistry.ClassKey = HKEY_CURRENT_USER
        aregistry.SectionKey = "Software\Microsoft\Internet Explorer\PageSetup"
        If Not aregistry.KeyExists Then
            aregistry.CreateKey
        End If
        
        If Not PageHeaderFlag Then PageHeader = ""
        Call aregistry.SetKeyValueStr("header", PageHeader)
        
        If Not PageFooterFlag Then PageFooter = ""
        Call aregistry.SetKeyValueStr("footer", PageFooter)
        
        If Trim(Pagemargin_left) = "" Then Pagemargin_left = "0.750000"
        Call aregistry.SetKeyValueStr("margin_left", Pagemargin_left)
        
        If Trim(Pagemargin_right) = "" Then Pagemargin_right = "0.750000"
        Call aregistry.SetKeyValueStr("margin_right", Pagemargin_right)
        
        If Trim(Pagemargin_top) = "" Then Pagemargin_top = "0.750000"
        Call aregistry.SetKeyValueStr("margin_top", Pagemargin_top)
        
        If Trim(Pagemargin_bottom) = "" Then Pagemargin_bottom = "0.750000"
        Call aregistry.SetKeyValueStr("margin_bottom", Pagemargin_bottom)
    
    End If
    
    Do
        DoEvents
    Loop Until Control.ReadyState = READYSTATE_COMPLETE And Control.Busy = False
    
    Dim eQuery As OLECMDF
    On Error Resume Next
    eQuery = Control.QueryStatusWB(OLECMDID_PRINT)
    If Err.number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            'Control.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT, Null, Null
            Control.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT, 2, Null '2=PRINT_WAITFORCOMPLETION
        End If
    End If
    
'    Do
'        DoEvents
'    Loop Until Control.ReadyState = READYSTATE_COMPLETE And Control.Busy = False And Control.QueryStatusWB(OLECMDID_PRINT) = OLECMDF_SUPPORTED + OLECMDF_ENABLED
'    On Error Resume Next
'    Control.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT, Null, Null
    Set aregistry = Nothing

End Sub

Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)

    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.value

    LoadFromIXMLDOMElement inNode
End Sub

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute, astr As String, urlchanged As Boolean
    urlchanged = False
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
            Case "ENABLED"
                If aattr.value = bvFalse Then
                    Enabled = False
                    'Control.Enabled = False
                    tEnabled = False
                ElseIf aattr.value = bvTrue Then
                    Enabled = True
                    'Control.Enabled = True
                    tEnabled = True
                End If
            Case "HEADER"
                PageHeaderFlag = True: PageHeader = aattr.value
            Case "FOOTER"
                PageFooterFlag = True: PageFooter = aattr.value
            Case "MARGIN_LEFT"
                Pagemargin_left = aattr.value
            Case "MARGIN_RIGHT"
                Pagemargin_right = aattr.value
            Case "MARGIN_TOP"
                Pagemargin_top = aattr.value
            Case "MARGIN_BOTTOM"
                Pagemargin_bottom = aattr.value
            Case "URL"
                astr = Replace(aattr.value, "&amp;", "&")
                Control.navigate astr
            Case "URLHEADER"
                If Trim(aattr.value) <> "" Then
                    urlHeader = Replace(aattr.value, "&amp;", "&")
                    urlchanged = True
                End If
            Case "URLPARAMS"
                If Trim(aattr.value) <> "" Then
                    urlParams = Replace(aattr.value, "&amp;", "&")
                    urlchanged = True
                End If
            
            Case "ACTION"
                astr = UCase(aattr.value)
                If astr = "PRINT" Then
                    PrintPage
                End If
        End Select
    Next aattr
    If urlchanged Then
        On Error GoTo errInvalidWebLink
        If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
            astr = WebLinks(UCase("EDUC" & urlHeader))
        ElseIf Left(Right(WorkEnvironment_, 8), 4) = "PROD" Then
            astr = WebLinks(UCase("PROD" & urlHeader))
        Else
            astr = ""
        End If
        On Error GoTo 0
        Control.navigate astr & urlParams
    End If
    Exit Sub
errInvalidWebLink:
    LogMsgbox "Λάθος Παράμετροι URL (" & urlHeader & ") " & Err.number & " - " & Err.description, vbCritical, "Λάθος..."
    Exit Sub
End Sub

Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("browser")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("enabled")
    attr.nodeValue = tEnabled
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = tVisible
    elm.setAttributeNode attr
    If PageHeaderFlag Then
        Set attr = XML.createAttribute("header")
        attr.nodeValue = PageHeader
        elm.setAttributeNode attr
    End If
    If PageFooterFlag Then
        Set attr = XML.createAttribute("footer")
        attr.nodeValue = PageFooter
        elm.setAttributeNode attr
    End If
    If Pagemargin_left <> "" Then
        Set attr = XML.createAttribute("margin_left")
        attr.nodeValue = Pagemargin_left
        elm.setAttributeNode attr
    End If
    If Pagemargin_right <> "" Then
        Set attr = XML.createAttribute("margin_right")
        attr.nodeValue = Pagemargin_right
        elm.setAttributeNode attr
    End If
    If Pagemargin_top <> "" Then
        Set attr = XML.createAttribute("margin_top")
        attr.nodeValue = Pagemargin_top
        elm.setAttributeNode attr
    End If
    If Pagemargin_bottom <> "" Then
        Set attr = XML.createAttribute("margin_bottom")
        attr.nodeValue = Pagemargin_bottom
        elm.setAttributeNode attr
    End If
    Set attr = XML.createAttribute("url")
    attr.nodeValue = Replace(Control.LocationURL, "&", "&amp;")
    elm.setAttributeNode attr
    If urlHeader <> "" Then
        Set attr = XML.createAttribute("urlheader")
        attr.nodeValue = Replace(urlHeader, "&", "&amp;")
        elm.setAttributeNode attr
    End If
    If urlParams <> "" Then
        Set attr = XML.createAttribute("urlparams")
        attr.nodeValue = Replace(urlParams, "&", "&amp;")
        elm.setAttributeNode attr
    End If
    
    Set IXMLDOMElementView = elm
End Function

Public Sub CleanUp()
    
    On Error Resume Next
    
    'Control.Quit
    Control.ExecWB OLECMDID_CLOSE, OLECMDEXECOPT_DONTPROMPTUSER
    Set Control.Document = Nothing
    'Unload Control
    
    Set owner = Nothing
End Sub

