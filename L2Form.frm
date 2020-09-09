VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form L2Form 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   8760
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   7455
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNUItem 
      Caption         =   "MNUItem"
      Visible         =   0   'False
      Begin VB.Menu MnuSub1 
         Caption         =   "MnuSub1"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem2 
      Caption         =   "MNUItem2"
      Visible         =   0   'False
      Begin VB.Menu MnuSub2 
         Caption         =   "MnuSub2"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem3 
      Caption         =   "MNUItem3"
      Visible         =   0   'False
      Begin VB.Menu MnuSub3 
         Caption         =   "MnuSub3"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem4 
      Caption         =   "MNUItem4"
      Visible         =   0   'False
      Begin VB.Menu MnuSub4 
         Caption         =   "MnuSub4"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem5 
      Caption         =   "MNUItem5"
      Visible         =   0   'False
      Begin VB.Menu MnuSub5 
         Caption         =   "MnuSub5"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem6 
      Caption         =   "MNUItem6"
      Visible         =   0   'False
      Begin VB.Menu MnuSub6 
         Caption         =   "MnuSub6"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem7 
      Caption         =   "MNUItem7"
      Visible         =   0   'False
      Begin VB.Menu MnuSub7 
         Caption         =   "MnuSub7"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem8 
      Caption         =   "MNUItem8"
      Visible         =   0   'False
      Begin VB.Menu MnuSub8 
         Caption         =   "MnuSub8"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem9 
      Caption         =   "MNUItem9"
      Visible         =   0   'False
      Begin VB.Menu MnuSub9 
         Caption         =   "MnuSub9"
         Index           =   0
      End
   End
End
Attribute VB_Name = "L2Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public owner As L2TrnHandler
Public LocalDocuments As New Collection

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
    WriteStatusMessage sMessage
End Sub

Public Sub WriteStatusMessage(Message As String)
    StatusBar.SimpleText = Message
End Sub

Private Sub SetActiveControl(controlname As String)
    On Error Resume Next
    Controls(controlname).SetFocus
    On Error GoTo 0
End Sub

Private Sub RefreshControl(Control)
    ScaleMode = vbPixels
    With Control
        .Left = .tLeft
        .Top = .tTop
        .width = .tWidth
        .height = .tHeight
        
        .Visible = .tVisible
        If Not (TypeOf Control Is shine.L2Label Or TypeOf Control Is shine.L2Browser) Then
            .TabStop = .tTabStop
            .TabIndex = .TTabIndex
        End If
    End With
End Sub

Private Sub Refreshform()
    Dim Control
    For Each Control In Controls
        If TypeOf Control Is shine.L2Label Then
            RefreshControl Control
        ElseIf TypeOf Control Is shine.L2TextBox Then
            RefreshControl Control
            Control.setBackColor
        ElseIf TypeOf Control Is shine.l2button Then
            RefreshControl Control
            Control.Cancel = Control.tCancel
        ElseIf TypeOf Control Is shine.L2ComboBox Then
            RefreshControl Control
        ElseIf TypeOf Control Is shine.L2CheckBox Then
            RefreshControl Control
        ElseIf TypeOf Control Is shine.L2ListBox Then
            RefreshControl Control
        ElseIf TypeOf Control Is shine.L2Grid Then
            RefreshControl Control
        ElseIf TypeOf Control Is shine.L2Browser Then
            RefreshControl Control
        End If
    Next Control
End Sub

Private Function FindControl(controltype As String, controlname As String)
    On Error GoTo ErrorPos
    Set FindControl = Controls(controlname)
    Exit Function
ErrorPos:
    Set FindControl = Controls.add(controltype, controlname)
End Function

Private Function FindLocalDocument(documentname As String)
    On Error GoTo ErrorPos
    Set FindLocalDocument = LocalDocuments(documentname)
    Exit Function
ErrorPos:
    Set FindLocalDocument = Nothing
End Function

Public Sub LoadXML(invalue As String)
    Dim workDocument As New MSXML2.DOMDocument30
    workDocument.LoadXML invalue
    Dim aLocalDocument As MSXML2.DOMDocument30
    Dim aL2TextBox As L2TextBox
    Dim aL2button As l2button
    Dim aL2combobox As L2ComboBox
    Dim aL2checkbox As L2CheckBox
    Dim aL2listbox As L2ListBox
    Dim aL2grid As L2Grid
    Dim aL2Label As L2Label
    Dim aL2Browser As L2Browser
    
    
    Dim elm As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    
    Dim cenabled As Boolean
    cenabled = Enabled
    Enabled = True
    For Each elm In workDocument.SelectNodes("//*")
        If UCase(elm.baseName) = UCase("textbox") Then
            Set aL2TextBox = FindControl("Shine.L2TextBox", elm.getAttribute("name"))
            aL2TextBox.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("button") Then
            Set aL2button = FindControl("Shine.L2button", elm.getAttribute("name"))
            aL2button.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("combobox") Then
            Set aL2combobox = FindControl("Shine.L2combobox", elm.getAttribute("name"))
            aL2combobox.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("checkbox") Then
            Set aL2checkbox = FindControl("Shine.L2checkbox", elm.getAttribute("name"))
            aL2checkbox.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("listbox") Then
            Set aL2listbox = FindControl("Shine.L2listbox", elm.getAttribute("name"))
            aL2listbox.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("grid") Then
            Set aL2grid = FindControl("Shine.L2grid", elm.getAttribute("name"))
            aL2grid.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("label") Then
            Set aL2Label = FindControl("Shine.L2Label", elm.getAttribute("name"))
            aL2Label.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("browser") Then
            Set aL2Browser = FindControl("Shine.L2Browser", elm.getAttribute("name"))
            aL2Browser.CreateFromIXMLDOMElement Me, elm
        ElseIf UCase(elm.baseName) = UCase("menu") Then
            AppendXMLMenu elm
        ElseIf UCase(elm.baseName) = UCase("localdocument") Then
            Set aLocalDocument = FindLocalDocument(elm.getAttribute("name"))
            If aLocalDocument Is Nothing Then
                Set aLocalDocument = New MSXML2.DOMDocument30
                aLocalDocument.LoadXML elm.XML
                LocalDocuments.add aLocalDocument, elm.getAttribute("name")
            Else
                aLocalDocument.LoadXML elm.XML
            End If
        ElseIf UCase(elm.baseName) = UCase("statusmessage") Then
            StatusBar.SimpleText = elm.Text
        End If
    Next elm
    Refreshform
    
    If workDocument.selectSingleNode("//form|//formupdate") Is Nothing Then
    Else
        For Each aattr In workDocument.selectSingleNode("//form|//formupdate").Attributes
            If UCase(aattr.baseName) = "ACTIVECONTROL" Then
                SetActiveControl aattr.Text
            ElseIf UCase(aattr.baseName) = "CAPTION" Then
                Caption = aattr.Text & " L2 Ver. "
            ElseIf UCase(aattr.baseName) = "SCROLL" Then
                If (aattr.Text = "true") Then
                   Me.VScroll1.Visible = True
                   Me.VScroll1.Enabled = True
                End If
'            ElseIf UCase(aattr.baseName) = "ENABLED" Then
'                cenabled = aattr.Text
            End If
        Next aattr
    End If
    Enabled = cenabled
End Sub

Public Function XML() As String
    XML = XMLFormView.XML
End Function

Public Function XMLFormView() As MSXML2.DOMDocument30
Dim xmlTrn As MSXML2.DOMDocument30
Dim elm As IXMLDOMElement
Dim i As Integer

Set xmlTrn = New MSXML2.DOMDocument30

xmlTrn.appendChild CreateXMLNode(xmlTrn, "", "TRN")

Dim Control
For Each Control In Controls
    If TypeOf Control Is shine.L2Label _
    Or TypeOf Control Is shine.L2TextBox _
    Or TypeOf Control Is shine.l2button _
    Or TypeOf Control Is shine.L2ComboBox _
    Or TypeOf Control Is shine.L2CheckBox _
    Or TypeOf Control Is shine.L2ListBox _
    Or TypeOf Control Is shine.L2Grid _
    Or TypeOf Control Is shine.L2Browser Then
        xmlTrn.documentElement.appendChild Control.IXMLDOMElementView
    End If
    
Next Control

Dim aLocalDocument As MSXML2.DOMDocument30
For Each aLocalDocument In LocalDocuments
    If aLocalDocument.documentElement Is Nothing Then
    Else
        xmlTrn.documentElement.appendChild aLocalDocument.documentElement.cloneNode(True)
    End If
Next aLocalDocument

Dim astr As String
astr = xmlTrn.XML
astr = Replace(astr, "xmlns=""""", "")
xmlTrn.LoadXML astr

Set XMLFormView = xmlTrn

End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Control, TControl, foundflag As Boolean
    If Me.Enabled Then
    
        foundflag = False
        Select Case KeyCode
            Case vbKeyEscape, vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, _
                vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12:
                For Each Control In Controls
                    If TypeOf Control Is shine.l2button Then
                        If Control.tEnabled Then
                            If Control.HotKeyValue = KeyCode Then
                                For Each TControl In Controls
                                    If TypeOf TControl Is shine.L2TextBox Then
                                        TControl.FinalizeEdit
                                    End If
                                Next
                                Control.Click: KeyCode = 0: foundflag = True: Exit For
                            End If
                        End If
                    End If
                Next Control
        End Select
        If KeyCode = vbKeyF10 Then
            KeyCode = 0
'            If cNewJournalType = False Then
'                eJournalFrm.Show vbModal, Nothing
'            Else
                Dim aTRNHandler As New L2TrnHandler
                aTRNHandler.ExecuteForm "9989"
                aTRNHandler.CleanUp
                Set aTRNHandler = Nothing
'            End If
        ElseIf KeyCode = vbKeyF11 Then
            Dim bTRNHandler As New L2TrnHandler
            bTRNHandler.ExecuteForm "9747"
            bTRNHandler.CleanUp
            Set bTRNHandler = Nothing
        ElseIf KeyCode = 76 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-l
            If SessID < 99 Then
                SessID = SessID + 1
            Else
                SessID = 1
            End If
        
            KeyCode = 0
            Dim aSelectFrm As New SelectTRNFrm
            aSelectFrm.Show vbModal, Me
            Set aSelectFrm = Nothing
        
            If SessID > 1 Then
                SessID = SessID - 1
            End If
        ElseIf KeyCode = 65 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-a
            KeyCode = 0
            Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = GenWorkForm.AppBuffers
            BufferViewer.Show vbModal, Me
            Unload BufferViewer
        ElseIf KeyCode = 66 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-b
            KeyCode = 0
            Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = owner.DocumentManager.TrnBuffers
            BufferViewer.Show vbModal, Me
            Unload BufferViewer
        ElseIf KeyCode = 67 Then
            If ((Shift And vbAltMask) > 0) Then 'alt-c  defghijklm
                
                For Each Control In Controls
                    If TypeOf Control Is shine.l2button Then
                        If Control.tEnabled Then
                            If Control.HotKeyValue = KeyCode Then
                                Control.Click: KeyCode = 0: foundflag = True: Exit For
                            End If
                        End If
                    End If
                Next Control
            End If
        ElseIf KeyCode = 77 Then
            If ((Shift And vbAltMask) > 0) Then 'alt-m
                
                For Each Control In Controls
                    If TypeOf Control Is shine.l2button Then
                        If Control.tEnabled Then
                            If Control.HotKeyValue = KeyCode Then
                                Control.Click: KeyCode = 0: foundflag = True: Exit For
                            End If
                        End If
                    End If
                Next Control
            End If
        ElseIf KeyCode = 13 Then
            For Each TControl In Controls
                If TypeOf TControl Is shine.L2TextBox Then
                   TControl.FinalizeEdit
                End If
            Next
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Control
    If Me.Enabled Then
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn:
                For Each Control In Controls
                    If TypeOf Control Is shine.l2button Then
                        If Control.tEnabled Then
                            If Control.HotKeyValue = KeyAscii Then
                                Control.Click: KeyAscii = 0: Exit For
                            End If
                        End If
                    End If
                Next Control
        End Select
     End If
End Sub

Private Sub Form_Load()
    Top = GenWorkForm.Top
    Left = GenWorkForm.Left
    width = GenWorkForm.width
    height = GenWorkForm.height
    
    With VScroll1
      .Top = 0
      .Left = Me.ScaleLeft + Me.ScaleWidth - 375
      .min = 0
      .max = 9000 - Me.height
      .SmallChange = Screen.TwipsPerPixelX * 10
      .LargeChange = .SmallChange
      .height = Me.ScaleHeight
      '.Enabled = (Picture1.ScaleHeight <= Picture2.ScaleHeight)
      .ZOrder 0
   End With
    Me.VScroll1.Enabled = False
    Me.VScroll1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set owner.Result = New MSXML2.DOMDocument30
    owner.Result.LoadXML XML
    
    Dim Control
    For Each Control In Controls
        If TypeOf Control Is shine.L2Label _
        Or TypeOf Control Is shine.L2TextBox _
        Or TypeOf Control Is shine.l2button _
        Or TypeOf Control Is shine.L2ComboBox _
        Or TypeOf Control Is shine.L2CheckBox _
        Or TypeOf Control Is shine.L2ListBox _
        Or TypeOf Control Is shine.L2Grid _
        Or TypeOf Control Is shine.L2Browser Then
            Control.CleanUp
            
            Controls.Remove Control
            Set Control = Nothing
        End If
    Next Control
    
    If LocalDocuments Is Nothing Then
    Else
        While LocalDocuments.Count > 0
            LocalDocuments.Remove 1
        Wend
    End If
    Set LocalDocuments = Nothing
    Set owner = Nothing
End Sub

Private Sub AppendXMLMenu(menuNode As IXMLDOMNode)
    Dim itemlist As IXMLDOMNodeList
    Dim itemattr_c As IXMLDOMAttribute
    Dim itemattr_v As IXMLDOMAttribute
    Dim itemattr_e As IXMLDOMAttribute
    Dim subattr_c As IXMLDOMAttribute
    Dim subattr_v As IXMLDOMAttribute
    Dim subattr_e As IXMLDOMAttribute
    Dim subattr_s As IXMLDOMAttribute

    Dim i As Integer
    Dim j As Integer
    Dim subitem As IXMLDOMNode
    Set itemlist = menuNode.SelectNodes("//menuitem")

    If itemlist.length > 0 Then
       For i = 0 To itemlist.length - 1
           Set itemattr_c = itemlist.Item(i).Attributes.getNamedItem("caption")
           Set itemattr_v = itemlist.Item(i).Attributes.getNamedItem("visible")
           Set itemattr_e = itemlist.Item(i).Attributes.getNamedItem("enabled")
           Select Case i
               Case 0:
               If Not (itemattr_c Is Nothing) Then MNUItem.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem.Enabled = itemattr_e.Text
               Case 1:
               If Not (itemattr_c Is Nothing) Then MNUItem2.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem2.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem2.Enabled = itemattr_e.Text
               Case 2:
               If Not (itemattr_c Is Nothing) Then MNUItem3.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem3.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem3.Enabled = itemattr_e.Text
               Case 3:
               If Not (itemattr_c Is Nothing) Then MNUItem4.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem4.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem4.Enabled = itemattr_e.Text
               Case 4:
               If Not (itemattr_c Is Nothing) Then MNUItem5.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem5.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem5.Enabled = itemattr_e.Text
               Case 5:
               If Not (itemattr_c Is Nothing) Then MNUItem6.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem6.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem6.Enabled = itemattr_e.Text
               Case 6:
               If Not (itemattr_c Is Nothing) Then MNUItem7.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem7.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem7.Enabled = itemattr_e.Text
               Case 7:
               If Not (itemattr_c Is Nothing) Then MNUItem8.Caption = itemattr_c.Text
               If Not (itemattr_v Is Nothing) Then MNUItem8.Visible = itemattr_v.Text
               If Not (itemattr_e Is Nothing) Then MNUItem8.Enabled = itemattr_e.Text
           End Select

           For j = 0 To itemlist(i).childNodes.length - 1
              Set subitem = itemlist(i).childNodes(j)
              If subitem.nodeType = NODE_ELEMENT Then
                 If j = 0 Then
                 Else
                    Select Case i
                      Case 0:
                            If j > MNUSub1.LBound And j < MNUSub1.UBound Then
                                If Me.Controls(MNUSub1(j).name) Is Nothing Then Load MNUSub1(j)
                            Else
                                If MNUSub1.UBound + 1 <= j Then Load MNUSub1(MNUSub1.UBound + 1)
                            End If
                      Case 1:
                            If j > MNUSub2.LBound And j < MNUSub2.UBound Then
                                If Me.Controls(MNUSub2(j).name) Is Nothing Then Load MNUSub2(j)
                            Else
                                If MNUSub2.UBound + 1 <= j Then Load MNUSub2(MNUSub2.UBound + 1)
                            End If
                      Case 2:
                            If j > MNUSub3.LBound And j < MNUSub3.UBound Then
                                If Me.Controls(MNUSub3(j).name) Is Nothing Then Load MNUSub3(j)
                            Else
                                If MNUSub3.UBound + 1 <= j Then Load MNUSub3(MNUSub3.UBound + 1)
                            End If
                      Case 3:
                            If j > MNUSub4.LBound And j < MNUSub4.UBound Then
                                If Me.Controls(MNUSub4(j).name) Is Nothing Then Load MNUSub4(j)
                            Else
                                If MNUSub4.UBound + 1 <= j Then Load MNUSub4(MNUSub4.UBound + 1)
                            End If
                      Case 4:
                            If j > MNUSub5.LBound And j < MNUSub5.UBound Then
                                If Me.Controls(MNUSub5(j).name) Is Nothing Then Load MNUSub5(j)
                            Else
                                If MNUSub5.UBound + 1 <= j Then Load MNUSub5(MNUSub5.UBound + 1)
                            End If
                      Case 5:
                            If j > MNUSub6.LBound And j < MNUSub6.UBound Then
                                If Me.Controls(MNUSub6(j).name) Is Nothing Then Load MNUSub6(j)
                            Else
                                If MNUSub6.UBound + 1 <= j Then Load MNUSub6(MNUSub6.UBound + 1)
                            End If
                      Case 6:
                            If j > MNUSub7.LBound And j < MNUSub7.UBound Then
                                If Me.Controls(MNUSub7(j).name) Is Nothing Then Load MNUSub7(j)
                            Else
                                If MNUSub7.UBound + 1 <= j Then Load MNUSub7(MNUSub7.UBound + 1)
                            End If
                      Case 7:
                            If j > MNUSub8.LBound And j < MNUSub8.UBound Then
                                If Me.Controls(MNUSub8(j).name) Is Nothing Then Load MNUSub8(j)
                            Else
                                If MNUSub8.UBound + 1 <= j Then Load MNUSub8(MNUSub8.UBound + 1)
                            End If
                    End Select
                 End If
                 Set subattr_c = subitem.Attributes.getNamedItem("caption")
                 Set subattr_v = subitem.Attributes.getNamedItem("visible")
                 Set subattr_e = subitem.Attributes.getNamedItem("enabled")
                 Set subattr_s = subitem.Attributes.getNamedItem("onselect")
                 Select Case i
                   Case 0:
                   If Not (subattr_c Is Nothing) Then MNUSub1(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub1(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub1(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub1(j).Tag = subattr_s.Text
                   Case 1:
                   If Not (subattr_c Is Nothing) Then MNUSub2(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub2(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub2(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub2(j).Tag = subattr_s.Text
                   Case 2:
                   If Not (subattr_c Is Nothing) Then MNUSub3(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub3(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub3(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub3(j).Tag = subattr_s.Text
                   Case 3:
                   If Not (subattr_c Is Nothing) Then MNUSub4(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub4(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub4(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub4(j).Tag = subattr_s.Text
                   Case 4:
                   If Not (subattr_c Is Nothing) Then MNUSub5(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub5(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub5(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub5(j).Tag = subattr_s.Text
                   Case 5:
                   If Not (subattr_c Is Nothing) Then MNUSub6(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub6(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub6(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub6(j).Tag = subattr_s.Text
                   Case 6:
                   If Not (subattr_c Is Nothing) Then MNUSub7(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub7(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub7(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub7(j).Tag = subattr_s.Text
                   Case 7:
                   If Not (subattr_c Is Nothing) Then MNUSub8(j).Caption = subattr_c.Text
                   If Not (subattr_v Is Nothing) Then MNUSub8(j).Visible = subattr_v.Text
                   If Not (subattr_e Is Nothing) Then MNUSub8(j).Enabled = subattr_e.Text
                   If Not (subattr_s Is Nothing) Then MNUSub8(j).Tag = subattr_s.Text
                 End Select
              End If
           Next j
       Next i
    End If
End Sub


Private Sub HandleEvent(sender As String)
   owner.DocumentManager.XmlObjectList.Item(sender).XML
End Sub

Private Sub MnuSub1_Click(Index As Integer)
    HandleEvent MNUSub1(Index).Tag
End Sub

Private Sub MnuSub2_Click(Index As Integer)
    HandleEvent MNUSub2(Index).Tag
End Sub

Private Sub MnuSub3_Click(Index As Integer)
    HandleEvent MNUSub3(Index).Tag
End Sub

Private Sub MnuSub4_Click(Index As Integer)
    HandleEvent MNUSub4(Index).Tag
End Sub

Private Sub MnuSub5_Click(Index As Integer)
    HandleEvent MNUSub5(Index).Tag
End Sub

Private Sub MnuSub6_Click(Index As Integer)
    HandleEvent MNUSub6(Index).Tag
End Sub

Private Sub MnuSub7_Click(Index As Integer)
    HandleEvent MNUSub7(Index).Tag
End Sub

Private Sub MnuSub8_Click(Index As Integer)
    HandleEvent MNUSub8(Index).Tag
End Sub

Private Sub MnuSub9_Click(Index As Integer)
    HandleEvent MNUSub9(Index).Tag
End Sub

Private Sub Command1_Click()
    Dim astr As String
    
astr = _
    "<clear/> " & _
"        <input name=""form""/>"

    astr = owner.DocumentManager.ExecCommand(astr)
    
    'owner.DocumentManager.Exec "INSERTjob"
    Exit Sub
    
    astr = owner.DocumentManager.xmlObjectContent(1)
    
    Dim jobhandler As New cXMLDocumentJob
    jobhandler.Title = "test"
    Set jobhandler.Manager = owner.DocumentManager
    
    Dim tmpdoc As New MSXML2.DOMDocument30
    tmpdoc.LoadXML "<job><input name=""form""/><if select=""//textbox[@validationerror!='']""><exitjob/></if><if notequal=""pinakioview"" select=""//TRN""><clear/><input name=""application""/><input name=""form""/><function name=""INSERTPINAKIOMESSAGE""/><function name=""EXECCOMMAND""/></if></job>"
    
    Dim res As Boolean
    res = jobhandler.ParseJob(tmpdoc.documentElement)
    
    astr = jobhandler.DebugMergedXml
    
    'Set jobview = xmlspy.Documents.NewFileFromText(jobHandler.DebugMergedXml, "xml")
    'jobview.SwitchViewMode 1
'"<job name=""csbtnjob"">" & _
'"    <function>" & _
'"        <xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
'"            <xsl:output method=""xml"" version=""1.0"" encoding=""UTF-8"" indent=""yes""/>" & _
'"            <xsl:template match=""/"">" & _
'"                <formupdate>" & _
'"                    <textbox name=""CSCD"">" & _
'"                        <xsl:if test=""//I_IP!='0'"">" & _
'"                            <xsl:attribute name=""TEXT""><xsl:value-of select=""//I_IP""/></xsl:attribute>" & _
'"                        </xsl:if>" & _
'"                    </textbox>" & _
'"                </formupdate>" & _
'"            </xsl:template>" & _
'"        </xsl:stylesheet>" & _
'"    </function>" & _
'"    <output name=""form""/>" & _
'"</job>" & _

End Sub


Private Sub Command2_Click()
Dim astr As String
astr = _
    "<clear/> " & _
"        <input name=""form""/>"

    astr = owner.DocumentManager.ExecCommand(astr)

End Sub

Sub ScrollForm(Direction As Byte, NewVal As Integer)
  
  Dim CTL As Control
  Static hOldVal As Integer
  Static vOldVal As Integer
  Dim hMoveDiff As Integer 'Diff in the horizontal controls movements
  Dim vMoveDiff As Integer 'Diff in the vertical controls Movements
  
  Select Case Direction
    
  Case 0 'Scroll Κάθετα
    
    If NewVal > vOldVal Then 'Scrolled From Top to Bottom
      vMoveDiff = -(NewVal - vOldVal)
    Else 'Scrolled From Bottom to Top
      vMoveDiff = (vOldVal - NewVal)
    End If
  
    For Each CTL In Me.Controls
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) And Not _
             (TypeOf CTL Is Menu) Then
        If TypeOf CTL Is Line Then
          CTL.Y1 = CTL.Y1 + vMoveDiff
          CTL.Y2 = CTL.Y2 + vMoveDiff
        Else
          CTL.Top = CTL.Top + vMoveDiff
        End If
      End If
    Next
      vOldVal = NewVal
  Case 1 'Scroll Οριζόντια
    If NewVal > hOldVal Then 'Scrolled From Left to Right
      hMoveDiff = -(NewVal - hOldVal)
    Else 'Scrolled From Right to Left
      hMoveDiff = (hOldVal - NewVal)
    End If
    For Each CTL In Me.Controls
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) And Not _
             (TypeOf CTL Is Menu) Then
        If TypeOf CTL Is Line Then
          CTL.X1 = CTL.X1 + hMoveDiff
          CTL.X2 = CTL.X2 + hMoveDiff
        Else
          CTL.Left = CTL.Left + hMoveDiff
        End If
      End If
    Next
    hOldVal = NewVal
  End Select
End Sub




Private Sub VScroll1_Change()
    ScrollForm 0, VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
  
  ScrollForm 0, VScroll1.Value

End Sub



