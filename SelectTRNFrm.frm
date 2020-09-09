VERSION 5.00
Begin VB.Form SelectTRNFrm 
   Caption         =   "Επιλογή Συναλλαγής"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox shortkey 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   1305
   End
   Begin VB.ListBox TitleList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   4320
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   150
      TabIndex        =   12
      Top             =   4590
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   150
      TabIndex        =   11
      Top             =   4230
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   150
      TabIndex        =   10
      Top             =   3870
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   150
      TabIndex        =   9
      Top             =   3510
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   150
      TabIndex        =   8
      Top             =   3150
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   150
      TabIndex        =   7
      Top             =   2790
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   2430
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   2070
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1740
      Width           =   1935
   End
   Begin VB.CommandButton MenuCommand 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton CancelBtn 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "&ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame BackFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   2400
      TabIndex        =   14
      Top             =   1830
      Width           =   915
   End
   Begin VB.Label lbShortKey 
      Caption         =   "Κωδικός Συναλλαγής"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "SelectTRNFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectedMenu As Integer
'Dim OldShortKey As String
Dim aStatus As Boolean

Private Sub SetSelectedMenu(Index As Integer)
Dim StartTop As Integer, i As Integer, k As Integer
    
    SelectedMenu = Index
    BackFrame.Top = 750: BackFrame.Left = 30
    BackFrame.width = 7815 - 60
    BackFrame.height = 6375 - 2 * 360
    TitleList.Left = 50
    TitleList.width = 7815 - 100
    StartTop = 900
    
    For i = 0 To Index
        MenuCommand(i).Left = 50
        MenuCommand(i).Top = StartTop
        MenuCommand(i).width = 7815 - 100
        StartTop = StartTop + 375
    Next i
    TitleList.Top = StartTop + 30
    StartTop = StartTop + TitleList.height
    StartTop = 6375
    For i = 9 To Index + 1 Step -1
        StartTop = StartTop - 375
        MenuCommand(i).Left = 50
        MenuCommand(i).Top = StartTop
        MenuCommand(i).width = 7815 - 100
    Next i
    If StartTop - TitleList.Top > 0 Then
        TitleList.height = StartTop - TitleList.Top - 60
    End If
    TitleList.Visible = True
    TitleList.Clear

Dim anode As Variant, bNode As Variant, astr As String
Dim HiddenFlag As Boolean
        
        Set anode = xmlNewMenu.documentElement.selectSingleNode("MenuItem[@CD='" & Trim(CStr(Index)) & "']")
        If Not (anode Is Nothing) Then
            For Each bNode In anode.childNodes
                HiddenFlag = False
                If Not (bNode.getAttributeNode("hidden") Is Nothing) Then
                    If Trim(bNode.getAttributeNode("hidden").nodeValue) = "1" Then
                        HiddenFlag = True: Exit For
                    End If
                End If
                If Not HiddenFlag Then
                    astr = bNode.getAttributeNode("id").nodeValue
                    astr = StrPad_(astr, 4, "0", "L")
                    TitleList.AddItem (astr & " - " & bNode.getAttributeNode("name").nodeValue)
                    TitleList.ItemData(TitleList.NewIndex) = CInt(astr)
                End If
            Next
        End If
        
'        Set anode = xmlMenu.documentElement.selectSingleNode("M" & Trim(CStr(Index + 1)))
'        If Not (anode Is Nothing) Then
'            For i = 1 To anode.childNodes.length - 1
'                Set bNode = anode.childNodes.item(i)
'                HiddenFlag = False
'                    If bNode.childNodes.length > 0 Then
'                        For k = 1 To bNode.childNodes.length - 1
'                            If UCase(bNode.childNodes.item(k).tagName) = "HIDDEN" Then
'                                HiddenFlag = True: Exit For
'                            End If
'                        Next k
'                    End If
'                If Not HiddenFlag Then
'                    astr = anode.childNodes.item(i).tagName
'                    astr = StrPad_(Right(astr, Len(astr) - 1), 4, "0", "L")
'                    TitleList.AddItem (astr & " - " & anode.childNodes.item(i).Text)
'                    TitleList.ItemData(TitleList.NewIndex) = CInt(astr)
'                End If
'            Next i
'        End If
        
        If SelectedMenu < 9 Then
            For i = 9 To SelectedMenu + 1 Step -1
                MenuCommand(i).TabIndex = i + 3
            Next i
        End If
        TitleList.TabIndex = SelectedMenu + 3
        For i = SelectedMenu To 0 Step -1
            MenuCommand(i).TabIndex = i + 2
        Next i
'    End If
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    shortkey.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
'        shortkey.Text = "1101"
    ElseIf KeyCode = vbKeyF4 And ((Shift And vbAltMask) = 0) Then
'        shortkey.Text = "1001"
    ElseIf KeyCode = vbKeyF5 Then
        shortkey.Text = "2100"
    ElseIf KeyCode = vbKeyF6 Then
        shortkey.Text = "2000"
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then shortkey.SetFocus
End Sub

Private Sub Form_Load()

    lbShortKey.Top = 200
    lbShortKey.Left = 50
    lbShortKey.width = 2175
    shortkey.Top = 200
    shortkey.Left = 2175 + 55
    shortkey.width = 7815 - 100 - (2175 + 55)
    
    Dim i As Integer
    i = 0
    Dim anode As IXMLDOMNode
    For Each anode In xmlNewMenu.documentElement.childNodes
        MenuCommand(i).Caption = anode.Attributes.getNamedItem("CD").Text & " - " & _
            anode.Attributes.getNamedItem("name").Text
        i = i + 1
    Next
    
    MenuCommand_Click (1)
    
    CenterFormOnScreen Me

End Sub

Private Sub MenuCommand_Click(Index As Integer)
    shortkey.Text = ""
    SetSelectedMenu (Index)
End Sub

Private Sub MenuCommand_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then _
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then shortkey.SetFocus
End Sub

Private Sub okBtn_Click()
Dim backFlag As Boolean ', backTrnCD As Integer
    
    backFlag = cEnableHiddenTransactions
'    backTrnCD = cTRNCode

    OpenTrnFrm
    
'    cTRNCode = backTrnCD
    cEnableHiddenTransactions = backFlag
    
    If aStatus Then Unload Me
End Sub

Private Sub shortkey_Change()

Dim avar As String, backTrnCD As Integer
Dim astr As String
    
    astr = shortkey.Text
    If Len(Trim(astr)) > 0 Then
        On Error GoTo clearShortkey
        
        avar = Str(CInt(astr))
        
        'If (Trim(avar) <> Trim(astr)) And ("0" & Trim(avar) <> Trim(astr)) And ("00" & Trim(avar) <> Trim(astr)) Then GoTo clearShortkey
        If Len(shortkey.Text) = 1 Then
            SetSelectedMenu (Round(Val(astr)))
        ElseIf Len(astr) = 4 And avar <> "" Then
            If avar = 610 Or avar = 611 Then
                shortkey.Text = ""
'                DoEvents: cTRNCode = CInt(astr): OldShortKey = "": shortkey.Text = ""
'                On Error GoTo 0
'                T0611Frm.Show vbModal, Me
            Else
                backTrnCD = cTRNCode
                cTRNCode = CInt(astr)
'                OldShortKey = ""
                On Error GoTo 0
                okBtn_Click
                shortkey.Text = ""
                cTRNCode = backTrnCD
            End If
        End If
        GoTo endShortKeyChk
clearShortkey:
'        shortkey.Text = OldShortKey
        shortkey.Text = ""
'        shortkey.SelStart = Len(OldShortKey)
'        shortkey.SelLength = 0
endShortKeyChk:
'        OldShortKey = shortkey.Text
End If
End Sub

Private Sub shortkey_Validate(Cancel As Boolean)
Dim aval As Double
On Error GoTo cancelupdate
    aval = Val(shortkey.Text)
    Cancel = False
    GoTo bye
cancelupdate:
    Cancel = True
bye:
    
End Sub

Private Sub TitleList_DblClick()
Dim avar As Long
    avar = TitleList.ItemData(TitleList.ListIndex)
    shortkey.Text = StrPad_(Trim(Str(avar)), 4, "0", "L")
End Sub

Private Sub TitleList_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then
    If KeyCode = vbKeyUp Then
        shortkey.SetFocus
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
        shortkey.SetFocus
        KeyCode = 0
    End If
ElseIf KeyCode = vbKeyReturn Then
Dim avar As Long
    avar = TitleList.ItemData(TitleList.ListIndex)
    shortkey.Text = StrPad_(Trim(Str(avar)), 4, "0", "L")
End If
End Sub

Private Sub OpenTrnFrm()
    cEnableHiddenTransactions = False
    On Error GoTo ExitError
StartPos:
Dim astr As String
    
    On Error GoTo 0
    Dim atrnnode As IXMLDOMElement
    Set atrnnode = TrnNodeFromTrnCode(Right("0000" & cTRNCode, 4))
    If Not (atrnnode Is Nothing) Then
        Dim HiddenFlag As Boolean
        HiddenFlag = HiddenFlagFromTrnNode(atrnnode)
'        If HiddenFlag Then
'            LogMsgbox "Λάθος Κωδικός Συναλλαγής", vbCritical, "Εφαρμογή OnLine"
'            Exit Sub
'        End If
        If HiddenFlag Then
            GoTo ExitError
        End If
        
        Dim aTRNHandler As New L2TrnHandler
        aTRNHandler.ExecuteForm Right("0000" & cTRNCode, 4)
        aTRNHandler.CleanUp
        Set aTRNHandler = Nothing
        aStatus = True
        Exit Sub
    Else
        On Error Resume Next
        Close #1
        astr = ReadDir & CStr(cTRNCode) & ".xml"
        On Error GoTo ExitError
        Open astr For Input As #1
        Close #1
        
'        On Error Resume Next
'        Dim anewTrnFrm As New TRNFrm
'        anewTrnFrm.Show vbModal, Me
'        Set anewTrnFrm = Nothing
        
        Dim aTRnFrm As New TRNFrm
        On Error Resume Next
        Load aTRnFrm
        On Error GoTo ExitError
        If aTRnFrm.CloseTransactionFlag Then
            Unload aTRnFrm
            Set aTRnFrm = Nothing
            aStatus = True
        Else
            aTRnFrm.Show vbModal, Me
            Unload aTRnFrm:
            Set aTRnFrm = Nothing
            aStatus = True
        End If
        
        If TRNQueue.count > 0 Then
            cEnableHiddenTransactions = True
            DoEvents
            cTRNCode = TRNQueue(1)
            GoTo StartPos
        End If
        Exit Sub
    End If
'ExitError:
'    aStatus = False
ExitError:
    LogMsgbox "Λάθος Κωδικός Συναλλαγής", vbCritical, "Εφαρμογή OnLine"
    aStatus = False
End Sub


