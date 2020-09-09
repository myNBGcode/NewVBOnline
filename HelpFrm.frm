VERSION 5.00
Begin VB.Form HelpFrm 
   Caption         =   "Πληροφόρηση"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelCmd 
      Cancel          =   -1  'True
      Caption         =   "Ακύρωση"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox HelpTxt 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton OkCmd 
      Caption         =   "Επιλογή"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ListBox HelpList 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label TitlesLbl 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pa-s
Public Selections As Collection, SelectedIndex As Integer

Private Sub CancelCmd_Click()
   HelpRetValue = ""
   Unload HelpFrm
End Sub

Private Sub Form_Activate()
    HelpRetValue = ""
    If HelpList.Visible Then
        If SelectedIndex <> -1 And SelectedIndex < HelpList.ListCount - 1 Then
            HelpList.ListIndex = SelectedIndex + 1
            DoEvents
            HelpList.ListIndex = SelectedIndex
            
            
        Else
            HelpList.ListIndex = 0
        End If
        HelpList.SetFocus
        
        'HelpList.SelCount = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim aNum, retnum As Integer
   Dim found As Boolean
   Dim aMsg As String
   
   If KeyAscii = vbKeyReturn Then SelectValue
'   aNum = KeyAscii - 48
'   Form1.rsHelp.MoveFirst
'   Do Until Form1.rsHelp.EOF
'        If Form1.rsHelp("SelNo").Value = aNum Then
'            retnum = aNum
'            found = True
'        Exit Do
'        End If
'        Form1.rsHelp.MoveNext
'   Loop
'
'   If found = False Then aMsg = MsgBox("Ανύπαρκτη επιλογή", vbOKOnly)
End Sub
Private Sub SelectValue()
   Dim aIndex As Integer
   HelpRetValue = Selections.Item(HelpList.ListIndex + 1)
   Set selection = Nothing
   Unload HelpFrm
End Sub

Private Sub Form_Load()
    Form_Resize
    CenterFormOnScreen Me
    Set Selections = New Collection
End Sub

Private Sub Form_Resize()
    ScaleMode = vbPixels
        
    HelpTxt.Left = 0: HelpTxt.Width = (Width / Screen.TwipsPerPixelX) - 6
    HelpList.Left = 0: HelpList.Width = (Width / Screen.TwipsPerPixelX) - 6
    OkCmd.Width = (Width / Screen.TwipsPerPixelX - 8) / 2
    CancelCmd.Width = (Width / Screen.TwipsPerPixelX - 8) / 2
    OkCmd.Left = 0: CancelCmd.Left = OkCmd.Width
    OkCmd.Height = 25
    OkCmd.Top = (Height / Screen.TwipsPerPixelY) - OkCmd.Height - 28
    CancelCmd.Top = OkCmd.Top
    TitlesLbl.Left = 0:  TitlesLbl.Width = Width
    If HelpTxt.Text = "" Then
        HelpTxt.Visible = False
        HelpList.Top = IIf(TitlesLbl.Caption = "", 0, TitlesLbl.Height)
        HelpList.Height = OkCmd.Top - HelpList.Top
        
    Else
        HelpTxt.Visible = True
        HelpTxt.Top = 0
        If HelpList.ListCount = 0 Then
            
            HelpTxt.Height = (Height / Screen.TwipsPerPixelY) - 28
            HelpList.Visible = False
            HelpList.Height = 0
            
            OkCmd.Visible = False
            CancelCmd.Visible = False
        Else
            HelpList.Visible = True
            HelpTxt.Height = (Height / Screen.TwipsPerPixelY) / 3
            HelpList.Top = HelpTxt.Height
            HelpList.Height = OkCmd.Top - HelpTxt.Height
        End If
    End If
'    If SelectedIndex <> -1 Then HelpList.ListIndex = SelectedIndex: SelectedIndex = -1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Selections = Nothing
End Sub

Private Sub HelpList_DblClick()
   SelectValue
End Sub

Private Sub HelpList_KeyPress(KeyAscii As Integer)
Dim i As Integer, foundflag As Boolean
    foundflag = False
    For i = IIf(HelpList.ListIndex < HelpList.ListCount - 2, HelpList.ListIndex + 1, 0) To HelpList.ListCount - 1
        If Left(HelpList.List(i), 1) = UCase(Chr(KeyAscii)) Then
            HelpList.ListIndex = i: foundflag = True: Exit For
        End If
    Next i
    If Not foundflag Then
        For i = 0 To HelpList.ListCount - 1
            If Left(HelpList.List(i), 1) = UCase(Chr(KeyAscii)) Then
                HelpList.ListIndex = i: foundflag = True: Exit For
            End If
        Next i
    End If
End Sub

Private Sub okCmd_Click()
   SelectValue
End Sub
'pa-e
