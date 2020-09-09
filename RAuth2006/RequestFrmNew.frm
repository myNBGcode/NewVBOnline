VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form RequestFrmNew 
   Caption         =   "Αίτηση Χορήγησης Κωδικού"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5565
   Begin SHDocVwCtl.WebBrowser JournalPage 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   5530
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
      Location        =   "http:///"
   End
   Begin VB.TextBox InputFld 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3900
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4020
      Width           =   1485
   End
   Begin VB.CommandButton AcceptBtn 
      Caption         =   "Αποδοχή"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4020
      Width           =   1125
   End
   Begin VB.CommandButton RejectBtn 
      Cancel          =   -1  'True
      Caption         =   "Απόρριψη"
      Height          =   495
      Left            =   1170
      TabIndex        =   1
      Top             =   4020
      Width           =   1305
   End
   Begin VB.Label PromptLabel 
      Height          =   765
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   5355
   End
   Begin VB.Label InputLbl 
      Caption         =   "Κωδικός:"
      Height          =   315
      Left            =   2820
      TabIndex        =   3
      Top             =   4020
      Width           =   1095
   End
End
Attribute VB_Name = "RequestFrmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Prompt As String
Public AcceptFlag As Boolean
Public user As String

Private Sub AcceptBtn_Click()
    If InputFld.Text <> ToolBarFrm.ActivePassword Then MsgBox "ΛΑΘΟΣ ΚΩΔΙΚΟΣ", vbOKOnly, "ΕΓΚΡΙΣΕΙΣ": Exit Sub
    AcceptFlag = True: Unload Me
End Sub

Private Sub Form_Activate()
    AcceptFlag = False
    InputFld.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        AcceptFlag = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Beep
    Beep
    
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    PromptLabel.Caption = Prompt
    AcceptFlag = False
    JournalPage.Navigate WebLink("Browser") & "SimpleBrowser.aspx?action=lastonuser&username=" & user & "&offset=1&accesstype=default&Timestamp=" & Format(Now, "yyyymmddhhmmss")

End Sub

Private Sub Form_Resize()
    If Height > 4000 And Width > 300 Then
        AcceptBtn.Top = Height - AcceptBtn.Height - 500
        RejectBtn.Top = AcceptBtn.Top
        InputFld.Top = AcceptBtn.Top
        InputLbl.Top = AcceptBtn.Top
    
        With JournalPage
            '.Width = Width - 120
            .Width = Width - 300
            .Height = AcceptBtn.Top - 50 - .Top
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    AppActivate "Εφαρμογή OnLine Συναλλαγών", False
End Sub

Private Sub RejectBtn_Click()
    AcceptFlag = False
    Unload Me
End Sub



