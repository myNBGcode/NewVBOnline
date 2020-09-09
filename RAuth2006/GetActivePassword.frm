VERSION 5.00
Begin VB.Form GetActivePassword 
   Caption         =   "Κωδικός Εκχώρησης Άδειας"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox OldInputFld 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton CancelBtn 
      Cancel          =   -1  'True
      Caption         =   "Ακύρωση"
      Height          =   435
      Left            =   2670
      TabIndex        =   4
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton AcceptBtn 
      Caption         =   "Αποδοχή"
      Default         =   -1  'True
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   1440
      Width           =   1005
   End
   Begin VB.TextBox InputFld 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Τελευταίος Κωδικός:"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Νέος Κωδικός:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1425
   End
End
Attribute VB_Name = "GetActivePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AcceptBtn_Click()
    ToolBarFrm.OldPassword = OldInputFld.Text
    If InputFld.Text = "" Then InputFld.Text = OldInputFld.Text
    ToolBarFrm.ActivePassword = InputFld.Text
    Unload Me
End Sub

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    ToolBarFrm.ActivePassword = ""
    InputFld.SelStart = 0
    InputFld.SelLength = 4
End Sub

