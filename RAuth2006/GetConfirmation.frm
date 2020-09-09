VERSION 5.00
Begin VB.Form GetConfirmation 
   Caption         =   "Επιβεβαίωση Νέου Κωδικού"
   ClientHeight    =   2145
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
   ScaleHeight     =   2145
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Ακύρωση"
      Height          =   435
      Left            =   2670
      TabIndex        =   2
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton ConfirmBtn 
      Caption         =   "Αποδοχή"
      Default         =   -1  'True
      Height          =   435
      Left            =   1620
      TabIndex        =   1
      Top             =   1440
      Width           =   1005
   End
   Begin VB.TextBox NewPwdConfirm 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Πληκτρείστε ξανά το νέο σας κωδικό"
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "GetConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public confirmPwd As String

Private Sub CancelBtn_Click()
   Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ConfirmBtn_Click()

   confirmPwd = NewPwdConfirm.Text
   ToolBarFrm.ChangePwd = ChkChangePwd(Trim(ToolBarFrm.ActivePassword), Trim(confirmPwd))
   Unload Me

End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2 + 1440
    
End Sub
