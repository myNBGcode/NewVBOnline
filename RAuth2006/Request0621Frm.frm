VERSION 5.00
Begin VB.Form Request0621Frm 
   Caption         =   "Αίτηση Μυστικού Chief Teller"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox InputFld 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4140
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1140
      Width           =   1245
   End
   Begin VB.CommandButton AcceptBtn 
      Caption         =   "Αποδοχή"
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
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1140
      Width           =   1125
   End
   Begin VB.CommandButton RejectBtn 
      Cancel          =   -1  'True
      Caption         =   "Απόρριψη"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      TabIndex        =   2
      Top             =   1140
      Width           =   1305
   End
   Begin VB.Label PromptLabel 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   5355
   End
   Begin VB.Label InputLbl 
      Caption         =   "Μυστικός CT:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   1140
      Width           =   1575
   End
End
Attribute VB_Name = "Request0621Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Prompt As String
Public AcceptFlag As Boolean

Private Sub AcceptBtn_Click()
    If Len(InputFld.Text) <> 7 Then MsgBox "ΛΑΘΟΣ ΜΥΣΤΙΚΟΣ", vbOKOnly, "ΕΓΚΡΙΣΕΙΣ": Exit Sub
    AcceptFlag = True: ToolBarFrm.Secret0621 = InputFld.Text:  Unload Me
End Sub

Private Sub Form_Activate()
    AcceptFlag = False
End Sub

Private Sub Form_Load()
    Beep
    Beep
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    PromptLabel.Caption = Prompt
    AcceptFlag = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    AppActivate "Εφαρμογή OnLine Συναλλαγών", False
End Sub

Private Sub RejectBtn_Click()
    AcceptFlag = False: Unload Me
End Sub
