VERSION 5.00
Begin VB.Form IRISKeyWarning 
   Caption         =   "Ενημέρωση για Έγκριση"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton YesBtn 
      Caption         =   "Ναι"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton NoBtn 
      Cancel          =   -1  'True
      Caption         =   "Όχι"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label MessageText 
      Caption         =   "MessageText"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "IRISKeyWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AMOUSEPOINTER As Integer
Public owner As Form

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAccepted = True
        Unload Me
    ElseIf KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo 0

Dim astr As String, i As Integer
    CenterFormOnScreen Me
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    astr = "Η Συναλλαγή θα γίνει με έγκριση. Να ολοκληρωθεί;"
    MessageText.Caption = astr
    KeyAccepted = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = AMOUSEPOINTER
End Sub

Private Sub NoBtn_Click()
    Unload Me
End Sub

Private Sub YesBtn_Click()
    KeyAccepted = True
    Unload Me
End Sub

