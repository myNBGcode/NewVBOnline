VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   2610
   ClientTop       =   1935
   ClientWidth     =   6720
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAbout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   3600
   End
   Begin VB.PictureBox Picture1 
      Height          =   1332
      Left            =   360
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   240
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "Εθνική Τράπεζα της Ελλάδος Α.Ε."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   4572
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  CenterFormOnScreen Me
  Me.Show
  DoEvents
  tmrAbout.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAbout.Enabled = False
End Sub



Private Sub tmrAbout_Timer()
  DoEvents
  gBoolStartingUp = False
  Unload frmAbout
End Sub


