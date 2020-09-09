VERSION 5.00
Begin VB.Form frmInformation 
   ClientHeight    =   6930
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9600
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   9600
   Begin VB.PictureBox SSPanel2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   6120
      Width           =   1812
      Begin VB.PictureBox TellerLogon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox ChiefTellerLogon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   600
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
      Begin VB.PictureBox ManagerLogon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pnlWorkArea 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5412
      Index           =   0
      Left            =   -120
      ScaleHeight     =   5355
      ScaleWidth      =   9930
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   9984
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1464
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1704
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   1944
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2184
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   2424
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   2664
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   2904
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   9510
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   3384
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   3624
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   14
         Left            =   120
         TabIndex        =   14
         Top             =   3864
         Width           =   9504
      End
      Begin VB.Label LBLOUTPUT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   13
         Top             =   4104
         Width           =   9504
      End
   End
   Begin VB.PictureBox pnlHeading 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9540
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox SSPanel1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9540
      TabIndex        =   4
      Top             =   6312
      Width           =   9600
      Begin VB.CommandButton cmdButton 
         Caption         =   "PageDown = епол. секида"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   7080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "PageUp = пяогц. секида"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Esc = айуяысг"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "F12 = диабибасг"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox stbStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   6120
      Width           =   7815
   End
   Begin VB.PictureBox stbStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9555
      TabIndex        =   3
      Top             =   5760
      Width           =   9615
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckKeyPressed KeyCode, cb.CurrentWorkArea
End Sub
Private Sub Form_Load()
    KeyPreview = True
    Call InitializeLogonStatus(Me)

    pnlHeading.Font.Size = 14
'biks
'    pnlHeading.Caption = "п к г я о ж о я и е с " & cb.curr_transaction & " - " & cb.Caption
 'biks
    Me.WindowState = vbMaximized
    Me.pnlWorkArea(0).Visible = True
    Me.Show
    DoEvents

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call InitializeLogonStatus(frmMenu)
End Sub


