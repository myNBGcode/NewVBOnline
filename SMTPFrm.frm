VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{7CE7A2CB-5104-4C11-B873-22E0D02DB883}#1.0#0"; "SMTPPanelXControl.ocx"
Begin VB.Form SMTPFrm 
   Caption         =   "Αποστολή Μυνήματος"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin SMTPPanelXControl.SMTPPanelX SMTPPanel 
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
      SMTP_Server     =   ""
      FROM_Name       =   ""
      FROM_Address    =   ""
      TO_Name         =   ""
      TO_Address      =   ""
      Subject         =   ""
      SMTP_Port       =   25
      Action          =   0
      KeepConnectionOpen=   0   'False
      WinsockStarted  =   0   'False
      TimeoutConnect  =   0
      TimeoutArp      =   0
      Alignment       =   2
      AutoSize        =   0   'False
      BevelInner      =   0
      BevelOuter      =   2
      BorderStyle     =   0
      Caption         =   ""
      Color           =   -2147483633
      Ctl3D           =   -1  'True
      UseDockManager  =   -1  'True
      DockSite        =   0   'False
      DragCursor      =   -12
      Object.DragMode        =   0
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   0   'False
      ParentColor     =   0   'False
      ParentCtl3D     =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   0   'False
      Cursor          =   0
   End
   Begin VB.CommandButton cancelCmd 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TextFld 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox TrnNumFld 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   690
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox JournalBox 
      Height          =   1575
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"SMTPFrm.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Κείμενο"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Α/Α Συναλλαγής"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label TitleLabel 
      Alignment       =   2  'Center
      Caption         =   "Αποστολή Μυνήματος στο ΚΜ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "SMTPFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    TextFld.width = width - 3 * TextFld.Left
    
    If (height - okCmd.height - 500) > 0 Then okCmd.Top = height - okCmd.height - 500
    If (height - cancelCmd.height - 500) > 0 Then cancelCmd.Top = height - cancelCmd.height - 500
    If (okCmd.Top - TextFld.Top - 100) > 0 Then TextFld.height = okCmd.Top - TextFld.Top - 100
End Sub

Private Sub SMTPPanel_OnDone()
    Unload Me
End Sub

Private Sub SMTPPanel_OnMailError(ByVal error As SMTPPanelXControl.TxSendMailError, ByVal addinfo As String)
    MsgBox addinfo, vbOKOnly, "Πρόβλημα EMail"
End Sub
