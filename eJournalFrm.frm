VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form eJournalFrm 
   Caption         =   "Form2"
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar CmdToolbar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   6570
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   582
      ButtonWidth     =   2699
      ButtonHeight    =   556
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "CommandList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Εκτύπωση"
            Key             =   "PRINT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Επιστροφή"
            Key             =   "RETURN"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Αποθήκευση"
            Key             =   "SAVE"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Αποστολή eMail"
            Key             =   "EMAIL"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList CommandList 
      Left            =   8100
      Top             =   4290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "eJournalFrm.frx":0000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "eJournalFrm.frx":0542
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "eJournalFrm.frx":0A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "eJournalFrm.frx":0FC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame SrchFrame 
      Height          =   945
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      Begin VB.TextBox SrchString 
         Height          =   315
         Left            =   5760
         TabIndex        =   5
         Top             =   510
         Width           =   2655
      End
      Begin VB.TextBox SrchTRNNo 
         Height          =   315
         Left            =   3510
         TabIndex        =   4
         Top             =   510
         Width           =   975
      End
      Begin VB.CommandButton PrintAll_Cmd 
         Appearance      =   0  'Flat
         Caption         =   "Κατάστημα"
         Height          =   765
         Left            =   10680
         Picture         =   "eJournalFrm.frx":12E0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton PrintLcl_Cmd 
         Appearance      =   0  'Flat
         Caption         =   "Τερματικό"
         Height          =   765
         Left            =   9600
         Picture         =   "eJournalFrm.frx":13E2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox SrchTerminal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3510
         TabIndex        =   2
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox SrchTRN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   510
         Width           =   1365
      End
      Begin VB.CommandButton SrchBtn 
         Appearance      =   0  'Flat
         Caption         =   "Ανάκτηση"
         Height          =   765
         Left            =   8520
         Picture         =   "eJournalFrm.frx":14E4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker SrchDate 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66846721
         CurrentDate     =   36284
      End
      Begin VB.Label Label6 
         Caption         =   "Αναζήτηση"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Σταθμός"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2700
         TabIndex        =   14
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "A/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Συναλλαγή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Ημερομηνία"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   1035
      End
   End
   Begin RichTextLib.RichTextBox vJournal 
      Height          =   3975
      Left            =   90
      TabIndex        =   0
      Top             =   1710
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7011
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"eJournalFrm.frx":1926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "eJournalFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 83 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-s
        MailSlotFrm.Show vbModal, Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim I As Integer
    I = KeyAscii
    If I = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Top = GenWorkForm.Top + 1000
    Left = GenWorkForm.Left
    width = GenWorkForm.width
    height = GenWorkForm.height - 1000
    
    Caption = "Ημερολόγιο της: " & cPOSTDATE
    
    SrchDate.value = format(cPOSTDATE, "dd/mm/yyyy")   'pa
    SrchTerminal.Text = cTERMINALID
    
'    If vJournal <> vbNull Then
'        vJournal.SetFocus
        vJournal.TextRTF = GenWorkForm.vJournal.TextRTF
        vJournal.SelStart = Len(vJournal.Text)
        vJournal.SelLength = 0
'    End If
End Sub

Private Sub Form_Resize() 'pa
    If WindowState <> vbMinimized Then
        With SrchFrame
            .Top = 0
            .Left = 0
            .width = Me.ScaleWidth
        End With
        With CmdToolbar
            .Top = Me.ScaleHeight - .height
        End With
        With vJournal
            .Left = 0
            .Top = SrchFrame.height
            .width = Me.ScaleWidth
            .height = CmdToolbar.Top - SrchFrame.height
        End With
    End If
'    With BtnFrame
'        .Top = Me.ScaleHeight - .Height
'        .Left = 0
'        .Width = Me.ScaleWidth
'    End With
'    With vJournal
'        .Left = 0
'        .Top = 0
'        .Width = Me.ScaleWidth
'        .Height = BtnFrame.Top
'    End With
End Sub

Private Sub PrintBtn_Click()
    vJournal.SelPrint (Printer.hdc)
End Sub

Private Sub PrintAll_Cmd_Click()
    On Error GoTo Error_Pos
    MousePointer = vbHourglass
    vJournal.Text = ""
    
    MousePointer = vbDefault
    vJournal.SetFocus
    vJournal.SelStart = Len(vJournal.Text)
    vJournal.SelLength = 0
Error_Pos:
End Sub

Private Sub PrintLcl_Cmd_Click()
    vJournal.SelPrint (Printer.hdc)
End Sub

Private Sub ResetBtn_Click()
    Dim aTerminal As String
    
    aTerminal = cTERMINALID
    
    vJournal.SetFocus
    vJournal.SelStart = Len(vJournal.Text)
    vJournal.SelLength = 0

End Sub

Private Sub SrchBtn_Click()
    Dim aTerminal As String
    Dim res As Boolean
    MousePointer = vbHourglass
    vJournal.Text = ""
    aTerminal = SrchTerminal.Text
    On Error GoTo Error_Pos
    Dim aValue As Integer, bvalue As Integer
    DecodeRange_ SrchTRNNo.Text, aValue, bvalue
    
    If aTerminal <> "" Then
        On Error GoTo HostNotFound
         Dim path As String, filename As String
         path = WorkDir & "\" & WorkstationParams.ComputerName
         
         filename = WorkstationParams.ComputerName & "_" & Replace(SrchDate.value, "/", "_") & ".rtf"
         If ChkXmlFileExistRemote(filename, WorkstationParams.ComputerName) Then
             vJournal.Text = ""
             vJournal.LoadFile (path & "\" & filename)
             vJournal.SelStart = Len(vJournal.Text): vJournal.SelLength = 0
         End If
    End If
    
    vJournal.SetFocus
    vJournal.SelStart = Len(vJournal.Text)
    vJournal.SelLength = 0

    MousePointer = vbDefault
Error_Pos:
    Exit Sub
HostNotFound:
    NBG_MsgBox "Δεν εντοπίστηκε το αρχείο " & path & "... (Α6)  " & error(), True, "ΛΑΘΟΣ"
    Exit Sub

End Sub
Private Sub CmdToolbar_ButtonClick(ByVal Button As MSComctlLib.Button) 'pa
If Button.Key = "PRINT" Then
    vJournal.SelPrint (Printer.hdc)
ElseIf Button.Key = "RETURN" Then
    Unload Me
ElseIf Button.Key = "SAVE" Then
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text " & _
        "Files (*.txt)|*.txt|RTF Files (*.rtf)|*.rtf"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.ShowSave
    vJournal.SaveFile CommonDialog1.filename, rtfText
ElseIf Button.Key = "EMAIL" Then
End If
End Sub

