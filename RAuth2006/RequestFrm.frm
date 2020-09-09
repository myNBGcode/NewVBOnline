VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form RequestFrm 
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
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox JournalBox 
      Height          =   2595
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   4577
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"RequestFrm.frx":0000
   End
   Begin VB.CommandButton RejectBtn 
      Cancel          =   -1  'True
      Caption         =   "Απόρριψη"
      Height          =   495
      Left            =   1230
      TabIndex        =   1
      Top             =   3960
      Width           =   1305
   End
   Begin VB.CommandButton AcceptBtn 
      Caption         =   "Αποδοχή"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   3960
      Width           =   1125
   End
   Begin VB.TextBox InputFld 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label InputLbl 
      Caption         =   "Κωδικός:"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label PromptLabel 
      Height          =   1005
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   5355
   End
End
Attribute VB_Name = "RequestFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Prompt As String
Public AcceptFlag As Boolean

Public Function NVLString_(inValue, retValue As String) As String
    If (VarType(inValue) = vbNull) _
    Or (VarType(inValue) = vbEmpty) Then NVLString_ = retValue _
    Else NVLString_ = inValue
End Function

Private Sub AcceptBtn_Click()
    If InputFld.Text <> ToolBarFrm.ActivePassword Then MsgBox "ΛΑΘΟΣ ΚΩΔΙΚΟΣ", vbOKOnly, "ΕΓΚΡΙΣΕΙΣ": Exit Sub
    AcceptFlag = True: Unload Me
End Sub

Private Sub Form_Activate()
    InputFld.SetFocus
    AcceptFlag = False
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
    
    Dim astr As String, ars As ADODB.Recordset
    Dim aUser As String
    
    If DisableSqlServer = False Then
    
        astr = _
        "select FldTitle, DataLine, TRNCount, TRNCode, UName, ChiefUName, ManagerUName, FldNo, FinalFlag from Journal " & _
        " where postDate = '" & Format(ToolBarFrm.ClientPostDate, "yyyy/mm/dd") & "' and TerminalID = '" & ToolBarFrm.ClientTerminalID & "'" & _
        " and TrnCount >= " & CStr(ToolBarFrm.ClientTrnNum - 2) & " Order by TRNCount, SerialNo"
        
    '    MsgBox astr
        Set ars = New ADODB.Recordset
        ars.Open astr, ToolBarFrm.ado_DB, adOpenStatic, adLockReadOnly
    
    Dim aTrncount As Integer, aTrnCode As String
        aTrncount = 0: aTrnCode = ""
        If ars.RecordCount > 0 Then
            ars.MoveLast
            While Not ars.BOF
                If aTrncount <> ars!TrnCount _
                Or aTrnCode <> ars!TRNCode Then
                    If aTrncount <> 0 Then
                        aUser = Trim(NVLString_(ars!UName, ""))
                        
                        With JournalBox
                            .SelStart = 0: .SelLength = 0: .SelText = vbCrLf
                            .SelStart = 0: .SelLength = 0: .SelBold = True
                            .SelText = "Συναλλαγή: " & aTrnCode & "   A/A: " & CStr(aTrncount) & "   Χρήστης: " & aUser
                            .SelStart = 0: .SelLength = 0: .SelText = vbCrLf
                        End With
                    End If
                    
                End If
    'FldTitle, DataLine, TRNCount, TRNCode, UName, ChiefUName, ManagerUName, FinalFlag
                With JournalBox
                    If (Len(ars!dataline) > 1) Or (ars!FldNo > 0) Then
                            .SelStart = 0: .SelLength = 0: .SelBold = False
                            .SelText = vbCrLf
                            .SelStart = 0: .SelLength = 0
                            If ars!FldTitle & ars!dataline <> vbNull Then .SelText = Trim(ars!FldTitle) & Trim(ars!dataline)
                    End If
                    If ars!FinalFlag <> 0 Then
                        .SelStart = 0: .SelLength = 0: .SelBold = False
                        .SelText = "Η ΣΥΝΑΛΛΑΓΗ " & ars!TRNCode & " ΟΛΟΚΛΗΡΩΘΗΚΕ" & vbCrLf
                    End If
                End With
                
                aTrncount = ars!TrnCount
                aTrnCode = ars!TRNCode
                
                ars.MovePrevious
            Wend
            JournalBox.SelStart = 0: JournalBox.SelLength = 0: JournalBox.SelBold = True
                        
            JournalBox.SelText = "Συναλλαγή: " & _
                aTrnCode & "   A/A: " & CStr(aTrncount) & "   Χρήστης: " & aUser
                        
            JournalBox.SelStart = Len(JournalBox.Text): JournalBox.SelLength = 0
        
        End If
    Else
            
            On Error GoTo HostNotFound
            Dim path As String, filename As String
            path = WorkDir & "\" & WorkstationParams.ComputerName
            
            filename = WorkstationParams.ComputerName & "_" & Replace(WorkstationParams.WorkDate, "/", "_") & ".rtf"
            If ChkXmlFileExistRemote(filename, WorkstationParams.ComputerName) Then
                JournalBox.LoadFile (path & "\" & filename)
                JournalBox.SelStart = Len(JournalBox.Text): JournalBox.SelLength = 0
            End If
            
    End If
    Exit Sub
HostNotFound:
    NBG_MsgBox "Δεν εντοπίστηκε το αρχείο " & path & "... (Α6)  " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
End Sub

Private Sub Form_Resize()
    If Height > 4000 And Width > 300 Then
        AcceptBtn.Top = Height - AcceptBtn.Height - 500
        RejectBtn.Top = AcceptBtn.Top
        InputFld.Top = AcceptBtn.Top
        InputLbl.Top = AcceptBtn.Top
    
        With JournalBox
            .Width = Width - 120
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


