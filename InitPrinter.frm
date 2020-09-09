VERSION 5.00
Begin VB.Form InitPrinter 
   Caption         =   "Ενεργοποίηση Εκτυπωτή"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label InitLbl 
      Caption         =   "Ο εκτυπωτής έχει ενεργοποιηθεί"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
End
Attribute VB_Name = "InitPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub TRUSTPRINTER()


Dim BufPtr() As Byte
Dim BufLen As Long, EntriesNum As Long, res As Long, i As Integer, _
    k As Integer, l As Integer, astr As String

On Error GoTo ErrorExit

res = NDdeShareEnum(vbNullString, 0, 0, 0, EntriesNum, BufLen)
If (res <> NDDE_NO_ERROR) And (res <> NDDE_BUF_TOO_SMALL) Then
    MsgBox "Πρόβλημα στην ανάκτηση του NETDDEShare (1)."
    Exit Sub
End If
ReDim BufPtr(BufLen)
res = NDdeShareEnum(vbNullString, 0, BufPtr(0), BufLen, EntriesNum, BufLen)
If (res <> NDDE_NO_ERROR) Then
    MsgBox "Πρόβλημα στην ανάκτηση του NETDDEShare (2)."
    Exit Sub
End If
l = 0
For i = 1 To EntriesNum
    astr = "": k = 1:
    While (l + k <= BufLen) And (k > 0)
        If BufPtr(l + k) <> 0 Then
            astr = astr + Chr(BufPtr(l + k))
            k = k + 1
        Else
            If UCase(Left(astr, 5)) = "NEXUS" Then
                res = NDdeSetTrustedShare(vbNullString, astr, _
                                    NDDE_TRUST_SHARE_INIT + _
                                    NDDE_TRUST_SHARE_START)
                If Not (res = NDDE_NO_ERROR) Then
                    MsgBox "Error in trusting DDE Share: " & astr
                End If
            End If
            l = l + k: k = 0
        End If
    Wend
Next i
i = 0

Exit Sub
ErrorExit:
    Exit Sub
End Sub


Private Sub Form_Load()
Dim i As Integer
Dim aMachineName As String * MAX_COMPUTERNAME_LENGTH
Dim res As Integer

    InitLbl.Top = Screen.Height / 2 - InitLbl.Height / 2
    InitLbl.Left = Screen.Width / 2 - InitLbl.Width / 2

    TRUSTPRINTER
    
    res = GetComputerName(aMachineName, MAX_COMPUTERNAME_LENGTH)
    MachineName = ClearFixedString(aMachineName)
    LogonServer = ClearFixedString(GetPrimaryDCName("", ""))
    If Environ("VBOnline_SERVER") <> "" Then LogonServer = Environ("VBOnline_SERVER")
    LogonDir = LogonServer & "\VBOnline\VBLogon\"
    If Command() <> "" Then LogonDir = LogonServer & Command()
    If Environ("VBonline_LOGONDIR") <> "" Then LogonDir = LogonServer & Environ("VBonline_LOGONDIR")
'    s = Environ("VBonline_SQLSERVER")

    DoEvents
    
    cPDC = GetPrimaryDCName("", ""): cDebug = 0: cPassbookPrinter = 0: cListToPassbook = 1
    cPrinterName = "Document"
    
    Dim fso As New FileSystemObject, ts As TextStream
    Dim s As String, apos As Integer, sHead As String, sBody As String
    On Error GoTo InvalidLogonPath
    
    Set ts = fso.OpenTextFile(LogonDir & "server.cfg", ForReading)
    Do While ts.AtEndOfStream <> True
        s = ts.ReadLine
        apos = InStr(1, s, "=")
        If apos > 1 Then
            sHead = Trim(UCase(Left(s, apos - 1)))
            sBody = Trim(Right(s, Len(s) - apos))
            If sHead = "BRANCH" Then
                cBRANCH = sBody
            ElseIf sHead = "BRANCHNAME" Then
                cBRANCHName = sBody
            ElseIf sHead = "TERMINALID" Then
                cTERMINALID = sBody
            ElseIf sHead = "LUNAME" Then
                cLUName = sBody
            ElseIf sHead = "APPLID" Then
                cApplID = sBody
            ElseIf sHead = "TIMEOUT" Then
                cTimeOut = sBody
            ElseIf sHead = "WORKDIR" Then
                WorkDir = LogonServer & sBody
            ElseIf sHead = "READDIR" Then
'                ReadDir = LogonServer & sBody
                ReadDir = sBody
            ElseIf sHead = "AUTHDIR" Then
                AuthDir = LogonServer & sBody
            ElseIf sHead = "PRINTERNAME" Then
                cPrinterName = sBody
            ElseIf sHead = "PASSBOOKPRINTER" Then
                cPassbookPrinter = sBody
            ElseIf sHead = "LISTTOPASSBOOK" Then
                cListToPassbook = sBody
            ElseIf sHead = "PDC" Then
                cPDC = sBody
            ElseIf sHead = "DEBUG" Then
                cDebug = CInt(Trim(sBody))
            ElseIf sHead = "SMTPSERVER" Then
                cSMTPServer = Trim(sBody)
            ElseIf sHead = "SMTPPORT" Then
                cSMTPPort = CInt(Trim(sBody))
            ElseIf sHead = "HELPDESKADDRESS" Then
                cHELPDESKADDRESS = Trim(sBody)
            ElseIf sHead = "HELPDESKSUBJECT" Then
                cHELPDESKSUBJECT = Trim(sBody)
            End If
        End If
    Loop
    ts.Close: Set ts = Nothing
    Set ts = fso.OpenTextFile(LogonDir & MachineName & ".cfg", ForReading)
    Do While ts.AtEndOfStream <> True
        s = ts.ReadLine
        apos = InStr(1, s, "=")
        If apos > 1 Then
            sHead = Trim(UCase(Left(s, apos - 1)))
            sBody = Trim(Right(s, Len(s) - apos))
            If sHead = "BRANCH" Then
                cBRANCH = sBody
            ElseIf sHead = "BRANCHNAME" Then
                cBRANCHName = sBody
            ElseIf sHead = "TERMINALID" Then
                cTERMINALID = sBody
            ElseIf sHead = "LUNAME" Then
                cLUName = sBody
            ElseIf sHead = "APPLID" Then
                cApplID = sBody
            ElseIf sHead = "TIMEOUT" Then
                cTimeOut = sBody
            ElseIf sHead = "WORKDIR" Then
                WorkDir = LogonServer & sBody
            ElseIf sHead = "READDIR" Then
'                ReadDir = LogonServer & sBody
                ReadDir = sBody
            ElseIf sHead = "AUTHDIR" Then
                AuthDir = LogonServer & sBody
            ElseIf sHead = "PRINTERNAME" Then
                cPrinterName = sBody
            ElseIf sHead = "PASSBOOKPRINTER" Then
                cPassbookPrinter = sBody
            ElseIf sHead = "LISTTOPASSBOOK" Then
                cListToPassbook = sBody
            ElseIf sHead = "PDC" Then
                cPDC = sBody
            ElseIf sHead = "DEBUG" Then
                cDebug = CInt(Trim(sBody))
            ElseIf sHead = "SMTPSERVER" Then
                cSMTPServer = Trim(sBody)
            ElseIf sHead = "SMTPPORT" Then
                cSMTPPort = CInt(Trim(sBody))
            ElseIf sHead = "HELPDESKADDRESS" Then
                cHELPDESKADDRESS = Trim(sBody)
            ElseIf sHead = "HELPDESKSUBJECT" Then
                cHELPDESKSUBJECT = Trim(sBody)
            End If
        End If
    Loop
    ts.Close: Set ts = Nothing
    
    On Error GoTo InvalidChkUser:
    ChkUser
    If UCase(cUserName) <> UCase("Printer") Then
        MsgBox "ΛΑΘΟΣ ΟΝΟΜΑ ΧΡΗΣΤΗ.": Unload Me
    Else
        EventLogWrite = False
        SendJournalWrite = False
        ReceiveJournalWrite = False
        SRJournal = False
        
        Dim aflag As Boolean, aresult As Long, trycount As Integer
        aflag = (cPassbookPrinter = 1 Or cPassbookPrinter = 2)
        On Error GoTo InvalidPrinterInit:
        If aflag Then
            Dim Status As Long
            Set docPrinter = New NXPrinter
step1:
            Status = docPrinter.WStart
            If Status <> 0 And Status <> -1 And i > 10 Then
                aresult = MsgBox("Πρόβλημα στην έναρξη του εκτυπωτή A:" & CStr(Status) & ". Να διακοπεί η διαδικασία;", vbYesNo, "Πρόβλημα Εκτύπωσης")
                If aresult = vbYes Then GoTo InvalidPrinterInit Else i = 0
            End If
            If Status <> 0 And Status <> -1 Then
                GoTo step1
            End If
            trycount = 0
step2:
            trycount = trycount + 1
            If cPassbookPrinter = 1 Then
                Status = docPrinter.WOpen(cPrinterName, ReadDir & "PR2Lib.vmd", "TestForm", 10000)
            ElseIf cPassbookPrinter = 2 Then
                Status = docPrinter.WOpen(cPrinterName, ReadDir & "HP4905Lib.vmd", "TestForm", 10000)
            End If
            
            If Status <> 0 Then
                aresult = MsgBox("Πρόβλημα στο άνοιγμα του εκτυπωτή B:" & CStr(Status) & ". Να διακοπεί η διαδικασία;", vbYesNo, "Πρόβλημα Εκτύπωσης")
                If aresult = vbYes Then GoTo InvalidPrinterInit Else i = 0
            End If
            If Status <> 0 Then GoTo step2
        End If
    '    Status = docPrinter.WClose
step3:
    End If
Exit Sub

InvalidLogonPath:
    MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α1) " & Error(), True, "ΛΑΘΟΣ"
    Unload Me
InvalidChkUser:
    MsgBox "Λάθος στην Ταυτοποίηση Χρήστη... (Α6) " & Error(), True, "ΛΑΘΟΣ"
    Unload Me
InvalidPrinterInit:
    MsgBox "Λάθος στην Έναρξη Λειτουργίας Εκτυπωτή... (Α7) " & Error(), True, "ΛΑΘΟΣ"
    Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 And Shift = vbAltMask Then   ' Display key combinations.
      Unload Me
   End If

End Sub

