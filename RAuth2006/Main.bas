Attribute VB_Name = "Mainmdl"
'Attribute VB_Name = "Initial"
Option Explicit

Public FUserName As String * MAX_USERNAME_LENGTH
Public FMachineName As String * MAX_COMPUTERNAME_LENGTH
Public FPDCName As String * MAX_COMPUTERNAME_LENGTH
Public LocalFlag As Boolean

'---------------------------------------------------
Public MachineName As String
Public LogonServer As String
Public LogonDir As String
Public ReadDir As String
Public WorkDir As String
Public AuthDir As String
Public connect_status As Integer

Public gBoolStartingUp As Boolean
Public ASCII_CP_STRING As String
Public EBCDIC_CP_STRING As String

Public Strpin(15, 3) As String

Public LogonShare As String
Public cClientName As String
Public cClientIP As String

Public cBRANCH As String
Public cBRANCHIndex As String
Public cBRANCHName As String
Public cTERMINALID As String
Public cPOSTDATE As Date
Public cNextDateFlag
Public cTRNNum As Integer
Public cTRNCode As Integer
Public cPDC As String
Public cLogonServer As String
Public cUserName As String
Public cFullUserName As String
Public cHostUserName As String
Public cHostUserPassword As String
Public cJournalName As String
Public cBatchTotalsName As String
Public DisableSqlServer As Boolean
Public cUseCicsUserInfo As Boolean

Public RequestFromMachine As String
Public ChiefRequest As Boolean
Public ManagerRequest As Boolean
Public KeyAccepted As Boolean

Public WorkEnvironment As String

Public RauthParams As cRauthParams
Public WorkstationParams As cWorkStationParams

Public UseActiveDirectory As Boolean
Public RauthFile As String '= "RauthInfo.xml"

Public xmlWebLinks As New MSXML2.DOMDocument30
Public xmlstation As cXmlWorkstation
Public WebLinks As New Collection

Public xmlEnvironment As New MSXML2.DOMDocument30

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const INFINITE = -1&

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal FLAGS As Long) As Long



Private Const NORMAL_PRIORITY_CLASS = &H20&

Public Function ExecCmd(cmdline$) As Long
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret As Long

   ' Initialize the STARTUPINFO structure:
   start.cb = Len(start)

   ' Start the shelled application:
   ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


   ' Wait for the shelled application to finish:
      ret& = WaitForSingleObject(proc.hProcess, INFINITE)
      Call GetExitCodeProcess(proc.hProcess, ret&)
      Call CloseHandle(proc.hThread)
      Call CloseHandle(proc.hProcess)
      ExecCmd = ret&
End Function

'------------------------------------------------
Public Function StrPad_(PString As String, PIntLen As Integer, Optional PStrChar As Variant, _
    Optional PStrLftRgt As Variant) As String

' Η Function StrPad δέχεται ένα string πεδίο και
' επιστρέφει ένα string του μήκους που ορίζεται
' προσθέτοντας δεξιά ή αριστερά τον χαρακτήρα που
' ορίζεται όσες φορές χρειάζεται ώστε το Input πεδίο
' να γίνει στο επιθυμητό μήκος
'
' Παράμετροι :
' PString    το input String
' PIntLen    το μήκος του string που θα επιτρέψει
' PStrChar   προαιρετικά ο χαρακτήρας που θα γεμίσει
'            το υπόλοιπο μήκος default <SPACE>
' PStrLftRgt προαιρετικά αν θα προσθέσει χαρακτήρες
'            δεξιά (R) ή αριστερά (L) default δεξιά (L)
'
' π.χ.
' StrPad("12345",10)         -> "     12345"
' StrPad("12345",10, ,"R")   -> "12345     "
' StrPad("12345",10,"0")     -> "0000012345"
' StrPad("12345",10,"0","R") -> "1234500000"
' StrPad("12345",4)          -> "2345"
' StrPad("12345",4, ,"R")    -> "1234"
    
    If PIntLen <= 0 Then StrPad_ = "": Exit Function
    Dim MString As String, minti As Integer
    
    If IsMissing(PStrChar) Then PStrChar = " "
    If IsMissing(PStrLftRgt) Then PStrLftRgt = "L"
    
    For minti = 1 To PIntLen: MString = MString + PStrChar: Next

    If PStrLftRgt Like "[Ll]" Then StrPad_ = Right(MString + PString, PIntLen) _
    Else StrPad_ = Left(PString + MString, PIntLen)
End Function

Public Sub NBG_MsgBox(PStrMessage As String, _
                      Optional PBolBeep As Variant, Optional pstrTitle As Variant)
                      
' Η Procedure NBG_MsgBox πρέπει να χρησιμοποιείται
' για την εμφάνιση διαφόρων μηνυμάτων στην οθόνη
'
' Παράμετροι :
' PStrMessage το μήνυμα που θέλουμε να εμφανισθεί
' PBolBeep    προεραιτικά flag (True, False) αν
'             θέλουμε να κάνει Beep default True
' PstrTitle    ο τίτλος στο παράθυρο
' π.χ.
'
' Call NBG_MsgBox("Λανθασμένος Κωδικός !!", True,"Μήνυμα Λάθους")

  
  If IsMissing(PBolBeep) Then PBolBeep = True
  If PBolBeep Then Beep
                    
  MsgBox PStrMessage, , pstrTitle
  
  DoEvents
End Sub

Private Function ReplaceCommandFileVariables(inCmd As String) As String
Dim apos As Integer, oldLine As String
    
    oldLine = ""
    While oldLine <> inCmd
    
        oldLine = inCmd
        apos = InStr(inCmd, "%VBONLINESERVER")
        If apos > 0 Then
            inCmd = Left(inCmd, apos - 1) & Right(LogonServer, Len(LogonServer) - 2) & _
                    Right(inCmd, Len(inCmd) - apos - 14)
        End If
        apos = InStr(inCmd, "%COMPUTERNAME")
        If apos > 0 Then
            inCmd = Left(inCmd, apos - 1) & MachineName & _
                    Right(inCmd, Len(inCmd) - apos - 12)
        End If
        
    Wend
    ReplaceCommandFileVariables = inCmd

End Function

Private Function ProcessPublicCommandFile() As Boolean
Dim s As String, sHead As String, sBody As String, sCMD As String, apos As Integer
    On Error GoTo errorpos
    Open LogonDir & "PublicCMD.cfg" For Input As #1
    Do While Not EOF(1)
        Line Input #1, s
        
        apos = InStr(1, s, "=")
        If apos > 1 Then
            sHead = Trim(UCase(Left(s, apos - 1)))
            sBody = Trim(Right(s, Len(s) - apos))
            If sHead = "BATCHRUN" Then
                sCMD = ReplaceCommandFileVariables(sBody)
                ExecCmd "cmd /c" & sCMD
            ElseIf sHead = "EXERUN" Then
                sCMD = ReplaceCommandFileVariables(sBody)
                ExecCmd sCMD
            End If
        End If
   
    Loop
    Close #1
errorpos:

End Function

Public Function ChkFieldExist(ado_DB As ADODB.Connection, tablename As String, fieldname As String) As Boolean
    Dim ars As New ADODB.Recordset
    Dim astr As String
    If DisableSqlServer = False Then
     
        astr = "select count(*) as Counter " & _
                "from syscolumns join sysobjects on syscolumns.id = sysobjects.id " & _
                "Where " & _
                "sysobjects.name = '" & tablename & "' and " & _
                "syscolumns.name = '" & fieldname & "'"
    
        ars.Open astr, ado_DB, adOpenForwardOnly + adOpenStatic, adLockReadOnly
        
        If ars!Counter = 0 Then ChkFieldExist = False Else ChkFieldExist = True
        ars.Close: Set ars = Nothing
    End If
    ChkFieldExist = False
End Function

Public Sub Main()

Dim res As Integer
    
    If App.PrevInstance <> 0 Then GoTo ApplicationRunning
    
    res = GetUserName_(FUserName, MAX_USERNAME_LENGTH)
    res = GetComputerName(FMachineName, MAX_COMPUTERNAME_LENGTH)
    MachineName = ClearFixedString(FMachineName)
    cUserName = UCase(ClearFixedString(GetUserName))

    cLogonServer = Trim(Environ("LOGONSERVER"))
    cPDC = GetPrimaryDCName("", "")
    If cLogonServer = "" Then cLogonServer = cPDC
    
    LogonServer = ClearFixedString(cPDC)
    LogonShare = "VBOnline"
    cClientName = MachineName
    RauthFile = UCase(ClearFixedString(FUserName)) & ".xml"
    
    Dim commandstr As String
    commandstr = Command()
    
    Dim commandargs() As String
    
    cClientIP = ""
    
    commandargs = Split(commandstr, " ")
    If UBound(commandargs) >= 0 Then LogonServer = commandargs(0)
    If UBound(commandargs) >= 1 Then LogonShare = commandargs(1)
    If UBound(commandargs) >= 2 Then cClientName = commandargs(2)
    If UBound(commandargs) >= 3 Then cClientIP = commandargs(3)
    
    MachineName = cClientName
    
    LogonDir = LogonServer & "\" & LogonShare & "\VBLogon\"
    ReadDir = LogonServer & "\" & LogonShare & "\VBRead\"
    WorkDir = LogonServer & "\" & LogonShare & "\Network\"
    AuthDir = WorkDir
    
    On Error GoTo WebLinksError
    prepareWebLinks
    
    Call PrepareEnv
    
    cBRANCHIndex = "0"
    DisableSqlServer = True

    On Error GoTo NoCfgError
    Dim station As cWorkstationConfigurationMessage
    Set station = New cWorkstationConfigurationMessage
    Set station = station.Initialize(ReadDir + "\XmlBlocks.xml")
    station.ComputerName = UCase(MachineName)
    station.UserName = UCase(cUserName)

    Dim Method As cXMLWebMethod
    Set Method = New cXMLWebMethod
    Dim aWeblink As cXMLWebLink
    Set aWeblink = New cXMLWebLink
    aWeblink.VirtualDirectory = WebLink("OBJECTDISPATCHER_WEBLINK")
    Set Method = aWeblink.DefineDocumentMethod("DispatchObject", "http://www.nbg.gr/online/obj")
          
    Dim ares As String
    ares = Method.LoadXmlNoTrnUpdate(station.Message)
    Dim tempdoc As New MSXML2.DOMDocument60
    tempdoc.LoadXml ares
    Dim returnNode As IXMLDOMElement
    Set returnNode = GetXmlNode(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text = "0" Then Exit Sub
    
    Set xmlstation = New cXmlWorkstation
    Set xmlstation = xmlstation.Initialize(tempdoc.documentElement.XML)
    
    Set tempdoc = Nothing
    Set aWeblink = Nothing
    Set Method = Nothing
    Set station = Nothing
    
    If Not xmlstation.branch Is Nothing Then cBRANCH = xmlstation.branch.Text
    If Not xmlstation.PDC Is Nothing Then cPDC = xmlstation.PDC.Text
    If Not xmlstation.UseActiveDirectory Is Nothing Then
        If xmlstation.UseActiveDirectory.Text = "1" Then UseActiveDirectory = True
        If xmlstation.UseActiveDirectory.Text = "0" Then UseActiveDirectory = False
    End If
    If Not xmlstation.CicsUserInfo Is Nothing Then
        If xmlstation.CicsUserInfo.Text = "1" Then
            cUseCicsUserInfo = True
        Else
            cUseCicsUserInfo = False
        End If
    Else
        cUseCicsUserInfo = False
    End If
    
    
    On Error GoTo InvalidUser
    isTeller = False
    isChiefTeller = False
    isManager = False
    ChkUser
    If UseActiveDirectory Then
        Dim adTool As New cADTool
        adTool.Initialize
        UserGroups.Add "test"
        While UserGroups.Count > 0
            UserGroups.Remove UserGroups.Count
        Wend
        Dim GroupName
        For Each GroupName In adTool.UserGroups
            If UCase(GroupName) = "TELLER" Then isTeller = True
            If UCase(GroupName) = "CHIEF TELLER" Then isChiefTeller = True
            If UCase(GroupName) = "MANAGER" Then isManager = True
            If UCase(GroupName) = "IMPORT USERS" Then isImportUser = True
            UserGroups.Add UCase(GroupName)
        Next GroupName
        
        UpdatexmlEnvironment "TELLER", CStr(isTeller)
        UpdatexmlEnvironment "CHIEFTELLER", CStr(isChiefTeller)
        UpdatexmlEnvironment "MANAGER", CStr(isManager)
        UpdatexmlEnvironment "IMPORTUSER", CStr(isImportUser)
    End If
    
    If cUseCicsUserInfo Then
        isTeller = False
        isChiefTeller = False
        isManager = False
        If Not xmlstation.IsHostTeller Is Nothing Then
            If xmlstation.IsHostTeller.Text = "1" Then
                isTeller = True
            End If
        End If
        If Not xmlstation.IsHostChief Is Nothing Then
            If xmlstation.IsHostChief.Text = "1" Then
                isChiefTeller = True
            End If
        End If
        If Not xmlstation.IsHostManager Is Nothing Then
            If xmlstation.IsHostManager.Text = "1" Then
                isManager = True
            End If
        End If
    End If

    ProcessPublicCommandFile
    
    Load ToolBarFrm
    ToolBarFrm.ServerSocket.LocalPort = 998
    ToolBarFrm.ServerSocket.Bind 998

    ToolBarFrm.Show
    
Exit Sub

ApplicationRunning:
    NBG_MsgBox "Η εφαρμογή βρίσκεται ήδη σε λειτουργία....", True, "ΛΑΘΟΣ"
    Exit Sub
InvalidLogonPath:
    NBG_MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α1) " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
WebLinksError:
    NBG_MsgBox "Λάθος κατά το διάβασμα του αρχείου WebLinks " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
NoCfgError:
    NBG_MsgBox "Λάθος κατά την Εναρξη Λειτουργίας Τερματικού (Α8) " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
InvalidUser:
    NBG_MsgBox "Λάθος στην Ταυτοποίηση του χρήστη... (Α9) " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
    
End Sub

Public Function UpdatexmlEnvironment(sHead As String, sBody As String)
    
    'Dim envElm As IXMLDOMElement
    'On Error Resume Next
    'If Left(sHead, 1) <> "-" Then
        'If Not (xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))) Is Nothing) Then
        '    xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))).Text = sBody
        'Else
        '    Set envElm = xmlEnvironment.createElement(UCase(Trim(sHead)))
        '    xmlEnvironment.documentElement.appendChild envElm
        '    envElm.Text = sBody
        'End If
        
    'End If
End Function

Public Sub prepareWebLinks()
    On Error GoTo defaultWebLinks
    xmlWebLinks.Load ReadDir & "WebLink.xml"
    On Error GoTo 0
    If xmlWebLinks.Text = "" Then GoTo defaultWebLinks
    
    Dim aattr As IXMLDOMAttribute

    Set aattr = xmlWebLinks.documentElement.Attributes.getNamedItem("environment")
    If Not (aattr Is Nothing) Then
        If Trim(aattr.Text) <> "" Then
            WorkEnvironment = Left(Trim(aattr.Text), 4)
            WorkEnvironment = String(7, "0") & "." & WorkEnvironment & String(4, "0")
        End If
    End If
    
    Dim links As IXMLDOMElement
    Set links = xmlWebLinks.documentElement.selectSingleNode("//WEBLINKS/V1")
    Dim link As IXMLDOMElement
    For Each link In links.childNodes
        If link.Text <> "" Then
            WebLinks.Add link.Text, link.nodename
        End If
    Next link
    GoTo ExitPoint

defaultWebLinks:
    WebLinks.Add "", "EDUCTRADEWEBLINK"
    WebLinks.Add "http://N00000032/VirtualTradeEduc/soap", "PRODTRADEWEBLINK"
    WebLinks.Add "", "EDUCADMINWEBLINK"
    WebLinks.Add "http://N00000032/TRNStatistics/soap", "PRODADMINWEBLINK"
    WebLinks.Add "", "EDUCKPSWEBLINK"
    WebLinks.Add "http://N00000032/KPSRequest/soap", "PRODKPSWEBLINK"
    
ExitPoint:

End Sub

Public Function fnChkFileExistAbs(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Open sFileName For Input As #1
    Close #1
    
    fnChkFileExistAbs = True
    Exit Function
returnfalse:
    fnChkFileExistAbs = False
End Function

Public Function WebLink(linkName As String) As String
    
    If Left(Right(WorkEnvironment, 8), 4) = "EDUC" Then
        WebLink = WebLinks(UCase("EDUC" & linkName))
    ElseIf Left(Right(WorkEnvironment, 8), 4) = "PROD" Then
        WebLink = WebLinks(UCase("PROD" & linkName))
    Else
        WebLink = ""
        MsgBox " Δεν βρέθηκε το Virtual Directory:" & linkName
    End If
    
End Function

Public Function GetxmlEnvironment(sHead As String) As String
    
    Dim envElm As IXMLDOMElement
    On Error Resume Next
    'If Left(sHead, 1) <> "-" Then
        If Not (xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))) Is Nothing) Then
            GetxmlEnvironment = xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))).Text
        Else
            GetxmlEnvironment = ""
        End If
        
    'End If
End Function

Public Sub PrepareEnv()

    If Trim(WorkEnvironment) = "" Then
        Dim aReg As New cRegistry
        aReg.ClassKey = HKEY_LOCAL_MACHINE
        aReg.SectionKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        aReg.ValueKey = "IRIS_DB_NAME"
    
        If aReg.KeyExists Then
            WorkEnvironment = aReg.value
        Else
            WorkEnvironment = Trim(Mid(LogonServer, 3, 9)) & ".PROD" & Right("0000" & cBRANCH, 4)
        End If
        If WorkEnvironment = "" Then WorkEnvironment = Trim(Mid(Left(LogonServer & String(9, "0"), 9), 3, 9)) & ".PROD" & Right("0000" & cBRANCH, 4)
    End If

End Sub


