VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ToolBarFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Έγκριση Προϊσταμένου"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ServerSocket 
      Left            =   1920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   998
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image CaptureImage 
      Height          =   345
      Left            =   0
      Top             =   0
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Εγκρίσεις"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3555
   End
End
Attribute VB_Name = "ToolBarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ActivePassword As String, OldPassword As String, ClientName As String, ClientTerminalID As String
Public ChangePwd As Boolean
Public ClientPostDate As Date, ClientTrnNum As Integer
Public RequestCmd As Integer

Public Secret0621 As String

Public ado_DB As ADODB.Connection

Private started As Boolean, terminate As Boolean

Private Response_Code As Integer, nCmd As Long, Response
Private nNameSize As Long, szName As String * 80

Dim initHeight As Integer, initWidth As Integer
Dim GetL As Variant
Private RAuthLogin As New cRAuthLogin

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
End Function
      
Private Sub ReadImage(aComputer As String)
    Set CaptureImage.Picture = LoadPicture(AuthDir & "\" & aComputer & ".bmp")
    Height = CaptureImage.Height + 915
    Width = CaptureImage.Width + 150
    AppActivate "Έγκριση Προϊσταμένου", False
End Sub

'Private Sub BuildFormCaption(aComputer As String, aIP As String, RequestType As Integer)
'' RequestType 1: ChiefTeller, 2: Manager, 3: Chief Teller Μυστικος, 4: IRIS
'    Dim ars As ADODB.Recordset
'    Dim aUFullName As String
'    If DisableSqlServer = False Then
'        Set ars = New ADODB.Recordset
'        ars.Open "select UFullName, TermID, TrnNum, PostDate from tbl_params where Machine = '" & aComputer & "'", _
'        ado_DB, adOpenStatic + adOpenForwardOnly, adLockReadOnly
'        If ars.RecordCount > 0 Then
'            aUFullName = ars!UFullName
'            ClientPostDate = ars!PostDate
'            ClientTrnNum = ars!TrnNum
'            ClientTerminalID = ars!Termid
'        Else
'            aUFullName = ""
'            ClientPostDate = Date
'            ClientTrnNum = 0
'            ClientTerminalID = ""
'        End If
'        ars.Close: Set ars = Nothing
'    Else
'
'        Set WorkstationParams = New cWorkStationParams
'        Set WorkstationParams = WorkstationParams.InitializeRemote(aComputer)
'        If WorkstationParams Is Nothing Then GoTo HostNotFound
'        aUFullName = WorkstationParams.UFullName
'        ClientPostDate = WorkstationParams.WorkDate
'        ClientTrnNum = WorkstationParams.TrnNum
'
'    End If
'    If RequestType = 1 Then
'        RequestFrm.Prompt = "Αίτηση για κλειδί Chief Teller" & vbCrLf & _
'            "από τον χρήστη: " & ClientName & " " & aUFullName
'    ElseIf RequestType = 2 Then
'        RequestFrm.Prompt = "Αίτηση για κλειδί Manager" & vbCrLf & _
'            "από τον χρήστη: " & ClientName & " " & aUFullName
'    ElseIf RequestType = 3 Then
'        Request0621Frm.Prompt = "Αίτηση για Μυστικό Chief Teller " & vbCrLf & _
'            "από τον χρήστη: " & ClientName & " " & aUFullName
'    ElseIf RequestType = 4 Then
'        RequestFrm.Prompt = "Αίτηση για έγκριση IRIS" & vbCrLf & _
'            "από τον χρήστη: " & ClientName & " " & aUFullName
'    End If
'    If (aIP <> "") Then
'        RequestFrm.Prompt = RequestFrm.Prompt & ", IP:" & aIP
'    End If
'    Exit Sub
'HostNotFound:
'    NBG_MsgBox "Δεν εντοπίστηκαν οι πληρoφορίες για τη μηχάνη " & aComputer & "... (Α6)  " & error(), True, "ΛΑΘΟΣ"
'    Exit Sub
'End Sub

Public Function fnChkFileExist(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Open WorkDir & sFileName For Input As #1
    Close #1
    
    fnChkFileExist = True
    Exit Function
returnfalse:
    fnChkFileExist = False
End Function

Public Function fnChkFileExistAbs(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Open sFileName For Input As #1
    Close #1
    
    fnChkFileExistAbs = True
    Exit Function
returnfalse:
    fnChkFileExistAbs = False
End Function


Function Prepare_RAuth_Server() As Boolean
Dim res As Long, astr As String
    Prepare_RAuth_Server = False

'On Error GoTo InvalidUser
'    isTeller = False
'    isChiefTeller = False
'    isManager = False
'    ChkUser
'    If UseActiveDirectory Then
'        Dim adTool As New cADTool
'        adTool.Initialize
'        UserGroups.Add "test"
'        While UserGroups.Count > 0
'            UserGroups.Remove UserGroups.Count
'        Wend
'        Dim GroupName
'        For Each GroupName In adTool.UserGroups
'            If UCase(GroupName) = "TELLER" Then isTeller = True
'            If UCase(GroupName) = "CHIEF TELLER" Then isChiefTeller = True
'            If UCase(GroupName) = "MANAGER" Then isManager = True
'            If UCase(GroupName) = "IMPORT USERS" Then isImportUser = True
'            UserGroups.Add UCase(GroupName)
'        Next GroupName
'
'        UpdatexmlEnvironment "TELLER", CStr(isTeller)
'        UpdatexmlEnvironment "CHIEFTELLER", CStr(isChiefTeller)
'        UpdatexmlEnvironment "MANAGER", CStr(isManager)
'        UpdatexmlEnvironment "IMPORTUSER", CStr(isImportUser)
'
'    End If
    
    res = GetComputerName(FMachineName, MAX_COMPUTERNAME_LENGTH)
    FPDCName = GetPrimaryDCName("", "")

    
    Dim aWeblink As New cXMLWebLink
    aWeblink.VirtualDirectory = WebLink("OBJECTDISPATCHER_WEBLINK")
    Dim Method As New cXMLWebMethod
    Set Method = aWeblink.DefineDocumentMethod("DispatchObject", "http://www.nbg.gr/online/obj")

    RAuthLogin.Initialize (ReadDir + "\XmlBlocks.xml")
    RAuthLogin.WebLink = aWeblink
    RAuthLogin.WebMethod = Method
    RAuthLogin.UserName = UCase(ClearFixedString(FUserName))
    RAuthLogin.Find

    RAuthLogin.UserFullName = ClearFixedString(cFullUserName)
    RAuthLogin.ComputerName = MachineName
    RAuthLogin.BranchCode = cBRANCH
    RAuthLogin.ΒranchΙndex = cBRANCHIndex
    If (cClientIP = "") Then
        RAuthLogin.IP = ServerSocket.LocalIP
    Else
        RAuthLogin.IP = cClientIP
    End If
    
        
    
    RAuthLogin.IsChief = IIf(isChiefTeller, "1", "0")
    RAuthLogin.ΙsManager = IIf(isManager, "1", "0")

    If RAuthLogin.Password = "" And Trim(OldPassword) <> "" Then GoTo InvalidPassword
    If RAuthLogin.Password = "" Then
        GetConfirmation.Show vbModal
        If ChangePwd = False Then GoTo InvalidPasswordConfirmation
        If ChangePwd = True Then
            RAuthLogin.Password = ActivePassword

            If RAuthLogin.Insert = False Then Exit Function
            MsgBox "Ο ΚΩΔΙΚΟΣ ΣΑΣ ΚΑΤΑΧΩΡΗΘΗΚΕ!"
        End If
    Else

        If RAuthLogin.Password <> "" And Trim(OldPassword) = "" Then GoTo InvalidPassword
        If RAuthLogin.Password <> "" Then
            If (RAuthLogin.Password <> Trim(OldPassword)) Then GoTo InvalidPassword
            If Trim(OldPassword) <> Trim(ActivePassword) Then
                GetConfirmation.Show vbModal
                If ChangePwd = False Then GoTo InvalidPasswordConfirmation2
                If ChangePwd = True Then
                    RAuthLogin.Password = ActivePassword

                    If RAuthLogin.ChangePassword = False Then Exit Function
                    MsgBox "Ο ΚΩΔΙΚΟΣ ΣΑΣ ΑΛΛΑΞΕ!"
                End If
            Else
                If RAuthLogin.Update = False Then Exit Function
            End If
        End If

    End If
        
    
    Prepare_RAuth_Server = True
    Exit Function

'InvalidUser:
'    NBG_MsgBox "Λάθος στην Ταυτοποίηση του χρήστη... (B1) " & error(), True, "ΛΑΘΟΣ"
'    Exit Function
InvalidDBInit:
    NBG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Βάσης Δεδομένων... (B2) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
InvalidDBUpdate:
    NBG_MsgBox "Λάθος στην Ενημέρωση της Βάσης Δεδομένων... (B3) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
InvalidPasswordConfirmation:
    NBG_MsgBox "Λάθος Επιβεβαίωση Νέου Κωδικού... (B5) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
InvalidPassword:
    NBG_MsgBox "Λάθος Τελευταίος Κωδικός... (B6) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
InvalidPasswordConfirmation2:
    NBG_MsgBox "Λάθος Επιβεβαίωση Νέου Κωδικού... (B7) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
RAuthInfoNotFound:
    NBG_MsgBox "Δεν βρέθηκε το RauthInfo... (B8) " & error(), True, "ΛΑΘΟΣ"
    Exit Function
End Function

Private Sub Form_Load()
    Left = 0: Top = 0: initHeight = Height: initWidth = Width
    GetActivePassword.Show vbModal, Me
    If ActivePassword = "" Then End
    If Not Prepare_RAuth_Server Then GoTo InvalidRAuthServer
    
    GetL = GetKeyboardLayout(0)
    If Right(CStr(Hex(GetL)), 2) <> "08" Then ActivateKeyboardLayout 0, 0
Exit Sub
InvalidRAuthServer:
    Unload Me:  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    terminate = True
    Dim res As Long
    
    On Error GoTo ExitPoint
    
    If RAuthLogin.UserName <> "" Then RAuthLogin.Disconnect
    
    ActivateKeyboardLayout GetL, 1
ExitPoint:
 
End Sub

Private Function CreateAnswerDoc() As MSXML2.DOMDocument60

    Dim answerdoc As New MSXML2.DOMDocument60
    Dim answElm As IXMLDOMElement
    Set answerdoc.documentElement = answerdoc.createElement("RAuthAnswer")
    
    Set answElm = answerdoc.createElement("AnswerType")
    answerdoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set answElm = answerdoc.createElement("RequestID")
    answerdoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set answElm = answerdoc.createElement("UserName")
    answerdoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set CreateAnswerDoc = answerdoc

End Function
Private Sub ServerSocket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, RequestType As String, RequestMachine As String, RequestID As String, i As Integer
    Dim ippos As Integer, RequestIP As String
    Dim lR As Long
    
    On Error GoTo ExitPos
    ServerSocket.GetData strData, vbString
    On Error GoTo 0
    ServerSocket.Close
    
    Dim RequestUserName As String, RequestFullUserName As String
        
    Dim xmldoc As New MSXML2.DOMDocument60
    xmldoc.LoadXml strData
    If Not xmldoc.selectSingleNode("//RequestType") Is Nothing Then RequestType = xmldoc.selectSingleNode("//RequestType").Text
    RequestMachine = xmldoc.selectSingleNode("//MachineName").Text
    RequestID = xmldoc.selectSingleNode("//RequestID").Text
    RequestIP = xmldoc.selectSingleNode("//RequestIP").Text
    RequestUserName = xmldoc.selectSingleNode("//UserName").Text
    RequestFullUserName = xmldoc.selectSingleNode("//FullUserName").Text
    Set xmldoc = Nothing

    For i = 1 To 100
        DoEvents
    Next i
    
    Dim answerdoc As New MSXML2.DOMDocument60
    Set answerdoc = CreateAnswerDoc
    answerdoc.selectSingleNode("//AnswerType").Text = "RECEIVED"
    answerdoc.selectSingleNode("//RequestID").Text = RequestID
    answerdoc.selectSingleNode("//UserName").Text = cUserName
    
    ServerSocket.remotehost = RequestIP
    ServerSocket.RemotePort = 997
    ServerSocket.SendData answerdoc.XML
    ServerSocket.Close
    
    Set answerdoc = Nothing

On Error Resume Next
        
    Dim strprompt As String
    ReadImage RequestMachine
    
    If RequestType = "CHIEF" Then
        strprompt = "Αίτηση για κλειδί Chief Teller" & vbCrLf & _
            "από τον χρήστη: " & ClientName & " " & RequestFullUserName
    ElseIf RequestType = "MANAGER" Then
        strprompt = "Αίτηση για κλειδί Manager" & vbCrLf & _
            "από τον χρήστη: " & ClientName & " " & RequestFullUserName
    ElseIf RequestType = "IRIS" Then
        strprompt = "Αίτηση για έγκριση IRIS" & vbCrLf & _
            "από τον χρήστη: " & ClientName & " " & RequestFullUserName
    Else
        strprompt = "Αίτηση για έγκριση" & vbCrLf & _
            "από τον χρήστη: " & ClientName & " " & RequestFullUserName
    End If
    If (RequestIP <> "") Then
        strprompt = strprompt & ", IP:" & RequestIP
    End If
    RequestFrmNew.Prompt = strprompt
    RequestFrmNew.user = RequestUserName
    
    lR = SetTopMostWindow(RequestFrmNew.hwnd, True)
    keybd_event 2, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    RequestFrmNew.Show vbModal, Me
    
    Dim answerdocb As New MSXML2.DOMDocument60
    Set answerdocb = CreateAnswerDoc
    If RequestFrmNew.AcceptFlag Then
        answerdocb.selectSingleNode("//AnswerType").Text = "ACCEPTED"
    Else
        answerdocb.selectSingleNode("//AnswerType").Text = "REJECTED"
    End If
    answerdocb.selectSingleNode("//RequestID").Text = RequestID
    answerdocb.selectSingleNode("//UserName").Text = cUserName
    
    ServerSocket.remotehost = RequestIP
    ServerSocket.RemotePort = 997
    ServerSocket.SendData answerdocb.XML
    
    Set answerdocb = Nothing
    
    Set CaptureImage.Picture = Nothing
    Height = initHeight
    Width = initWidth
    
    
    ServerSocket.Bind 998
ExitPos:
End Sub

Private Sub ServerSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description
End Sub
