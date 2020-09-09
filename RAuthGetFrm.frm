VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form RAuthGetFrm 
   Caption         =   "Αναμονή Έγκρισης"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Ακύρωση"
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      Top             =   1110
      Width           =   1005
   End
   Begin MSWinsockLib.Winsock ClientSocket 
      Left            =   100
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   998
   End
   Begin VB.Label StatusLabel 
      Caption         =   "Διαδικασία:"
      Height          =   465
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   4245
   End
   Begin VB.Label PromptLabel 
      Caption         =   "Παρακαλώ περιμένετε για Έγκριση από το σταθμό: "
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4245
   End
End
Attribute VB_Name = "RAuthGetFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AnswerStr As String
Public ReceivePending As Boolean
Public requestid As String
Public requestip As String

Private Sub ClientSocket_DataArrival(ByVal bytesTotal As Long)
    Dim idpart As String
    KeyAccepted = False
    On Error GoTo errhandler
    
    ClientSocket.GetData AnswerStr
    
    Dim AnswerType As String, AnsertRequestID As String, AnswerUserName As String
    Dim xmldoc As New MSXML2.DOMDocument60
    xmldoc.LoadXML AnswerStr
    AnswerType = xmldoc.selectSingleNode("//AnswerType").Text
    AnsertRequestID = xmldoc.selectSingleNode("//RequestID").Text
    AnswerUserName = xmldoc.selectSingleNode("//UserName").Text
    Set xmldoc = Nothing

    If AnsertRequestID <> requestid Then Exit Sub
    
    If AnswerType = "RECEIVED" Then
        ReceivePending = False
        StatusLabel.Caption = "Διαδικασία: Αναμονή Απάντησης"
    ElseIf AnswerType = "ACCEPTED" Then
        If ChiefRequest Then
            If cCHIEFUserName = "" Then
                cCHIEFUserName = AnswerUserName
            Else
                If UCase(cCHIEFUserName) <> UCase(AnswerUserName) Then
                    Exit Sub
                End If
            End If
        End If
        If ManagerRequest Then
            If ManagerRequest And cMANAGERUserName = "" Then
                cMANAGERUserName = AnswerUserName
            Else
                If UCase(cMANAGERUserName) <> UCase(AnswerUserName) Then
                    Exit Sub
                End If
            End If
        End If
        ClientSocket.Close
        KeyAccepted = True
        Unload Me
    ElseIf AnswerType = "REJECTED" Then
        LogMsgbox "Η ΕΓΚΡΙΣΗ ΑΠΟΡΡΙΦΘΗΚΕ...", vbOKOnly, "On Line Εφαρμογή"
        cCHIEFUserName = ""
        cMANAGERUserName = ""
        KeyAccepted = False
        ClientSocket.Close
        Unload Me
    End If
    
    Exit Sub
errhandler:
    If Err.number = 10054 Then
        If ReceivePending Then
            ProcessRequestNew
        End If
    Else
        LogMsgbox "ΠΡΟΒΛΗΜΑ ΣΤΟ ΣΥΣΤΗΜΑ ΕΓΚΡΙΣΕΩΝ. " & vbCrLf & "ΕΠΑΝΑΛΑΒΑΤΕ ΤΗ ΔΙΑΔΙΚΑΣΙΑ..." & vbCrLf & Err.number & " " & Err.description, vbOKOnly, "On Line Εφαρμογή"
        cCHIEFUserName = ""
        cMANAGERUserName = ""
        KeyAccepted = False
        Unload Me
    End If
End Sub

Private Sub ProcessRequest()
    StatusLabel.Caption = "Διαδικασία: Αποστολή Αίτησης"
    
Dim res As Integer, i As Integer, k As Integer
On Error Resume Next
    ReceivePending = True
    ClientSocket.Bind 997
On Error GoTo 0
    AnswerStr = ""
    If RequestFromIP <> "" Then
        ClientSocket.remotehost = RequestFromIP
    Else
        ClientSocket.remotehost = RequestFromMachine
    End If
    ClientSocket.RemotePort = 998

    Randomize
    requestid = "&RequestID=" & CStr(Rnd(100000))
    If (cClientIP = "") Then
        requestip = "&RequestIP=" & Right(String(15, " ") & ClientSocket.LocalIP, 15)
    Else
        requestip = "&RequestIP=" & Right(String(15, " ") & cClientIP, 15)
    End If
    
    If ChiefRequest And Not SecretRequest Then
        ClientSocket.SendData Left("CHIEF" & String(7, " "), 7) & Left(MachineName & String(15, " "), 15) & requestid & requestip
    ElseIf ManagerRequest Then
        ClientSocket.SendData Left("MANAGER" & String(7, " "), 7) & Left(MachineName & String(15, " "), 15) & requestid & requestip
    ElseIf ChiefRequest And SecretRequest Then
        ClientSocket.SendData Left("0621" & String(7, " "), 7) & Left(MachineName & String(15, " "), 15) & requestid & requestip
    End If
End Sub

Private Sub ProcessRequestNew()
    StatusLabel.Caption = "Διαδικασία: Αποστολή Αίτησης"
    Dim randomno As Single

On Error Resume Next
    ReceivePending = True
    ClientSocket.Bind 997
On Error GoTo 0
    AnswerStr = ""
    If RequestFromIP <> "" Then
        ClientSocket.remotehost = RequestFromIP
    Else
        ClientSocket.remotehost = RequestFromMachine
    End If
    ClientSocket.RemotePort = 998
    
    Randomize
    randomno = Rnd(100000)
    requestid = CStr(randomno)
    If (cClientIP = "") Then
        requestip = ClientSocket.LocalIP
    Else
        requestip = cClientIP
    End If
    

    Dim xmldoc As New MSXML2.DOMDocument60
    Dim answElm As IXMLDOMElement
    Set xmldoc.documentElement = xmldoc.createElement("RAuthMessage")

    Set answElm = xmldoc.createElement("RequestID")
    answElm.Text = requestid
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set answElm = xmldoc.createElement("RequestIP")
    answElm.Text = requestip
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set answElm = xmldoc.createElement("RequestType")
    If ChiefRequest Then
        answElm.Text = "CHIEF"
    ElseIf ManagerRequest Then
        answElm.Text = "MANAGER"
    End If
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing

    Set answElm = xmldoc.createElement("MachineName")
    answElm.Text = MachineName
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    Set answElm = xmldoc.createElement("UserName")
    answElm.Text = cUserName
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing

    Set answElm = xmldoc.createElement("FullUserName")
    answElm.Text = cFullUserName
    xmldoc.documentElement.appendChild answElm
    Set answElm = Nothing
    
    ClientSocket.SendData xmldoc.XML
    Set xmldoc = Nothing
    
End Sub

Private Sub CancelBtn_Click()
    AnswerStr = "REJECTED"
    DoEvents
    
    KeyAccepted = False
    cCHIEFUserName = ""
    cMANAGERUserName = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Left = (Screen.width - width) / 2
    Top = (Screen.height - height) / 2
    PromptLabel.Caption = PromptLabel.Caption & " " & RequestFromMachine
    
    ProcessRequestNew
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        AnswerStr = "REJECTED"
        DoEvents
        KeyAccepted = False
        cCHIEFUserName = ""
        cMANAGERUserName = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    AnswerStr = "REJECTED"
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoEvents
End Sub

