VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form IRISRAuthGetFrm 
   Caption         =   "Αναμονή Εγκρισης"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Ακύρωση"
      Height          =   405
      Left            =   3375
      TabIndex        =   0
      Top             =   1110
      Width           =   1005
   End
   Begin MSWinsockLib.Winsock ClientSocket 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   998
   End
   Begin VB.Label PromptLabel 
      Caption         =   "Παρακαλώ περιμένετε για Έγκριση από το σταθμό: "
      Height          =   465
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   4245
   End
   Begin VB.Label StatusLabel 
      Caption         =   "Διαδικασία:"
      Height          =   465
      Left            =   135
      TabIndex        =   1
      Top             =   540
      Width           =   4245
   End
End
Attribute VB_Name = "IRISRAuthGetFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AnswerStr As String
Public requestid As String
Public requestip As String


Private Sub ClientSocket_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ContinuePos
RetryPos:
    ClientSocket.GetData AnswerStr
    GoTo SkipPos
ContinuePos:
    GoTo RetryPos
SkipPos:
    On Error GoTo 0
        
    If AnswerStr <> "" Then
        Dim AnswerType As String, AnswerUserName As String
        Dim xmldoc As New MSXML2.DOMDocument60
        xmldoc.LoadXML AnswerStr
        AnswerType = xmldoc.selectSingleNode("//AnswerType").Text
        AnswerUserName = xmldoc.selectSingleNode("//UserName").Text
        Set xmldoc = Nothing
        
        If AnswerType <> "RECEIVED" Then
            ClientSocket.Close
            If AnswerType = "ACCEPTED" Then
                KeyAccepted = True
                cIRISAuthUserName = AnswerUserName
            ElseIf AnswerType = "REJECTED" Then
                MsgBox "Η ΕΓΚΡΙΣΗ ΑΠΟΡΡΙΦΘΗΚΕ...", vbOKOnly, "On Line Εφαρμογή"
                cCHIEFUserName = ""
                cMANAGERUserName = ""
                KeyAccepted = False
            End If
            Unload Me
        End If
    End If

ExitPos:
End Sub

Private Sub ProcessRequest()
    StatusLabel.Caption = "Διαδικασία: Αποστολή Αίτησης"
    
Dim res As Integer, i As Integer, k As Integer
On Error Resume Next
    ClientSocket.Bind 997
On Error GoTo 0
    AnswerStr = ""
    Dim atime, btime
    
    While AnswerStr = ""
        atime = Time: atime = (hour(atime) * 60 + Minute(atime)) * 60 + Second(atime): btime = atime
        While Abs(btime - atime) < 2
            DoEvents
            btime = Time: btime = (hour(btime) * 60 + Minute(btime)) * 60 + Second(btime):
        Wend
        If AnswerStr = "" Then
            ClientSocket.remotehost = RequestFromMachine
            ClientSocket.RemotePort = 998

            ClientSocket.SendData Left("IRIS" & "       ", 7) & Left(MachineName & String(15, " "), 15)
        End If
        StatusLabel.Caption = "Διαδικασία: Αναμονή Απάντησης"
        atime = btime
    Wend
    
End Sub

Private Sub ProcessRequestNew()
    StatusLabel.Caption = "Διαδικασία: Αποστολή Αίτησης"
    Dim randomno As Single

Dim res As Integer, i As Integer, k As Integer
On Error Resume Next
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

    Dim atime, btime
    
    While AnswerStr = ""
        atime = Time: atime = (hour(atime) * 60 + Minute(atime)) * 60 + Second(atime): btime = atime
        While Abs(btime - atime) < 2
            DoEvents
            btime = Time: btime = (hour(btime) * 60 + Minute(btime)) * 60 + Second(btime):
        Wend
        
        If AnswerStr = "" Then
            
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
            answElm.Text = "IRIS"
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
        End If
        StatusLabel.Caption = "Διαδικασία: Αναμονή Απάντησης"
        atime = btime
    Wend
    
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

