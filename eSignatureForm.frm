VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form eSignatureForm 
   Caption         =   "e-Signature"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "&Ακύρωση"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock ClientSocket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Έλεγχος ηλεκτρονικής υπογραφής..."
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "eSignatureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ReceivePending As Boolean
Public Canceled As Boolean
Public gError As Integer

Private Sub Cancel_Click()
    ReceivePending = False
    Canceled = True
    DoEvents
    Unload Me
End Sub

Private Sub ClientSocket_DataArrival(ByVal bytesTotal As Long)
    Dim AnswerStr As String
 On Error GoTo errhandler
    ClientSocket.GetData AnswerStr
    ReceivePending = False
    Unload Me
    
    If AnswerStr = "ESIGN_FAIL" Then
        gError = 1
    End If
    
 Exit Sub
errhandler:
  gError = Err.number
    If Err.number = 10054 Then
    
    Else
   
    End If
    ReceivePending = False
    Unload Me

End Sub

Public Sub SendData(Data As String)
    ClientSocket.Bind 921
    ClientSocket.SendData Data
End Sub

Private Sub Form_Initialize()
    
    If (cClientIP = "") Then
        ClientSocket.remotehost = ClientSocket.LocalIP
    Else
        ClientSocket.remotehost = cClientIP
    End If
    ClientSocket.RemotePort = 920

    gError = 0

End Sub

Private Sub Form_Load()
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoEvents
End Sub



