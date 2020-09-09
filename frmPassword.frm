VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Chief-Teller Logon"
   ClientHeight    =   2010
   ClientLeft      =   1485
   ClientTop       =   2055
   ClientWidth     =   4200
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
   ScaleHeight     =   2010
   ScaleWidth      =   4200
   Begin VB.TextBox Password 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   18
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox UserId 
      Height          =   420
      Left            =   1440
      MaxLength       =   18
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   372
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1116
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1212
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
    If cTRNCode = 0 Then
        GenWorkForm.sbWriteStatusMessage sMessage
    Else
        TRNFrm.sbWriteStatusMessage sMessage
    End If
End Sub

Public Function fnReadStatusMessage() As String
    If cTRNCode = 0 Then
        fnReadStatusMessage = GenWorkForm.fnReadStatusMessage
    Else
        fnReadStatusMessage = TRNFrm.fnReadStatusMessage
    End If
End Function

Private Sub Form_Load()
    CenterFormOnScreen Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            If cb.curr_transaction <> "" Then
                If cb.Password Then
                    ContinueCommunication = False
                    
                    'Call TerminateTransaction
                    sbWriteStatusMessage "TERMINATING TRANSACTION. PLEASE WAIT..."
                End If
                cb.TransTerminating = True
                Unload Me
                Exit Sub
            End If
    End Select

'    Call Key_Control(KeyCode)

End Sub
Private Sub UserId_KeyPress(KeyAscii As Integer)
'    Call Text_Keypress(KeyAscii)
End Sub
Private Sub UserId_LostFocus()
    UserId.Text = UserId.Text
End Sub
Private Sub Password_KeyPress(KeyAscii As Integer)
'    Call Text_Keypress(KeyAscii)
End Sub
Private Sub Command1_Click()
Dim pUser As String, _
    pPass As String, _
    pDevice As String, _
    pRemote As String, _
    pMyCompName As String, _
    pFilePattern As String, _
    pFileName As String
Dim RetCode As Long, _
    pRetCode As Long, _
    handle As Long, _
    fSizeHigh As Long, _
    fAttrib As Long, _
    fSizeLow As Long


pRetCode = 0
pMyCompName = GKNetName(pRetCode) & Chr$(0)
pUser = UserId.Text & Chr$(0)
pPass = Password.Text & Chr$(0)
'biks
Dim ares As Boolean
ares = GKIsInGroup("\\D0YYY0\TELLER" & Chr$(0))
ares = GKIsInGroup("CHIEF TELLER" & Chr$(0))
'biks
pDevice = "X:" & Chr$(0)

Select Case Me.Caption
    Case "Teller Logon"
        pRemote = "\\" & Mid(pMyCompName, 1, Len(pMyCompName) - 1) & "\TELLER" & Chr$(0)
    Case "Chief-Teller Logon"
'biks
        pRemote = "\\" & Mid(pMyCompName, 1, Len(pMyCompName) - 1) & "\CHIEF" & Chr$(0)
 'biks
    Case "Manager Logon"
        pRemote = "\\" & Mid(pMyCompName, 1, Len(pMyCompName) - 1) & "\MANAGER" & Chr$(0)
End Select

'biks
'RetCode = GKNetUnUse(pDevice, 1)
'RetCode = GKNetUse(pDevice, pRemote, pUser, pPass)
'biks
If RetCode <> 0 Then GoTo PermissionDenied

pFilePattern = "X:\*.*" & Chr$(0)
fSizeLow = -2
pFileName = "          "
'biks
'pFileName = GKFindFirst(pFilePattern, fSizeLow, fSizeHigh, fAttrib)
'biks
If fSizeLow = -1 Then GoTo PermissionDenied

Select Case Me.Caption
    Case "Teller Logon"
        isTeller = True
        Call NBG_MsgBox("Teller Permission Granted !!", True)
    Case "Chief-Teller Logon"
        isChiefTeller = True
        Call NBG_MsgBox("Chief-Teller Permission Granted !!", True)
    Case "Manager Logon"
        isManager = True
        Call NBG_MsgBox("Manager Permission Granted !!", True)
End Select


handle = 1
'biks
'pFileName = GKFindClose(handle)
'RetCode = GKNetUnUse(pDevice, 1)
'biks

Unload Me

Exit Sub
    
PermissionDenied:
Select Case Me.Caption
    Case "Teller Logon"
        isTeller = False
        Call NBG_MsgBox("Teller Permission Denied !!", True)
    Case "Chief-Teller Logon"
        isChiefTeller = False
        Call NBG_MsgBox("Chief-Teller Permission Denied !!", True)
    Case "Manager Logon"
        isManager = False
        Call NBG_MsgBox("Manager Permission Denied !!", True)
End Select
    handle = 1
'biks
'    pFileName = GKFindClose(handle)
'    RetCode = GKNetUnUse(pDevice, 1)
'biks
End Sub

