VERSION 5.00
Begin VB.Form frmContinue 
   ClientHeight    =   2295
   ClientLeft      =   1680
   ClientTop       =   4425
   ClientWidth     =   7110
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
   ScaleHeight     =   2295
   ScaleWidth      =   7110
   Begin VB.CommandButton Continue_NO 
      Caption         =   "ΟΧΙ"
      Height          =   372
      Left            =   6120
      TabIndex        =   2
      Top             =   1800
      Width           =   852
   End
   Begin VB.CommandButton Continue_YES 
      Caption         =   "ΝΑΙ"
      Height          =   372
      Left            =   5160
      TabIndex        =   1
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label msgLabel 
      Caption         =   "Θέλετε να συνεχίσετε τη συναλλαγή;"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmContinue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AMOUSEPOINTER As Integer
Dim aMessage As String
Public aOwner As TRNFrm

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
On Error Resume Next
    If aOwner Is Nothing Then
        GenWorkForm.sbWriteStatusMessage sMessage
    Else
        aOwner.sbWriteStatusMessage sMessage
    End If
End Sub

Public Function fnReadStatusMessage() As String
On Error Resume Next
    If aOwner Is Nothing Then
        fnReadStatusMessage = GenWorkForm.fnReadStatusMessage
    Else
        fnReadStatusMessage = aOwner.fnReadStatusMessage
    End If
End Function

Private Sub Form_Activate()
    msgLabel.Caption = aMessage & vbCrLf & vbCrLf & "Θέλετε να συνεχίσετε τη συναλλαγή;"
    Screen.ActiveForm.sbWriteStatusMessage aMessage 'cb.receive_str
End Sub

Private Sub Form_GotFocus()
    Screen.ActiveForm.sbWriteStatusMessage aMessage
End Sub

Private Sub Form_Load()
    CenterFormOnScreen Me
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    aMessage = Screen.ActiveForm.fnReadStatusMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyEscape
                ContinueCommunication = False
                Unload Me
                Exit Sub
    End Select
'    Call Key_Control(KeyCode)
End Sub

Public Sub Continue_NO_Click()
    ContinueCommunication = False
    Unload Me
    Exit Sub
End Sub

Private Sub Continue_YES_Click()
    ContinueCommunication = True
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = AMOUSEPOINTER
End Sub
