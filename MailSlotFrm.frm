VERSION 5.00
Begin VB.Form MailSlotFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Αποστολή Μυνήματος"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelCmd 
      Cancel          =   -1  'True
      Caption         =   "Ακύρωση"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton SendCmd 
      Caption         =   "Αποστολή"
      Default         =   -1  'True
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox MsgText 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "MailSlotFrm.frx":0000
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox UsrList 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Επιλογή Παραλήπτη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Μύνημα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "MailSlotFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MachineList As Collection

Private Sub GetUsersList()

End Sub

Private Sub CancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GetUsersList
    MsgText.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    If MachineList Is Nothing Then Exit Sub
    For i = MachineList.count To 1 Step -1
        MachineList.Remove i
    Next i
    Set MachineList = Nothing
End Sub

Private Sub SendCmd_Click()
Dim astr As String, bstr As String, bstrPtr As Long, res As Long, wsize As Long, wstrPtr As Long, WSTR() As Byte
Dim mstr As String
    If MsgText.Text = "" Then Exit Sub
    
    astr = MachineList.item(UsrList.ListIndex + 1) & vbNullChar
    
    mstr = MsgText.Text & vbNullChar
    mstr = StrConv(mstr, vbUnicode)
    
    res = NetMessageBufferSend("", StrConv(astr, vbUnicode), StrConv(MachineName, vbUnicode), _
        mstr, Len(mstr) * 2)
End Sub
