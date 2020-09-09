VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DepositMessageForm 
   Caption         =   "ÃÁÌ˝Ï·Ù·"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Cancel          =   -1  'True
      Caption         =   "œ◊…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "Õ¡…"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox MessageList 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"DepositMessageForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "»›ÎÂÙÂ Ì· ÛıÌÂ˜ﬂÛÂÙÂ;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "DepositMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MessageDocument As MSXML2.DOMDocument

Private Sub CancelBtn_Click()
    Dim Text
    Text = "« ”’Õ¡ÀÀ¡√« ¡ ’—Ÿ»« ≈ ¡–œ ‘œ ◊—«”‘«..."
    MessageDocument.LoadXml "<MESSAGE><ERROR><LINE>« ”’Õ¡ÀÀ¡√« ¡ ’—Ÿ»« ≈ ¡–œ ‘œ ◊—«”‘«...</LINE></ERROR></MESSAGE>"
    eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(Text)
    Unload Me
End Sub

Private Sub Form_Activate()
   
    Dim elm As IXMLDOMElement
    Dim Text
    For Each elm In MessageDocument.documentElement.SelectNodes("//STRMSG")
        Text = elm.selectSingleNode("./MSG_TEXT").Text
        If Text <> "" Then MessageList.Text = MessageList.Text & elm.selectSingleNode("./MSG_TEXT").Text & vbCrLf
         
         ActiveL2TrnHandler.activeform.WriteStatusMessage CStr(Text)
         eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(Text)
    Next elm
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'UnloadMode = vbFormControlMenu = 0
    If (UnloadMode = vbFormControlMenu) Then CancelBtn_Click
    'UnloadMode = vbFormCode = 1 ‰ÁÎ·‰Á ÔÔı Í·Î˘ Unload Ne
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = vbKeyEscape Then
'        CancelBtn_Click
'    End If
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyEscape Then
'        CancelBtn_Click
'    End If
'End Sub
'
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
' If KeyCode = vbKeyEscape Then
'        CancelBtn_Click
'    End If
'End Sub
'
'
'
'Private Sub Form_Terminate()
'    CancelBtn_Click
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    CancelBtn_Click
'End Sub

Private Sub okBtn_Click()
 'MessageDocument.LoadXml "<MESSAGE>F12</MESSAGE>"
    MessageDocument.LoadXml ""
    Unload Me
End Sub

