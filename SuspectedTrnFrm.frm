VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SuspectedTrnFrm 
   Caption         =   "Μηνύματα"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "ok"
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
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox MessageList 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"SuspectedTrnFrm.frx":0000
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
End
Attribute VB_Name = "SuspectedTrnFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MessageDocument As MSXML2.DOMDocument
Public RequiredKey As String
Public KeyDoc As New MSXML2.DOMDocument30


Private Sub CancelBtn_Click()
    Dim resultdoc As IXMLDOMElement
        
    Set resultdoc = MessageDocument.documentElement.selectSingleNode("//SUSPECTED_TRN_RESULT")
    resultdoc.Text = "N"
    
    Unload Me
End Sub

Private Sub Form_Activate()
    'Για ύποπτες συναλλαγές
    Dim aelm As IXMLDOMElement, belm As IXMLDOMElement
    Dim aText
    For Each aelm In MessageDocument.documentElement.SelectNodes("//SKEYS1")
        For Each belm In aelm.SelectNodes("MSG_TXT")
            aText = belm.Text
            If aText <> "" Then
                MessageList.Text = MessageList.Text & aText & vbCrLf
                eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(aText)
            End If
        Next belm
        For Each belm In aelm.SelectNodes("HASH_VALUE")
            aText = belm.Text
            If aText <> "0" Then
                eJournalWriteAll ActiveL2TrnHandler.activeform, CStr("HASH_VALUE: " & aText)
            End If
        Next belm
        For Each belm In aelm.SelectNodes("EFARMOGH")
            aText = belm.Text
            If aText <> "" Then
                eJournalWriteAll ActiveL2TrnHandler.activeform, CStr("EFARMOGH: " & aText)
            End If
        Next belm
        For Each belm In aelm.SelectNodes("MSG_KWD")
            aText = belm.Text
            If aText <> "" Then
                eJournalWriteAll ActiveL2TrnHandler.activeform, CStr("MSG_KWD: " & aText)
            End If
        Next belm
        For Each belm In aelm.SelectNodes("MSG_TYPE")
            aText = belm.Text
            If aText <> "" Then
                eJournalWriteAll ActiveL2TrnHandler.activeform, CStr("MSG_TYPE: " & aText)
            End If
        Next belm
    Next aelm
End Sub


Private Sub okBtn_Click()
Dim KeyResult As String
Dim Node As IXMLDOMElement
    If RequiredKey = "M" Then
        KeyResult = L2Lib.L2ManagerKey(Nothing)
        KeyDoc.LoadXml KeyResult
        If KeyDoc.SelectNodes("//MESSAGE/ERROR").length > 0 Then
        Else
           Set Node = MessageDocument.documentElement.selectSingleNode("//NT_HEADER/CLIENT_DATA/TRAN_KEY")
           If Not (Node Is Nothing) Then Node.Text = "MANAGER"
           Set Node = MessageDocument.documentElement.selectSingleNode("//NT_HEADER/AUTHORISATION/AUTH_USER")
           If Not (Node Is Nothing) Then Node.Text = KeyDoc.selectSingleNode("//MESSAGE/MANAGER/").Text
        End If
    ElseIf RequiredKey = "C" Then
        KeyResult = L2Lib.L2ChiefKey(Nothing)
        KeyDoc.LoadXml KeyResult
        If KeyDoc.SelectNodes("//MESSAGE/ERROR").length > 0 Then
        Else
           Set Node = MessageDocument.documentElement.selectSingleNode("//NT_HEADER/CLIENT_DATA/TRAN_KEY")
           If Not (Node Is Nothing) Then Node.Text = "CHIEFTEL"
           Set Node = MessageDocument.documentElement.selectSingleNode("//NT_HEADER/AUTHORISATION/AUTH_USER")
           If Not (Node Is Nothing) Then Node.Text = KeyDoc.selectSingleNode("//MESSAGE/CHIEF/").Text
        End If
        
    End If
        
    Unload Me
End Sub
