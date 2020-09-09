VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form CurrencyMessageForm 
   Caption         =   "Μηνύματα"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton okBtn 
      Cancel          =   -1  'True
      Caption         =   "ok"
      Default         =   -1  'True
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
      Left            =   7440
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
      TextRTF         =   $"CurrencyMessageForm.frx":0000
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
Attribute VB_Name = "CurrencyMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MessageDocument As MSXML2.DOMDocument

Private Sub Form_Activate()
    Dim elm As IXMLDOMElement
    
    Dim Text
    For Each elm In MessageDocument.documentElement.SelectNodes("//STRMSG")
        Text = elm.selectSingleNode("./MSG_TEXT").Text
        If Text <> "" Then MessageList.Text = MessageList.Text & elm.selectSingleNode("./MSG_TEXT").Text & vbCrLf

         'ActiveL2TrnHandler.activeform.WriteStatusMessage CStr(Text)
         eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(Text)
    Next elm
    
End Sub
Private Sub okBtn_Click()
    Unload Me
End Sub

