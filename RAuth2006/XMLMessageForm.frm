VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form XMLMessageForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Μηνύματα"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox MessageList 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"XMLMessageForm.frx":0000
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
   Begin VB.CommandButton okBtn 
      Cancel          =   -1  'True
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "XMLMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MessageDocument As MSXML2.DOMDocument

Private Sub Form_Activate()
    Dim elm As IXMLDOMElement
    For Each elm In MessageDocument.documentElement.selectNodes("//ERROR|//WARNING|//MESSAGE")
        If elm.baseName = "ERROR" Then
            MessageList.Text = MessageList.Text & "ΛΑΘΟΣ:" & elm.selectSingleNode("./LINE").Text & vbCrLf
        ElseIf elm.baseName = "WARNING" Then
            MessageList.Text = MessageList.Text & "Ειδοποίηση:" & elm.selectSingleNode("./LINE").Text & vbCrLf
        ElseIf elm.baseName = "MESSAGE" And elm.selectSingleNode("./ERROR/LINE") Is Nothing Then
            MessageList.Text = MessageList.Text & "ΛΑΘΟΣ:" & elm.Text & vbCrLf
        End If
    Next elm
    
    'Για 1041newver
    Dim Text
    For Each elm In MessageDocument.documentElement.selectNodes("//STRMSG")
        Text = elm.selectSingleNode("./MSG_TEXT").Text
        If Text <> "" Then MessageList.Text = MessageList.Text & elm.selectSingleNode("./MSG_TEXT").Text & vbCrLf

         'ActiveL2TrnHandler.ActiveForm.WriteStatusMessage CStr(Text)
         'eJournalWriteAll ActiveL2TrnHandler.ActiveForm, CStr(Text)
    Next elm
    
End Sub

Private Sub okBtn_Click()
    Unload Me
End Sub
