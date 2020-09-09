VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form mailFrm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2445
   End
   Begin MSMAPI.MAPIMessages aMessage 
      Left            =   1080
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession aMAPISession 
      Left            =   330
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "mailFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim oSess      As Mapi.Session
    Dim oMsg       As Mapi.Message
    Dim oRecipTo   As Mapi.Recipient
    Dim oRecipCC   As Mapi.Recipient

    Set oSess = CreateObject("Mapi.Session")
    oSess.Logon "MS Exchange Settings"

    Set oMsg = oSess.Outbox.Messages.Add
    oMsg.Subject = "Test Subject from Active Messaging"
    oMsg.Text = "Test Text from Active Messaging"

    Set oRecipTo = oMsg.Recipients.Add
    oRec.Name = "biks@nbg.gr"
    oRec.Type = ActMsgTo
    oRec.Resolve

    Set oRecipCC = oMsg.Recipients.Add
    oRec.Name = "pchoun@nbg.gr"
    oRec.Type = ActMsgCC
    oRec.Resolve

    oMsg.Update
    oMsg.Send

    Set oRecipTo = Nothing
    Set oRecipCC = Nothing
    Set oMsg = Nothing
    Set oSess = Nothing

    
    
    'aMAPISession.UserName = "U34000"
    'aMAPISession.SignOn
    'aMessage.SessionID = aMAPISession.SessionID
    'aMessage.Compose
    'aMessage.RecipAddress = "biks@usa.net"
    'aMessage.MsgSubject = "Shine Transaction Error"
    'aMessage.MsgNoteText = "Journal Part"
    'aMessage.Send True
    'aMAPISession.SignOff

End Sub
