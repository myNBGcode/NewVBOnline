VERSION 5.00
Begin VB.Form MsgMemberTestFrm 
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
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "MsgMemberTestFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim aconstructor As msgmemberwsconstructor
    Set aconstructor = New msgmemberwsconstructor
    
    Dim amember As msgmember
    Set amember = aconstructor.buildmessage("http://www.nbg.gr/online/NBGApplForm", "Application")
    amember.SaveWorkDocument "c:\res1.xml"
    
    Dim astr As String
    astr = amember.XML
    Dim newdoc As New MSXML2.DOMDocument30
    newdoc.Load "c:\response_f.xml"
    amember.XML = newdoc.XML
    
    amember.SaveWorkDocument "c:\res1.xml"
    
    Dim bconstructor As msgwrapperwsconstructor
    Set bconstructor = New msgwrapperwsconstructor
    Dim awrapper As msgwrapper, bwrapper As msgwrapper
    Set awrapper = bconstructor.buildwrapper("http://www.nbg.gr/online/MassTelCat", "UpdateMonadaType")
    awrapper.workDocument.save "c:\res1.xml"
    Set bwrapper = awrapper.find("//ExistingMonadaType")
    bwrapper.value("./Monada/MonadaId") = 10
    awrapper.workDocument.save "c:\res1.xml"
    Set bwrapper = awrapper.find("//NewMonadaType")
    bwrapper.value("./Monada/MonadaId") = 20
    awrapper.workDocument.save "c:\res1.xml"
    awrapper.XML = "c:\response.xml"
    MsgBox awrapper.value("//Mother")
    'Dim bmember As msgmember
    'Set bmember = amember.find("//NewMonadaType")
    'bmember.value(".//MonadaId") = "1000"
    'amember.SaveWorkDocument "c:\res2.xml"
    
    'bmember.value("lala") = 10
    
End Sub
