VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SelKeyFrm 
   Caption         =   "Επιλογή"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3810
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox MsgLst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin MSComctlLib.ImageList SelKeyImgLst 
      Left            =   3690
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelKeyFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelKeyFrm.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelKeyFrm.frx":08A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar CmdToolBar 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   3180
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   1111
      ButtonWidth     =   1535
      ButtonHeight    =   1111
      Style           =   1
      ImageList       =   "SelKeyImgLst"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Επιλογή"
            Key             =   "SELECT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Τοπικά"
            Key             =   "LOCAL"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Επιστροφή"
            Key             =   "RETURN"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.ListBox KeysLst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3075
   End
   Begin VB.Label Reason 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
   End
End
Attribute VB_Name = "SelKeyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AMOUSEPOINTER As Integer
Public owner As Form
Dim ReasonText As String

Public Sub SetReasonText(Text As String)
    ReasonText = Text
End Sub

Private Sub Form_Activate()
Dim i As Integer, x
On Error Resume Next
    If G0Data.count > 0 Then MsgLst.Visible = True Else MsgLst.Visible = False
    For Each x In G0Data
        MsgLst.AddItem x
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If KeysLst.ListIndex + 1 < 1 Then Exit Sub
        SelectValue
    ElseIf KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo 0

Dim astr As String, i As Integer
    Reason.Font.Bold = True
    CenterFormOnScreen Me
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    For i = MachineList.count To 1 Step -1
        MachineList.Remove i
    Next i
    For i = UserList.count To 1 Step -1
        UserList.Remove i
    Next i
    For i = UserKeysList.count To 1 Step -1
        UserKeysList.Remove i
    Next i
    For i = IPList.count To 1 Step -1
        IPList.Remove i
    Next i
    Reason.Caption = ReasonText
    
    Dim RAuthLogin As New cRAuthLogin
    Dim Rauthorities As New Collection
    
    Dim aWeblink As New cXMLWebLink
    aWeblink.VirtualDirectory = TRNFrm.WebLink("OBJECTDISPATCHER_WEBLINK")
    Dim method As New cXMLWebMethod
    Set method = aWeblink.DefineDocumentMethod("DispatchObject", "http://www.nbg.gr/online/obj")
    
    RAuthLogin.Initialize (ReadDir + "\XmlBlocks.xml")
    RAuthLogin.WebLink = aWeblink
    RAuthLogin.WebMethod = method
    RAuthLogin.BranchCode = cBRANCH
    RAuthLogin.ΒranchΙndex = cBRANCHIndex
    RAuthLogin.ConnectionTimestamp = Date
    Set Rauthorities = RAuthLogin.FindConnectedUsers
    If Not Rauthorities Is Nothing Then
        Dim login As cRAuthLogin
        For Each login In Rauthorities
            If AnyRequest Or (ChiefRequest And login.IsChief = "1") Or (ManagerRequest And login.ΙsManager = "1") Then
                astr = login.ComputerName & "-" & login.UserName & "-" & login.UserFullName & _
                IIf(login.IsChief = "1", " (C)", " ") & IIf(login.ΙsManager = "1", " (M)", " ")
                KeysLst.AddItem astr
                MachineList.add login.ComputerName
                UserList.add login.UserName
                IPList.add login.IP
                If login.ΙsManager = "1" Then
                    UserKeysList.add "MANAGER"
                ElseIf login.IsChief = "1" Then
                    UserKeysList.add "CHIEF"
                Else
                    UserKeysList.add ""
                End If
            End If
        Next
        Set login = Nothing
    End If
    Set RAuthLogin = Nothing
    Set Rauthorities = Nothing

End Sub
Private Sub SelectValue()
    If KeysLst.ListIndex < 0 Then MsgBox "Επιλέξτε Χρήστη....", vbOKOnly, "Εφαρμογή OnLine": Exit Sub
    RequestFromMachine = MachineList.item(KeysLst.ListIndex + 1)
    RequestFromIP = ""
    
    RequestFromIP = IPList.item(KeysLst.ListIndex + 1)
    
    If AnyRequest Then
        cANYKEY = UserKeysList.item(KeysLst.ListIndex + 1)
        If cANYKEY = "CHIEF" Then ChiefRequest = True
        If cANYKEY = "MANAGER" Then ManagerRequest = True
    End If
    If ChiefRequest Then cCHIEFUserName = UserList.item(KeysLst.ListIndex + 1)
    If ManagerRequest Then cMANAGERUserName = UserList.item(KeysLst.ListIndex + 1)
    
    Hide
    DoEvents
    ActiveWindowToFile
    RAuthGetFrm.Show vbModal, owner
    Unload Me
End Sub

Private Sub SelectLocal()
    If KeysLst.ListIndex < 0 Then MsgBox "Επιλέξτε Χρήστη....", vbOKOnly, "Εφαρμογή OnLine": Exit Sub
    
    getkeyfrm.SelectedUser = ""
    If KeysLst.ListIndex >= 0 Then _
        getkeyfrm.SelectedUser = UserList.item(KeysLst.ListIndex + 1)
    Unload Me
    getkeyfrm.Show vbModal, owner
End Sub

Private Sub CmdToolbar_ButtonClick(ByVal Button As MSComctlLib.Button) 'pa
If Button.key = "SELECT" Then
    SelectValue
ElseIf Button.key = "LOCAL" And Not SecretRequest Then
    SelectLocal
ElseIf Button.key = "RETURN" Then
    Unload Me
End If
End Sub
Private Sub Form_Resize()
    DoEvents
    ScaleMode = vbTwips
    MsgLst.width = ScaleWidth
    MsgLst.Top = 0
    MsgLst.Left = 0
   
    KeysLst.width = ScaleWidth
    KeysLst.Top = 0
    KeysLst.Left = 0
    
    Reason.width = ScaleWidth
    Reason.Left = 0
    If Trim(Reason.Caption) = "" Then
        Reason.height = 0
    End If
    Reason.Top = ScaleHeight - CmdToolBar.height - Reason.height
    
    Dim fixedHeight As Single
    fixedHeight = CmdToolBar.height + Reason.height
    
    If MsgLst.Visible Then
        MsgLst.height = (ScaleHeight - fixedHeight) \ 2
        KeysLst.Top = MsgLst.height
        KeysLst.height = ScaleHeight - fixedHeight - MsgLst.height
    Else
        If ScaleHeight - fixedHeight < 0 Then
            KeysLst.height = 0
        Else
            KeysLst.height = ScaleHeight - fixedHeight
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = AMOUSEPOINTER
    ReasonText = ""
End Sub

Private Sub KeysLst_DblClick()
    SelectValue
End Sub

