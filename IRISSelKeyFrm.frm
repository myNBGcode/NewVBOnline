VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form IRISSelKeyFrm 
   Caption         =   "Επιλογή"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3075
   End
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin MSComctlLib.ImageList SelKeyImgLst 
      Left            =   3570
      Top             =   1560
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
            Picture         =   "IRISSelKeyFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IRISSelKeyFrm.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IRISSelKeyFrm.frx":08A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar CmdToolBar 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   3570
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   1164
      ButtonWidth     =   1773
      ButtonHeight    =   1164
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
End
Attribute VB_Name = "IRISSelKeyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AMOUSEPOINTER As Integer
Public owner As Form
Public levelAuth As String

Public Function GetSupervisor(EmpCD, SupervisorOnly As Boolean)
    Dim res, i As Integer, k As Integer, astr As String, acount As Integer, Supervisor
    GetSupervisor = False
    With GenWorkForm.AppBuffers.ByName("TR_CON_LISTA_EMPL_EMPL_TRN_I")
       .v2Value("SCROLLABLE_OCCURS") = 50
       .v2Value("COD_NRBE_EN") = "0011"
       .v2Value("ID_INTERNO_EMPL_2") = EmpCD
       .ByName("STD_TRN_I_PARM_V").Data = GenWorkForm.AppBuffers.ByName("STD_TRN_I_PARM_V").Data
       .v2Value("COD_TX") = "GCA11CON"
    End With
    res = IRISCom_(Nothing, "INEIO", "AIEDOY", GenWorkForm.AppBuffers.ByName("TR_CON_LISTA_EMPL_EMPL_TRN_I"), GenWorkForm.AppBuffers.ByName("TR_CON_LISTA_EMPL_EMPL_TRN_O"))
    If res <> 0 Then Exit Function
    If Not ChkIRISOutput_(GenWorkForm.AppBuffers.ByName("TR_CON_LISTA_EMPL_EMPL_TRN_O"), True) Then Exit Function
    With GenWorkForm.AppBuffers.ByName("TR_CON_LISTA_EMPL_EMPL_TRN_O")
        acount = .v2Value("NUMBER_OF_RECORDS")
        If acount > 0 Then
           
           For i = 1 To acount
               astr = .v2Value("COD_RL_EMPL_EMPL", i)
               If SupervisorOnly And astr <> "01" Then
               Else
                   IRISAuthList.add UCase(Trim(UCase(.v2Value("ID_INTERNO_EMPL_EP", i))))
                   IRISAuthNames.add UCase(Trim(UCase(.v2Value("NOMB_50", i))))
               End If
           Next i
        End If
    End With
    GetSupervisor = True
End Function

Public Function GetAuthList()
    If IRISAuthList.count <> 0 Then Exit Function
    If Not GenWorkForm.AppBuffers.Exists("TR_CON_LISTA_EMPL_EMPL_TRN_I") Then BuildIRISAppStruct "TR_CON_LISTA_EMPL_EMPL_TRN_I", "TR_CON_LISTA_EMPL_EMPL_TRN_I", True
    If Not GenWorkForm.AppBuffers.Exists("TR_CON_LISTA_EMPL_EMPL_TRN_O") Then BuildIRISAppStruct "TR_CON_LISTA_EMPL_EMPL_TRN_O", "TR_CON_LISTA_EMPL_EMPL_TRN_O", True
    Dim ListPos As Integer, Counter As Integer
    
    If GetSupervisor(GenWorkForm.AppBuffers.ByName("TR_APERTURA_PUESTO_TRN_I").v2Value("ID_INTERNO_EMPL_EP"), True) Then
       Counter = 0: ListPos = 1
       If IRISAuthList.count > 0 Then
          Do
             If ListPos > IRISAuthList.count Then Exit Do
             If Counter > 10 Then Exit Do
             GetSupervisor IRISAuthList.item(ListPos), False
             ListPos = ListPos + 1
             Counter = Counter + 1
          Loop
       End If

    End If
    
    
End Function
Public Function GetAuthListLevel() As Boolean
    Dim res, acount, i As Integer
    Dim astr As String
    If IRISAuthList.count <> 0 Then Exit Function
    
    'If IRISAuthList.count > 0 Then
    '   Dim k As Integer
    '   For k = IRISAuthList.count To 1 Step -1
    '        IRISAuthList.Remove k
    '   Next k
    'End If
    If Not GenWorkForm.AppBuffers.Exists("TR_EP_LISTA_EMPL_RL_EMPL_TRN_I") Then BuildIRISAppStruct "TR_EP_LISTA_EMPL_RL_EMPL_TRN_I", "TR_EP_LISTA_EMPL_RL_EMPL_TRN_I", True
    If Not GenWorkForm.AppBuffers.Exists("TR_EP_LISTA_EMPL_RL_EMPL_TRN_O") Then BuildIRISAppStruct "TR_EP_LISTA_EMPL_RL_EMPL_TRN_O", "TR_EP_LISTA_EMPL_RL_EMPL_TRN_O", True
    With GenWorkForm.AppBuffers.ByName("TR_EP_LISTA_EMPL_RL_EMPL_TRN_I")
       .v2Value("SCROLLABLE_OCCURS") = 50
       .v2Value("COD_NRBE_EN") = "0011"
       .v2Value("IND_ATRIB") = levelAuth
       .v2Value("ID_INTERNO_EMPL_2") = GenWorkForm.AppBuffers.ByName("TR_APERTURA_PUESTO_TRN_I").v2Value("ID_INTERNO_EMPL_EP")
       .ByName("STD_TRN_I_PARM_V").Data = GenWorkForm.AppBuffers.ByName("STD_TRN_I_PARM_V").Data
       .v2Value("COD_TX") = "GCA71CON"
    End With
    res = IRISCom_(Nothing, "INEF", "BV5PDV", GenWorkForm.AppBuffers.ByName("TR_EP_LISTA_EMPL_RL_EMPL_TRN_I"), GenWorkForm.AppBuffers.ByName("TR_EP_LISTA_EMPL_RL_EMPL_TRN_O"))
    If res <> 0 Then Exit Function
    If Not ChkIRISOutput_(GenWorkForm.AppBuffers.ByName("TR_EP_LISTA_EMPL_RL_EMPL_TRN_O"), True) Then Exit Function
    With GenWorkForm.AppBuffers.ByName("TR_EP_LISTA_EMPL_RL_EMPL_TRN_O")
        acount = .v2Value("NUMBER_OF_RECORDS")
        If acount > 0 Then
           For i = 1 To acount
               astr = .v2Value("COD_RL_EMPL_EMPL", i)
               IRISAuthList.add UCase(Trim(UCase(.v2Value("ID_INTERNO_EMPL_EP", i))))
               IRISAuthNames.add UCase(Trim(UCase(.v2Value("NOMB_50", i))))
           Next i
        End If
    End With
    GetAuthListLevel = True
End Function


Private Sub Form_Activate()
Dim i As Integer, x
On Error Resume Next
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

Dim ado_RAuth As ADODB.Recordset
Dim astr As String, i As Integer, Supervisor, aUser As String
    
    If Trim(levelAuth) = "" Then
        GetAuthList
    Else
        GetAuthListLevel
    End If
    CenterFormOnScreen Me
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    For i = MachineList.count To 1 Step -1
        MachineList.Remove i
    Next i
    For i = UserList.count To 1 Step -1
        UserList.Remove i
    Next i
    For i = IPList.count To 1 Step -1
        IPList.Remove i
    Next i
    
    If GenWorkForm.AppBuffers.Exists("TR_APERTURA_PUESTO_TRN_I") Then
        Dim aUserName As String
        aUserName = GenWorkForm.AppBuffers.ByName("TR_APERTURA_PUESTO_TRN_I").v2Value("ID_INTERNO_EMPL_EP")
        KeysLst.AddItem MachineName & "-" & aUserName & "-" & cFullUserName
        MachineList.add MachineName
        UserList.add aUserName
        IPList.add ""
    End If
        
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
            astr = login.ComputerName & "-" & login.UserName & "-" & login.UserFullName
            aUser = login.UserName
            For Each Supervisor In IRISAuthList
                If UCase(Supervisor) = UCase(aUser) Then
                    KeysLst.AddItem UCase(astr)
                    MachineList.add UCase(login.ComputerName)
                    UserList.add UCase(aUser)
                    IPList.add login.IP
                    Exit For
                End If
            Next Supervisor
        Next
    End If
    
    Set RAuthLogin = Nothing
    Set Rauthorities = Nothing
    
End Sub
Private Sub SelectValue()
    If KeysLst.ListIndex < 0 Then MsgBox "Επιλέξτε Χρήστη....", vbOKOnly, "Εφαρμογή OnLine": Exit Sub
    RequestFromIP = ""
    RequestFromMachine = MachineList.item(KeysLst.ListIndex + 1)
    RequestFromIP = IPList.item(KeysLst.ListIndex + 1)
    cIRISAuthUserName = UserList.item(KeysLst.ListIndex + 1)
    
    Hide
    DoEvents
    If KeysLst.ListIndex > 0 Then
        ActiveWindowToFile
        IRISRAuthGetFrm.Show vbModal, owner
    ElseIf KeysLst.ListIndex = 0 Then
        KeyAccepted = True
    End If
    Unload Me
End Sub

Private Sub SelectLocal()
    If KeysLst.ListIndex < 0 Then MsgBox "Επιλέξτε Χρήστη....", vbOKOnly, "Εφαρμογή OnLine": Exit Sub
    
    If KeysLst.ListIndex >= 1 Then
        IRISGetKeyFrm.SelectedUser = ""
        IRISGetKeyFrm.SelectedUser = UserList.item(KeysLst.ListIndex + 1)
        Unload Me
        IRISGetKeyFrm.Show vbModal, owner
    Else
        KeyAccepted = True
        cIRISAuthUserName = UserList.item(KeysLst.ListIndex + 1)
        Unload Me
    End If
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
    If MsgLst.Visible Then
        MsgLst.height = (ScaleHeight - CmdToolBar.height) \ 2
        KeysLst.Top = MsgLst.height
        KeysLst.height = ScaleHeight - CmdToolBar.height - MsgLst.height
    Else
        KeysLst.height = ScaleHeight - CmdToolBar.height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = AMOUSEPOINTER
End Sub

Private Sub KeysLst_DblClick()
    If KeysLst.ListIndex >= 0 Then
        SelectValue
    Else
        SelectLocal
    End If
End Sub



