VERSION 5.00
Begin VB.Form getkeyfrm 
   Caption         =   "Αίτηση Έγκρισης"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBtn 
      Cancel          =   -1  'True
      Caption         =   "Ακύρωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "Συνέχεια"
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
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox PassFld 
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
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox CodeFld 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Χρήστης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "getkeyfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedUser As String
Dim aLayout, bLayout As Long
Dim GetL As Variant
Dim GetLStr As String
Dim AMOUSEPOINTER As Integer


Private Sub CancelBtn_Click()
    KeyAccepted = False
    Unload Me
End Sub

Private Sub Form_Load()
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    GetL = GetKeyboardLayout(0)
    GetLStr = CStr(Hex(GetL))
    If Right(GetLStr, 2) = "08" Then aLayout = ActivateKeyboardLayout(0, 0)
    
    CodeFld.Text = SelectedUser
    KeyAccepted = False
'    : KeyAnswer = "REJECTED"
End Sub

Private Sub okBtn_Click()
    
    Dim RAuthLogin As New cRAuthLogin

    Dim aWeblink As New cXMLWebLink
    aWeblink.VirtualDirectory = TRNFrm.WebLink("OBJECTDISPATCHER_WEBLINK")
    Dim Method As New cXMLWebMethod
    Set Method = aWeblink.DefineDocumentMethod("DispatchObject", "http://www.nbg.gr/online/obj")

    RAuthLogin.Initialize (ReadDir + "\XmlBlocks.xml")
    RAuthLogin.WebLink = aWeblink
    RAuthLogin.WebMethod = Method
    RAuthLogin.UserName = CodeFld.Text
    RAuthLogin.find
    If RAuthLogin.ConnectionTimestamp > RAuthLogin.DisConnectionTimestamp And _
        ((AnyRequest And (RAuthLogin.ΙsManager = "1" Or RAuthLogin.IsChief = "1")) Or _
        (ChiefRequest And RAuthLogin.IsChief = "1") Or (ManagerRequest And RAuthLogin.ΙsManager = "1")) Then
        If (RAuthLogin.Password <> PassFld.Text) Or Trim(PassFld.Text) = "" Then
            KeyAccepted = False
            MsgBox "Δεν βρέθηκε ο χρήστης ή το password είναι λάθος...", vbOKOnly, "Λάθος"
        Else
            KeyAccepted = True
            If AnyRequest Then
                If RAuthLogin.ΙsManager = "1" Then
                   cANYKEY = "MANAGER"
                   ManagerRequest = True
                ElseIf RAuthLogin.IsChief = "1" Then
                    cANYKEY = "CHIEF"
                    ChiefRequest = True
                End If
            End If
            If ChiefRequest Then cCHIEFUserName = CodeFld.Text
            If ManagerRequest Then cMANAGERUserName = CodeFld.Text
            Unload Me
        End If
    Else
        KeyAccepted = False
        MsgBox "Δεν βρέθηκε ο χρήστης ή το password είναι λάθος...", vbOKOnly, "Λάθος"
    End If
    Set RAuthLogin = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ActivateKeyboardLayout GetL, 1
    Screen.MousePointer = AMOUSEPOINTER
End Sub

