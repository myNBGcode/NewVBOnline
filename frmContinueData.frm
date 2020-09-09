VERSION 5.00
Begin VB.Form frmContinueData 
   ClientHeight    =   2745
   ClientLeft      =   1755
   ClientTop       =   3105
   ClientWidth     =   6675
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2745
   ScaleWidth      =   6675
   Begin VB.ListBox G0ListBox 
      Height          =   1230
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   6495
   End
   Begin VB.TextBox txtinput 
      Height          =   420
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   6495
   End
   Begin VB.CommandButton Continue_NO 
      Cancel          =   -1  'True
      Caption         =   "Ακύρωση"
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   2310
      Width           =   1695
   End
   Begin VB.CommandButton Continue_YES 
      Caption         =   "Συνέχεια"
      Default         =   -1  'True
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   2310
      Width           =   1815
   End
   Begin VB.Label FldLabel 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6495
   End
End
Attribute VB_Name = "frmContinueData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public owner As Form
Dim AMOUSEPOINTER As Integer, Selfield As Integer, EditLength As Integer, EditType As Integer, OLDTEXT As String

Public Sub Continue_NO_Click()
    ContinueCommunication = False
    
'    Call TerminateTransaction
'    Call PrintJrnl("ΟΧΙ", False)
    Unload Me
End Sub
Private Sub Continue_YES_Click()
Dim Send_status As Integer
Dim prefix As String
    
    eJournalWrite "ΝΑΙ" & txtinput(0).Text

    cb.read_again = True

    If txtinput(0).Text <> "" Then prefix = "0002" Else prefix = "9999"
    
Dim astr As String, DataLength As Integer
    
    astr = txtinput(0).Text
    If Selfield > 0 And Trim(astr) <> "" Then
        If Not ChkFldType_(astr, owner.fields(Selfield).ValidationCode) Then Exit Sub
        
        DataLength = owner.fields(Selfield).GetOutBuffLength(owner.processphase) - _
            IIf(owner.fields(Selfield).GetOutCode(owner.processphase) > 0, 2, 0)
        astr = FormatFldBeforeOut_(astr, owner.fields(Selfield).ValidationCode, owner.fields(Selfield).OutMask)
        If Len(astr) > DataLength Then
            astr = Left$(astr, DataLength)
            'Left part για τον Λογαριασμό δανείου όταν κόβεται το δευτερο CD
        Else
            If owner.fields(Selfield).EditType = etTEXT _
            Or owner.fields(Selfield).EditType = etNONE Then
                astr = StrPad_(astr, DataLength, " ", "R")
            Else
                astr = StrPad_(astr, DataLength, "0", "L")
            End If
        End If
    
        owner.fields(Selfield).Text = txtinput(0).Text
    End If
    'Mid(cb.receive_str, 3, 2) &
    cb.send_str = prefix & cHEAD & owner.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L") & astr
    'cb.send_str_length = Len(cb.send_str)
    Send_status = SEND(owner)
    
    If Send_status <> SEND_OK Then
        cb.read_again = False
        Call NBG_error("Continue_YES_Click", Send_status)
    Else
        ContinueCommunication = True
    End If
    
    Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Integer
    
    Selfield = -1
    
    For i = 1 To owner.fields.count
        If Len(owner.fields(i).Prompt) >= 3 Then
            If Mid(owner.fields(i).Prompt, 2, 2) = Mid(cb.receive_str, 3, 2) Then
                Selfield = i
                If owner.fields(Selfield).GetOutBuffLength(owner.processphase) > 0 Then Exit For Else Selfield = -1
            End If
        End If
    Next i
    
    If Selfield > 0 Then
        FldLabel.Caption = owner.fields(Selfield).Prompt
        EditLength = owner.fields(Selfield).EditLength
        EditType = owner.fields(Selfield).EditType
    Else
        EditLength = 0: EditType = 0
    End If
    OLDTEXT = ""
End Sub

Private Sub Form_Initialize()
    ShowScrollBar (200)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
            Case vbKeyEscape
                Call Continue_NO_Click
                Unload Me
                Exit Sub
    End Select
'    Call Key_Control(KeyCode)
End Sub
Private Sub Form_Load()
    'CenterFormOnScreen Me
Dim i As Integer, astr As String, alength As Integer

    G0ListBox.Clear
    For i = 1 To G0Data.count
        G0ListBox.AddItem G0Data.item(i)
    Next i

    Dim MF_Byte, MN_Bytes As String
    
  
    MF_Byte = Mid(cb.receive_str, 2, 1)
    MN_Bytes = Mid(cb.receive_str, 3, 2)
    
    alength = Len(cb.receive_str)
    
    If alength > 5 Then
        astr = Right(cb.receive_str, alength - 5)
        alength = Len(astr)
        If alength > 1 Then
            astr = Left(astr, alength - 1)
        Else
            astr = ""
        End If
    Else
        astr = ""
    End If
    G0ListBox.AddItem astr
    'FldLabel.Caption = astr
    
    Select Case MF_Byte
        Case "2"
            Select Case MN_Bytes
               Case "03"
                    Strpin(0, 0) = "03": Strpin(0, 1) = "04"
                Case "32"
                    Strpin(0, 0) = "32": Strpin(0, 1) = "01"
            End Select
    End Select
    Strpin(0, 2) = "00"
    
    CenterFormOnScreen Me
    AMOUSEPOINTER = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    G0ListBox.Clear
    For i = G0Data.count To 1 Step -1: G0Data.Remove (i): Next i
    
    Screen.MousePointer = AMOUSEPOINTER
End Sub

Private Sub txtinput_Change(Index As Integer)
Dim i As Integer
Dim astr As String
Dim aNum As Double
    If EditLength > 0 Then
        If Len(txtinput(0).Text) > EditLength Then
            txtinput(0).Text = OLDTEXT
            txtinput(0).SelStart = Len(txtinput(0).Text)
            txtinput(0).SelLength = 0
            Beep
        ElseIf Len(txtinput(0).Text) > 0 Then
            
            If EditType = etNUMBER Then
                astr = txtinput(0).Text
                On Error GoTo ErrorPos
                aNum = CDbl(astr)
                GoTo noErrorPos
ErrorPos:
                txtinput(0).Text = OLDTEXT
                txtinput(0).SelStart = Len(txtinput(0).Text)
                txtinput(0).SelLength = 0
                Beep
noErrorPos:
            Else
                astr = txtinput(0).Text
                If astr <> UCase(astr) Then
                    Dim aselpos As Integer, asellength As Integer
                    aselpos = txtinput(0).SelStart
                    asellength = txtinput(0).SelLength
                    
                    astr = UCase(astr)
                    txtinput(0).Text = astr
                    txtinput(0).SelStart = aselpos
                    txtinput(0).SelLength = asellength
                    
                End If
            End If
            
        End If
    End If
    OLDTEXT = txtinput(0).Text

End Sub

Public Sub ShowScrollBar(Optional maxlinelength As Integer)
    If IsMissing(maxlinelength) Then maxlinelength = 200
    Dim ascalemode As Integer, aSize As Integer, astr As String
    ScaleMode = 3
    astr = String(maxlinelength, "W")
    aSize = TextWidth(astr)
    SendMessage G0ListBox.hWnd, LB_SETHORIZONTALEXTENT, aSize, 0
End Sub



