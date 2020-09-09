VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOverwrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Data As String

Public Property Get Header() As String
    Header = Left(Data, 39)
End Property

Public Property Get content() As String
    content = Trim(Right(Data, Len(Data) - 39))
    Dim apos As Integer
    apos = InStr(1, content, "`")
    If apos <> 0 Then
        content = Trim(Left(content, apos - 1))
    End If
End Property

Public Property Get IsInfo() As Boolean
    Dim opart As String
    opart = Right(Header, 11)
    IsInfo = Left(opart, 2) = "II"
End Property

Public Property Get IsKey() As Boolean
    Dim opart As String
    opart = Right(Header, 11)
    IsKey = Left(opart, 2) = "KK"
End Property

Public Property Get RequiredKey() As String
    If IsKey Then
        RequiredKey = Mid(Header, 33, 1)
    End If
End Property

Public Property Get UpdatedHeader() As String
    UpdatedHeader = Mid(Header, 29, 11)
    If IsKey Then UpdatedHeader = Left(UpdatedHeader, 8) & Mid(UpdatedHeader, 5, 1) & Right(UpdatedHeader, 2)
End Property
Public Property Get TimeStamp() As String
    TimeStamp = Left(Header, 28)
End Property
Sub ShowMessage(Message As String)
    Dim Control
    For Each Control In Screen.activeform.Controls
        If TypeOf Control Is StatusBar Then
            Control.Panels(1).Text = Message
            Exit For
        End If
    Next Control
End Sub

Function GetKey() As Boolean
    KeyAccepted = False
    If RequiredKey = "C" Then
        ChiefRequest = True:
        ManagerRequest = False:
        If Not isChiefTeller Then
            Set SelKeyFrm.owner = Screen.activeform
            SelKeyFrm.Show vbModal, Screen.activeform
        Else
            Load KeyWarning:
            Set KeyWarning.owner = Screen.activeform
            KeyWarning.Show vbModal, Screen.activeform
        End If
    ElseIf RequiredKey = "M" Then
        ChiefRequest = False:
        ManagerRequest = True:
        If Not isManager Then
            Set SelKeyFrm.owner = Screen.activeform
            SelKeyFrm.Show vbModal, Screen.activeform
        Else
            Load KeyWarning:
            Set KeyWarning.owner = Screen.activeform
            KeyWarning.Show vbModal, Screen.activeform
        End If
    End If
    GetKey = KeyAccepted
End Function

Public Function HandleMessage() As VbMsgBoxResult
    If IsInfo Then
        eJournalWriteAll Nothing, content
        ShowMessage content
        
        HandleMessage = MsgBox(content & vbCrLf & vbCrLf & "»›ÎÂÙÂ Ì· ÛıÌÂ˜ﬂÛÂÙÂ ÙÁ ÛıÌ·ÎÎ·„ﬁ;", vbOKCancel, "≈ÌÁÏ›Ò˘ÛÁ")
        If HandleMessage = vbOK Then
            eJournalWriteAll Nothing, "≈–…¬≈¬¡…Ÿ”« √…¡ ”’Õ≈◊≈…¡ ‘«” ”’Õ¡ÀÀ¡√«”"
        Else
            eJournalWriteAll Nothing, "ƒ…¡ œ–« ‘«” ”’Õ¡ÀÀ¡√«”"
        End If
    ElseIf IsKey Then
        eJournalWriteAll Nothing, content
        ShowMessage content
        If Not GetKey Then
            HandleMessage = vbCancel
        Else
            HandleMessage = vbOK
        End If
        
        If HandleMessage = vbOK Then
            eJournalWriteAll Nothing, "ƒœ»« ≈ ≈√ —…”«"
        Else
            eJournalWriteAll Nothing, "ƒ≈Õ ƒœ»« ≈ ≈√ —…”«"
        End If
    Else
        HandleMessage = vbCancel
    End If
End Function
