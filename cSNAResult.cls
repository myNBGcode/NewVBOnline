VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSNAResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ErrCode As Long
Public ErrMessage As String
Public SenseCodeMessage As String
Public MessageID As String

Public Sub UpdateXmlDocumentManager(aobj As cXMLDocumentManager)
    If (ErrCode <> 0 And ErrCode <> COM_OK) Or Me.SenseCodeMessage <> "" Then
        If Not (aobj Is Nothing) Then
            If Not (aobj.activeform Is Nothing) Then
                On Error Resume Next
                aobj.activeform.sbWriteStatusMessage "(" & CStr(ErrCode) & ") " & ErrMessage
                eJournalWrite "диабибасг стоивеиым: (" & CStr(ErrCode) & ") " & ErrMessage
                If Me.SenseCodeMessage <> "" Then
                    eJournalWrite Me.SenseCodeMessage
                End If
            End If
        End If
    Else
        If Not (aobj Is Nothing) Then
            If Not (aobj.activeform Is Nothing) Then
                On Error Resume Next
                aobj.activeform.sbWriteStatusMessage "H епийоимымиа окойкгяыхгйе...."
            End If
        End If
    End If
    If Not (aobj Is Nothing) Then
        If Not (aobj.activeform Is Nothing) Then
            aobj.activeform.Enabled = True
        End If
    End If
End Sub

Public Sub UpdateForm(aobj)
    If (ErrCode <> 0 And ErrCode <> COM_OK) Or Me.SenseCodeMessage <> "" Then
        If Not (aobj Is Nothing) Then
            On Error Resume Next
            ShowStatusMessage "(" & CStr(ErrCode) & ") " & ErrMessage
            eJournalWrite "диабибасг стоивеиым: (" & CStr(ErrCode) & ") " & ErrMessage
            If Me.SenseCodeMessage <> "" Then
                eJournalWrite Me.SenseCodeMessage
            End If
        End If
    Else
        If Not (aobj Is Nothing) Then
            On Error Resume Next
            aobj.sbWriteStatusMessage "H епийоимымиа окойкгяыхгйе...."
        End If
    End If
    If Not (aobj Is Nothing) Then aobj.Enabled = True
End Sub



