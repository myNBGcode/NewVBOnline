Attribute VB_Name = "Errors"
Option Explicit
Public Sub NBG_error(routine_name As String, error As Integer)

    Call EventLog(1, "NBG Error: " & error & " in " & routine_name & "()")
    Call NBG_MsgBox("Runtime Error: " & error & " in " & routine_name & "()", True, " ")

End Sub
Public Sub Runtime_error(routine_name As String, error As Integer, error_msg As String)

    Call EventLog(1, "Runtime Error: " & error & " in " & routine_name & "() " & error_msg)
    Call NBG_MsgBox("Runtime Error: " & error & " in " & routine_name & "() " & error_msg, True, " ")
End Sub

Public Sub NBG_LOG_MsgBox(PStrMessage As String, _
                      Optional PBolBeep As Variant, Optional pstrTitle As Variant)
                      
' � Procedure NBG_MsgBox ������ �� ���������������
' ��� ��� �������� �������� ��������� ���� �����
'
' ���������� :
' PStrMessage �� ������ ��� ������� �� ����������
' PBolBeep    ����������� flag (True, False) ��
'             ������� �� ����� Beep default True
' PstrTitle    � ������ ��� ��������
' �.�.
'
' Call NBG_MsgBox("����������� ������� !!", True,"������ ������")

  
  If IsMissing(PBolBeep) Then PBolBeep = True
  If PBolBeep Then Beep
  If IsMissing(pstrTitle) Then pstrTitle = "On Line ��������"
  'MsgBox PStrMessage, , pstrTitle
  LogMsgbox PStrMessage, , CStr(pstrTitle)
  
  DoEvents
End Sub

Public Sub NBG_MsgBox(PStrMessage As String, _
                      Optional PBolBeep As Variant, Optional pstrTitle As Variant)
                      
' � Procedure NBG_MsgBox ������ �� ���������������
' ��� ��� �������� �������� ��������� ���� �����
'
' ���������� :
' PStrMessage �� ������ ��� ������� �� ����������
' PBolBeep    ����������� flag (True, False) ��
'             ������� �� ����� Beep default True
' PstrTitle    � ������ ��� ��������
' �.�.
'
' Call NBG_MsgBox("����������� ������� !!", True,"������ ������")

  
  If IsMissing(PBolBeep) Then PBolBeep = True
  If PBolBeep Then Beep
  If IsMissing(pstrTitle) Then pstrTitle = "On Line ��������"
  MsgBox PStrMessage, , pstrTitle
  
  DoEvents
End Sub

Public Sub Xml_ParseError(docerror As IXMLDOMParseError)
    LogMsgbox "����� ���� ����������� ��������: " & docerror.errorCode & " ������ " & docerror.Line & " ���� " & docerror.linepos & " ���� ������� " & docerror.filepos & _
        vbCrLf & " ���������� " & docerror.reason & " ������� " & docerror.srcText, vbCritical, "�����", Err
    'MsgBox "����� ���� ����������� ��������: " & docerror.errorCode & " ������ " & docerror.Line & " ���� " & docerror.linepos & " ���� ������� " & docerror.filepos & _
    '   vbCrLf & " ���������� " & docerror.reason & " ������� " & docerror.srcText, vbCritical, "�����"
End Sub

Public Sub LogMsgbox(message As String, Optional style As VbMsgBoxStyle, Optional Title As String, Optional error As Variant)
    MsgBox message, style, Title
    'eJournalWriteAll
    If IsMissing(error) Then
        eJournalWrite Title & ":" & message
    ElseIf error Is Nothing Then
        eJournalWrite Title & ":" & message
    Else
        Dim errornumber As Integer
        errornumber = error.Number
        eJournalWrite Title & ":" & message & " " & " �����:" & error.Number & " " & error.description
        If (errornumber = 999) Then
            eJournalWrite "�������! ������� � ���� ���������.����� ������ ��� ��� ���� ��� ����������."
            MsgBox "�������! ������� � ���� ���������.����� ������ ��� ��� ���� ��� ����������."
        End If
    End If
    
    
End Sub

'Public Sub LogMsgbox(error, message As String, Optional style As VbMsgBoxStyle, Optional Title As String)
'    MsgBox message, style, Title
'    'eJournalWriteAll
'    If error Is Nothing Then
'        eJournalWrite Title & ":" & message
'    Else
'        eJournalWrite Title & ":" & message & " " & " �����:" & error.Number & " " & error.description
'    End If
'
'
'End Sub

