Attribute VB_Name = "GK_DLL"
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const INFINITE = -1&

Declare Function VB4SLICONNECT Lib "VB4SLI.DLL" (pLUName As String, pAppId As String, ByVal pConvertIt As Long, ByVal pTimeOut As Long, pRet1 As Long, pRet2 As Long, pRet3 As Long, ByVal pDebug As Long) As String

Declare Function VB4SLICONNECTEX Lib "VB4SLI.DLL" (pLUName As String, pAppId As String, ByVal pConvertIt As Long, ByVal pTimeOut As Long, pRet1 As Long, pRet2 As Long, pRet3 As Long, ByVal pDebug As Long, ByRef exEvent As Long) As String

Declare Function VB4SLIDISCONNECT Lib "VB4SLI.DLL" (ByVal pTimeOut As Long, pRet1 As Long, pRet2 As Long, pRet3 As Long, ByVal pDebug As Long) As String

Declare Function VB4SLISEND Lib "VB4SLI.DLL" (pData As String, ByVal pConvertIt As Long, ByVal pTimeOut As Long, pLen As Long, pMsgType As Long, pRet1 As Long, pRet2 As Long, pRet3 As Long, ByVal pDebug As Long) As String

Declare Function VB4SLIRECEIVE Lib "VB4SLI.DLL" (ByVal pConvertIt As Long, ByVal pTimeOut As Long, pLen As Long, pMsgType As Long, pRet1 As Long, pRet2 As Long, pRet3 As Long, ByVal pDebug As Long) As String

Declare Function VB4SLIReset Lib "VB4SLI.DLL" (ByVal ResetType As Long) As Integer

Declare Function VB4SLIWAIT Lib "VB4SLI.DLL" (ByVal Seconds As Long) As Integer
        
'Public Const BETB = 1
'Public Const SEND = 2
'Public Const RECV = 3
'Public pLUADirection As Long

'Global Connected As Integer
'Global Batch As Integer
'Global pBuff As String


Declare Function GKUpper Lib "GK_VB4.DLL" (Str As String) As String
Declare Function GKTranslate Lib "GK_VB4.DLL" (Str As String, SrcTable As String, TgtTable As String) As String
'Declare Function GKEncrypt Lib "GK_VB4.DLL" (Str As String, Password As String) As String
'Declare Function GKDecrypt Lib "GK_VB4.DLL" (Str As String, Password As String) As String
'Declare Function GKNetUse Lib "GK_VB4.DLL" '        (Device As String, '         Netpath As String, '         UserId As String, '          Password As String) '        As Long
'Declare Function GKNetUnUse Lib "GK_VB4.DLL" '        (Netpath As String, '         ForceDisconnect As Long) '         As Long
'Declare Function GKFindFirst Lib "GK_VB4.DLL" '        (FilePattern As String, '         FileSizeLow As Long, '         FileSizeHigh As Long, '         Attribs As Long) '         As String
'Declare Function GKFindNext Lib "GK_VB4.DLL" '        (FileSizeLow As Long, '         FileSizeHigh As Long, '         Attribs As Long) '         As String
'Declare Function GKFindClose Lib "GK_VB4.DLL" '        (Result As Long) '        As Long
'Declare Function GKDirChange Lib "GK_VB4.DLL" '        (Dir As String, '         RetCode As Long) '         As String
'Declare Function GKNetName Lib "GK_VB4.DLL" '        (RetCode As Long) '        As String
'Declare Function GKWriteToRegClassesRoot Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegVal As String) As Long
'Declare Function GKWriteToRegLocalMachine Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegVal As String) As Long
'Declare Function GKWriteToRegCurrentUser Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegVal As String) As Long
'Declare Function GKWriteToRegUsers Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegVal As String) As Long
'Declare Function GKReadFromRegClassesRoot Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegLen As Long) As String
'Declare Function GKReadFromRegLocalMachine Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegLen As Long) As String
'Declare Function GKReadFromRegCurrentUser Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegLen As Long) As String
'Declare Function GKReadFromRegUsers Lib "GK_VB4.DLL" '        (RegEntry As String, '         RegKey As String, '         RegLen As Long) As String
Declare Function GKReportAnEvent Lib "GK_VB4.DLL" (EventType As Long, Line1 As String, Line2 As String) As Long

' EvntType  EVENTLOG_SUCCESS                0
'           EVENTLOG_ERROR_TYPE             1
'           EVENTLOG_WARNING_TYPE           2
'           EVENTLOG_INFORMATION_TYPE       4
'           EVENTLOG_AUDIT_SUCCESS          8
'           EVENTLOG_AUDIT_FAILURE         10
'
'Declare Function GKLogon Lib "GK_VB4.DLL" '        (UserName As String, '         Password As String, '         UserToken As Long) As Long
'
'Declare Function GKLogoff Lib "GK_VB4.DLL" '        (UserToken As Long) As Long
'
'Declare Function GKSaveWindow Lib "GK_VB4.DLL" '        (FileName As String, '         ByVal hWnd As Long, '         DIBKind As Long) As Long
'
'Declare Function GKIsInGroup Lib "GK_VB4.DLL" '        (GroupName As String) As Boolean

'DIBKind  BI_RLE4     2
'         BI_RLE8     1
'         BI_RGB      0



Public Function EventLog(EventType As Long, string1 As String, Optional string2 As Variant) As String

Dim Retval As Long, Line1 As String, Line2 As String, Beep As Boolean
    
' EvntType  EVENTLOG_SUCCESS                0
'           EVENTLOG_ERROR_TYPE             1
'           EVENTLOG_WARNING_TYPE           2
'           EVENTLOG_INFORMATION_TYPE       4
'           EVENTLOG_AUDIT_SUCCESS          8
'           EVENTLOG_AUDIT_FAILURE         10
'

If IsMissing(string2) Then string2 = ""

If EventLogWrite Then
    If cb.app_debug = 1 Or cb.app_debug = 2 Then
        Line1 = string1 + Chr$(13) + Chr$(10)
        Line2 = string2 + Chr$(13) + Chr$(10)
        EventLog = GKReportAnEvent(EventType, Line1, Line2)
    End If
End If

If cb.app_debug = 2 Then
    If EventType <> 0 Then Beep = True Else Beep = False
'    Call NBG_MsgBox("EVENT LOG: " & Chr$(13) & Chr$(10) & Line1 & Line2, Beep)
End If

End Function


