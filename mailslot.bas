Attribute VB_Name = "mailslot"
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, _
    ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, _
    ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Const OPEN_EXISTING = 3
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Function SendToWinPopUp(PopFrom As String, PopTo As String, MsgText As String) As Long
' parms: PopFrom: user or computer that sends the message
' PopTo: computer that receives the  message
' MsgText: the text of the message to send
    Dim rc As Long
    Dim mshandle As Long
    Dim msgtxt As String
    Dim byteswritten As Long
    Dim mailslotname As String
' name of the mailslot
    mailslotname = "\\" + PopTo + "\mailslot\messngr"
    msgtxt = PopFrom + Chr(0) + PopTo + Chr(0) + _
    MsgText + Chr(0)
    mshandle = CreateFile(mailslotname, GENERIC_WRITE, FILE_SHARE_READ, 0, _
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    rc = WriteFile(mshandle, msgtxt, Len(msgtxt), byteswritten, 0)
    rc = CloseHandle(mshandle)
End Function
