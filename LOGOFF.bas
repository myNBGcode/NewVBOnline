Attribute VB_Name = "LOGOFF"
Option Explicit
Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_FORCE As Long = 4
Private Const EWX_POWEROFF As Long = 8

'The ExitWindowsEx function either logs off, shuts down, or shuts
'down and restarts the system.
Private Declare Function ExitWindowsEx Lib "USER32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

'The GetLastError function returns the calling thread's last-error
'code value. The last-error code is maintained on a per-thread basis.
'Multiple threads do not overwrite each other's last-error code.
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const mlngWindows95 = 0
Private Const mlngWindowsNT = 1

Public glngWhichWindows32 As Long

'The GetVersion function returns the operating system in use.
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Type LUID
   UsedPart As Long
   IgnoredForNowHigh32BitPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   TheLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   TheLuid As LUID
   Attributes As Long
End Type

'The GetCurrentProcess function returns a pseudohandle for the
'current process.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'The OpenProcessToken function opens the access token associated with
'a process.
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'The LookupPrivilegeValue function retrieves the locally unique
'identifier (LUID) used on a specified system to locally represent
'the specified privilege name.
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

'The AdjustTokenPrivileges function enables or disables privileges
'in the specified access token. Enabling or disabling privileges
'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Sub AdjustToken()

'********************************************************************
'* This procedure sets the proper privileges to allow a log off or a
'* shut down to occur under Windows NT.
'********************************************************************

   Const TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8
   Const SE_PRIVILEGE_ENABLED = &H2

   Dim hdlProcessHandle As Long
   Dim hdlTokenHandle As Long
   Dim tmpLuid As LUID
   Dim tkp As TOKEN_PRIVILEGES
   Dim tkpNewButIgnored As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long

   'Set the error code of the last thread to zero using the
   'SetLast Error function. Do this so that the GetLastError
   'function does not return a value other than zero for no
   'apparent reason.
   SetLastError 0

   'Use the GetCurrentProcess function to set the hdlProcessHandle
   'variable.
   hdlProcessHandle = GetCurrentProcess()

   If GetLastError <> 0 Then
      MsgBox "GetCurrentProcess error==" & GetLastError
   End If

   OpenProcessToken hdlProcessHandle, _
      (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle

   If GetLastError <> 0 Then
      MsgBox "OpenProcessToken error==" & GetLastError
   End If

   'Get the LUID for shutdown privilege
   LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

   If GetLastError <> 0 Then
      MsgBox "LookupPrivilegeValue error==" & GetLastError
   End If

   tkp.PrivilegeCount = 1    ' One privilege to set
   tkp.TheLuid = tmpLuid
   tkp.Attributes = SE_PRIVILEGE_ENABLED

   'Enable the shutdown privilege in the access token of this process
   AdjustTokenPrivileges hdlTokenHandle, _
                         False, _
                         tkp, _
                         Len(tkpNewButIgnored), _
                         tkpNewButIgnored, _
                         lBufferNeeded

   If GetLastError <> 0 Then
      MsgBox "AdjustTokenPrivileges error==" & GetLastError
   End If

End Sub

Public Sub cmdLogoff_()

   ExitWindowsEx (EWX_LogOff), &HFFFF
   MsgBox "ExitWindowsEx's GetLastError " & GetLastError

End Sub

Public Sub cmdForceLogoff_()
    MsgBox "Επανεκκίνηση Συστήματος...", vbCritical, "ΠΡΟΣΟΧΗ !!!!"
    DISCONNECT_
    ExitWindowsEx (EWX_LogOff Or EWX_FORCE), &HFFFF
'   MsgBox "ExitWindowsEx's GetLastError " & GetLastError

End Sub

Private Sub cmdShutdown_()

   If glngWhichWindows32 = mlngWindowsNT Then
      AdjustToken
      MsgBox "Post-AdjustToken GetLastError " & GetLastError
   End If

   ExitWindowsEx (EWX_SHUTDOWN), &HFFFF
   MsgBox "ExitWindowsEx's GetLastError " & GetLastError

End Sub

Private Sub cmdForceShutdown_()
   If glngWhichWindows32 = mlngWindowsNT Then
      AdjustToken
      MsgBox "Post-AdjustToken GetLastError " & GetLastError
   End If

ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), &HFFFF
MsgBox "ExitWindowsEx's GetLastError " & GetLastError

End Sub

Private Sub Initialize_()
'********************************************************************
'* When the project starts, check the operating system used by
'* calling the GetVersion function.
'********************************************************************
   Dim lngVersion As Long

   lngVersion = GetVersion()

   If ((lngVersion And &H80000000) = 0) Then
      glngWhichWindows32 = mlngWindowsNT
      MsgBox "Running Windows NT or Windows 2000"
   Else
      glngWhichWindows32 = mlngWindows95
      MsgBox "Running Windows 95 or 98"
   End If

End Sub

