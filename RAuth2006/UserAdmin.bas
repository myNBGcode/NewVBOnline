Attribute VB_Name = "UserAdmin"
Option Explicit
Option Base 0     ' Important assumption for this code
Const MAX_PREFERRED_LENGTH = -1
   
Public Declare Function NetMessageBufferSend Lib "Netapi32" (ByVal servername As String, ByVal msgname As String, ByVal fromname As String, ByVal buf As Any, ByVal BufLen As Integer) As Long

Private Type GROUP_USER_INFO_0
    gusri0_name As Long           'LPWSTR in SDK
End Type

Private Type USER_INFO_3
    usri3_name As Long           'LPWSTR in SDK
    usri3_password As Long       'LPWSTR in SDK
    usri3_password_age As Long      'DWORD in SDK
    usri3_priv As Long           'DWORD in SDK
    usri3_home_dir As Long       'LPWSTR in SDK
    usri3_comment As Long        'LPWSTR in SDK
    usri3_flags As Long          'DWORD in SDK
    usri3_script_path As Long    'LPWSTR in SDK
    usri3_auth_flags As Long        'DWORD in SDK
    usri3_full_name As Long         'LPWSTR in SDK
    usri3_usr_comment As Long    'LPWSTR in SDK
    usri3_parms As Long          'LPWSTR in SDK
    usri3_workstations As Long      'LPWSTR in SDK
    usri3_last_logon As Long        'DWORD in SDK
    usri3_last_logoff As Long    'DWORD in SDK
    usri3_acct_expires As Long      'DWORD in SDK
    usri3_max_storage As Long    'DWORD in SDK
    usri3_units_per_week As Long    'DWORD in SDK
    usri3_logon_hours As Long    'PBYTE in SDK
    usri3_bad_pw_count As Long      'DWORD in SDK
    usri3_num_logons As Long        'DWORD in SDK
    usri3_logon_server As Long      'LPWSTR in SDK
    usri3_country_code As Long      'DWORD in SDK
    usri3_code_page As Long         'DWORD in SDK
    usri3_user_id As Long        'DWORD in SDK
    usri3_primary_group_id As Long  'DWORD in SDK
    usri3_profile As Long        'LPWSTR in SDK
    usri3_home_dir_drive As Long    'LPWSTR in SDK
    usri3_password_expired As Long  'DWORD in SDK
End Type

   
Declare Function NetGetDCName Lib "netapi32.dll" (servername As Byte, DomainName As Byte, DCNPtr As Long) As Long

' Enumerate using Level 0 user structure
Declare Function NetUserEnum0 Lib "netapi32.dll" Alias "NetUserEnum" (servername As Byte, ByVal Level As Long, ByVal lFilter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, resumehandle As Long) As Long

Declare Function NetGroupEnumUsers0 Lib "netapi32.dll" Alias "NetGroupGetUsers" (servername As Byte, GroupName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, resumehandle As Long) As Long

Declare Function NetGroupEnum0 Lib "netapi32.dll" Alias "NetGroupEnum" (servername As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, resumehandle As Long) As Long

Declare Function NetUserGetGroups0 Lib "netapi32.dll" Alias "NetUserGetGroups" (servername As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long

Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long

Declare Function NetAPIBufferAllocate Lib "netapi32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long

Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (Retval As Byte, ByVal Ptr As Long) As Long

Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)

Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, source As Byte) As Long

Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (Retval As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long

Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Declare Function GetUserName_ Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal nFormat As Byte, ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function NetUserGetInfo Lib "netapi32.dll" (strServerName As Any, strUserName As Any, ByVal dwLevel As Long, pBuffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

' Converts a Unicode string to an ANSI string
' Specify -1 for cchWideChar and 0 for cchMultiByte to return string length.
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Const CP_ACP = 0        ' ANSI code page

Public Const MAX_COMPUTERNAME_LENGTH As Long = 15&
Public Const MAX_USERNAME_LENGTH As Long = 25&

Type MungeLong
     x As Long
     Dummy As Integer
End Type

Type MungeInt
     XLo As Integer
     XHi As Integer
     Dummy As Integer
End Type

'Type TUser0                    ' Level 0
'     ptrName As Long
'End Type
'
'Type TUser1                    ' Level 1
'     ptrName As Long
'     ptrPassword As Long
'     dwPasswordAge As Long
'     dwPriv As Long
'     ptrHomeDir As Long
'     ptrComment As Long
'     dwFlags As Long
'     ptrScriptPath As Long
'End Type

   '
   ' for dwPriv
   '
Const USER_PRIV_MASK = &H3
Const USER_PRIV_GUEST = &H0
Const USER_PRIV_USER = &H1
Const USER_PRIV_ADMIN = &H2

   '
   ' for dwFlags
   '
Const UF_SCRIPT = &H1
Const UF_ACCOUNTDISABLE = &H2
Const UF_HOMEDIR_REQUIRED = &H8
Const UF_LOCKOUT = &H10
Const UF_PASSWD_NOTREQD = &H20
Const UF_PASSWD_CANT_CHANGE = &H40
Const UF_NORMAL_ACCOUNT = &H200     ' Needs to be ORed with the other flags

   '
   ' for lFilter
   '
Const FILTER_NORMAL_ACCOUNT = &H2

Public isTeller As Boolean
Public isChiefTeller As Boolean
Public isManager As Boolean
Public isImportUser As Boolean
   
Public UserGroups As New Collection
'Public GroupUsers As New Collection

Dim GetPDCNameFailed As Boolean
   
   
Private Declare Function NetUserGetInfo_V2 Lib "netapi32.dll" Alias "NetUserGetInfo" (ByVal strServerName As Any, ByVal strUserName As Any, ByVal dwLevel As Long, pBuffer As Long) As Long

Declare Function NetGetDCName_V2 Lib "netapi32.dll" Alias "NetGetDCName" (ByVal servername As Any, ByVal DomainName As Any, DCNPtr As Long) As Long


   
Public Function CurrentMachineName() As String
    
    If cClientName = "" Then
        Dim lSize As Long, sBuffer As String
        sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
        lSize = Len(sBuffer)
        If GetComputerName(sBuffer, lSize) Then CurrentMachineName = Left$(sBuffer, lSize)
    Else
        CurrentMachineName = cClientName
    End If
    
    UpdatexmlEnvironment "COMPUTERNAME", CurrentMachineName
End Function
   
Function GetPrimaryDCName(ByVal MName As String, ByVal DName As String) As String
   Dim Result As Long, DCName As String, DCNPtr As Long
   Dim DCNArray(100) As Byte
     If LocalFlag Then Exit Function
     If GetPDCNameFailed Then Exit Function
     
     MName = StrConv(MName & vbNullChar, vbUnicode)
     DName = StrConv(DName & vbNullChar, vbUnicode)
     
     Result = NetGetDCName_V2(MName, DName, DCNPtr)
     If Result <> 0 Then GetPDCNameFailed = True: MsgBox "Error GetPDCName: " & Result: Exit Function
     
     Result = PtrToStr(DCNArray(0), DCNPtr)
     Result = NetAPIBufferFree(DCNPtr)
     DCName = DCNArray()
     GetPrimaryDCName = DCName
End Function

Public Sub GetUserGroups(ByVal sName As String, inUName As String)
Dim Result As Long, BufPtr As Long, BufLen As Long, EntriesRead As Long, TotalEntries As Long, SNArray() As Byte, UNArray() As Byte, GNArray(99) As Byte, gName As String, i As Integer, TempPtr As MungeLong, TempStr As MungeInt

    SNArray = sName & vbNullChar       ' Move to byte array
    UNArray = inUName & vbNullChar     ' Move to Byte array
    BufLen = 2048                       ' Buffer size
    
    Result = NetUserGetGroups0(SNArray(0), UNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries)
    If Result <> 0 Then
        'And Result <> 234 Then    ' 234 means multiple reads required
         MsgBox "Error NetUSerGetGroups: " & Result & " enumerating user " & EntriesRead & " of " & TotalEntries, vbOKOnly, "On Line Εφαρμογή"
         If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & gName & "'"
         Exit Sub
    End If
    For i = UserGroups.Count To 1 Step -1
        UserGroups.Remove (i)
    Next i
    
    For i = 1 To EntriesRead
        ' Get pointer to string from beginning of buffer
        ' Copy 4-byte block of memory in 2 steps
        Result = PtrToInt(TempStr.XLo, BufPtr + (i - 1) * 4, 2)
        Result = PtrToInt(TempStr.XHi, BufPtr + (i - 1) * 4 + 2, 2)
        LSet TempPtr = TempStr ' munge 2 integers into a Long
        ' Copy string to array
        Result = PtrToStr(GNArray(0), TempPtr.x)
        gName = Left(GNArray, StrLen(TempPtr.x))
        UserGroups.Add gName
    Next i

    Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory
End Sub

Public Function GetStrFromPtrW(lpszW As Long) As String
  Dim sRtn As String

  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char

' WideCharToMultiByte also returns Unicode string length
'  sRtn = String$(WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, 0, 0, 0, 0), 0)

  Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
  GetStrFromPtrW = GetStrFromBufferA(sRtn)

End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

Private Sub TestGetName(ByVal sName As String)
'εως 03/11/2000
    Dim Result As Long
    Dim pServer() As Byte, pUser() As Byte
    
    If LocalFlag Then Exit Sub

'    Dim sDomainDCName As String
    'You need to change "Your_User_Account_Domain" in the next
    'line to your valid NT User domain name.
'    sDomainDCName = GetPrimaryDCName("", "")
'    MsgBox sDomainDCName

    'You need to change "Your_Domain_Logon_Name" in the next
    'line to your valid NT domain logon name.

    pUser = cUserName & vbNullChar
    pServer = sName & vbNullChar
'    sDomainDCName & vbNullChar

    'The above two lines convert VB string to Unicode string.

    'To check a local user account on local machine
    'pUser = "Local_User_Name" & vbNullChar
    'pServer = "\\Machine_Name" & vbNullChar

    Dim dwLevel As Long
    dwLevel = 3

    Dim tmpBuffer As USER_INFO_3
    Dim ptmpBuffer As Long

    'As last param is dimmed as long, the pointer to ptmpBuffer
    'is passed to dll, and the function returns a pointer to a
    'pointer to our UDT. Therefore ptmpBuffer on return holds a
    'pointer to our UDT.
    Result = NetUserGetInfo(pServer(0), pUser(0), dwLevel, _
             ptmpBuffer)
    If Result <> 0 Then
      MsgBox "Error NetUserGetInfo: " & Result, vbOKOnly, "On Line Εφαρμογή"
      Exit Sub
    End If

    'Deference it!!!
    MoveMemory tmpBuffer, ByVal ptmpBuffer, Len(tmpBuffer)
''    CopyMemory tmpBuffer, ptmpBuffer, LenB(tmpBuffer)

    'Convert LPWSTR (Unicode string) to VB string.
'    CopyMemory sByte(0), tmpBuffer.usri3_name, 256
'    sUser = sByte
'    sUser = sUser & vbNullChar

    Dim fName As String

    fName = GetStrFromPtrW(tmpBuffer.usri3_full_name)

'    GetStrFromPtrW (ui3.usri3_name)
'    CopyMemory fByte(0), tmpBuffer.usri3_full_name, 256
'    fName = fByte
'    fName = fName & vbNullChar

    cFullUserName = fName: cFullUserName = ClearFixedString(cFullUserName)

    'Now I get my user name back, it's VB string now'
    Result = NetAPIBufferFree(ptmpBuffer)

    If Result <> 0 Then
      MsgBox "Error: " & Result, vbOKOnly, "On Line Εφαρμογή"
      Exit Sub
    End If
'εως 03/11/2000
    
'    Dim Result As Long
'    Dim pServer As String, pUser As String
'
'    Dim sDomainDCName As String
'    sDomainDCName = GetPrimaryDCName("", "")
'
'    pUser = StrConv(cUserName & vbNullChar, vbUnicode)
'    pServer = StrConv(sDomainDCName & vbNullChar, vbUnicode)
'
'    Dim dwLevel As Long
'    dwLevel = 3
'
'    Dim tmpBuffer As USER_INFO_3, ptmpBuffer As Long, fName As String
'
'    Result = NetUserGetInfo_V2(pServer, pUser, dwLevel, ptmpBuffer)
'    If Result <> 0 Then MsgBox "Error: " & Result, vbOKOnly, "On Line Εφαρμογή": Exit Sub
'
'    fName = GetStrFromPtrW(tmpBuffer.usri3_full_name)
'
'    cFullUserName = fName: cFullUserName = ClearFixedString(cFullUserName)
'
'    Result = NetAPIBufferFree(ptmpBuffer)
'
'    If Result <> 0 Then MsgBox "Error: " & Result, vbOKOnly, "On Line Εφαρμογή": Exit Sub
    
End Sub
   

Public Function GetUserName() As String
Dim lpBuff As String * MAX_USERNAME_LENGTH
Dim ret As Long
    
    ret = GetUserName_(lpBuff, MAX_USERNAME_LENGTH)
    GetUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function
   
   

Public Sub ChkUser()
Dim lpBuff As String * MAX_USERNAME_LENGTH
Dim lpAllBuff As String * MAX_USERNAME_LENGTH
Dim ret As Long
Dim res As String
    If LocalFlag Then
        isTeller = True
        isChiefTeller = True
        isManager = True
        isImportUser = True
        Exit Sub
    End If

    ret = GetUserName_(lpBuff, MAX_USERNAME_LENGTH)
    cUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    cUserName = UCase(cUserName)
    UpdatexmlEnvironment "USERNAME", cUserName
    
'    res = cPDC
   
'    TestGetName cPDC
'    GetUserGroups cPDC, cUserName

    TestGetName cLogonServer
    GetUserGroups cLogonServer, cUserName
    
    isTeller = False
    isChiefTeller = False
    isManager = False
    isImportUser = False
    
Dim i As Integer
    For i = 1 To UserGroups.Count
        If UCase(UserGroups(i)) = "TELLER" Then isTeller = True
        If UCase(UserGroups(i)) = "CHIEF TELLER" Then isChiefTeller = True
        If UCase(UserGroups(i)) = "MANAGER" Then isManager = True
        If UCase(UserGroups(i)) = "IMPORT USERS" Then isImportUser = True
    Next i
    
    UpdatexmlEnvironment "TELLER", CStr(isTeller)
    UpdatexmlEnvironment "CHIEFTELLER", CStr(isChiefTeller)
    UpdatexmlEnvironment "MANAGER", CStr(isManager)
    UpdatexmlEnvironment "IMPORTUSER", CStr(isImportUser)
    
End Sub
Public Function ChkChangePwd(inNewPwd1 As String, inNewPwd2 As String) As Boolean
    If Trim(inNewPwd1) <> Trim(inNewPwd2) Then
       ChkChangePwd = False
    Else
       ChkChangePwd = True
    End If
End Function

Public Function ClearFixedString(inString As String) As String
    If InStr(inString, Chr$(0)) > 0 Then
        ClearFixedString = Left(inString, InStr(inString, Chr$(0)) - 1)
    Else
        ClearFixedString = inString
    End If
End Function

'Public Function GetUserList(ByVal sName As String, ByVal gName As String) As Boolean
'    Dim Result As Long
'    Dim pServer() As Byte, pGroup() As Byte
'    Dim EntriesRead As Long, TotalEntries As Long, resumehandle As Long
'    Dim aname As String, i As Integer, TempPtr As MungeLong, TempStr As MungeInt
'    Dim GNArray(99) As Byte
'
'    pServer = sName & vbNullChar
'    pGroup = gName & vbNullChar
'
'    Dim tmpBuffer As GROUP_USER_INFO_0
'    Dim ptmpBuffer As Long
'
'    Result = NetGroupEnumUsers0(pServer(0), pGroup(0), 0, ptmpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, TotalEntries, resumehandle)
'
'    For i = GroupUsers.Count To 1 Step -1
'        GroupUsers.Remove (i)
'    Next i
'
'    For i = 1 To EntriesRead
'        ' Get pointer to string from beginning of buffer
'        ' Copy 4-byte block of memory in 2 steps
'        Result = PtrToInt(TempStr.XLo, ptmpBuffer + (i - 1) * 4, 2)
'        Result = PtrToInt(TempStr.XHi, ptmpBuffer + (i - 1) * 4 + 2, 2)
'        LSet TempPtr = TempStr ' munge 2 integers into a Long
'        ' Copy string to array
'        Result = PtrToStr(GNArray(0), TempPtr.x)
'        aname = Left(GNArray, StrLen(TempPtr.x))
'        GroupUsers.Add aname
'    Next i
'    Result = NetAPIBufferFree(ptmpBuffer)         ' Don't leak memory
'
'End Function
