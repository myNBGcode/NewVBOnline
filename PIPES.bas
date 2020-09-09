Attribute VB_Name = "PIPES"
Option Explicit

Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const PIPE_ACCESS_DUPLEX = &H3
Public Const PIPE_READMODE_MESSAGE = &H2
Public Const PIPE_TYPE_MESSAGE = &H4
Public Const PIPE_WAIT = 0
Public Const PIPE_NOWAIT = 1
Public Const INVALID_HANDLE_VALUE = -1
Public Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Public Const SECURITY_DESCRIPTOR_REVISION = (1)
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Const szClientPipeName = "\\.\pipe\RAuthResponse"

Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function CreateNamedPipe Lib "kernel32" Alias _
      "CreateNamedPipeA" ( _
      ByVal lpName As String, _
      ByVal dwOpenMode As Long, _
      ByVal dwPipeMode As Long, _
      ByVal nMaxInstances As Long, _
      ByVal nOutBufferSize As Long, _
      ByVal nInBufferSize As Long, _
      ByVal nDefaultTimeOut As Long, _
      lpSecurityAttributes As Any) As Long

Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal dwRevision As Long) As Long

Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal bDaclPresent As Long, _
      ByVal pDacl As Long, _
      ByVal bDaclDefaulted As Long) As Long

Declare Function ConnectNamedPipe Lib "kernel32" ( _
      ByVal hNamedPipe As Long, _
      lpOverlapped As Any) As Long

Declare Function DisconnectNamedPipe Lib "kernel32" ( _
      ByVal hNamedPipe As Long) As Long

Declare Function CreateFile Lib "kernel32" _
    Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Declare Function WriteFile Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpBuffer As Any, _
      ByVal nNumberOfBytesToWrite As Long, _
      lpNumberOfBytesWritten As Long, _
      lpOverlapped As Any) As Long

Declare Function ReadFile Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpBuffer As Any, _
      ByVal nNumberOfBytesToRead As Long, _
      lpNumberOfBytesRead As Long, _
      lpOverlapped As Any) As Long

Declare Function FlushFileBuffers Lib "kernel32" ( _
      ByVal hFile As Long) As Long

Declare Function CloseHandle Lib "kernel32" ( _
      ByVal hObject As Long) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Declare Function CallNamedPipe Lib "kernel32" Alias _
      "CallNamedPipeA" ( _
      ByVal lpNamedPipeName As String, _
      lpInBuffer As Any, _
      ByVal nInBufferSize As Long, _
      lpOutBuffer As Any, _
      ByVal nOutBufferSize As Long, _
      lpBytesRead As Long, _
      ByVal nTimeOut As Long) As Long

Public StopWritePipe As Boolean, StopReadPipe As Boolean
Public WaitingResponseFrom As String
Type RAuth_Request
    Request_Cmd As Long
    Client_Name As String * MAX_USERNAME_LENGTH
    Client_Machine As String * MAX_COMPUTERNAME_LENGTH
End Type

Type RAuth_Response
    Response_Code As Long
    Server_Machine As String * MAX_COMPUTERNAME_LENGTH
End Type

Public Const RAuth_Get_Chief_Key = 1000
Public Const RAuth_Get_Manager_Key = 1001
Public Const RAuth_Exit = 1100

Public Const RAuth_Chief_Key_On = 1200
Public Const RAuth_Chief_Key_Off = 1201
Public Const RAuth_Manager_Key_On = 1202
Public Const RAuth_Manager_Key_Off = 1203

Public hResponsePipe As Long
Private pResponseSD As Long
Private ResponseSA As SECURITY_ATTRIBUTES

'Public Function Read_PipeResponse() As Integer
'Dim aResponse As RAuth_Response, cbnCount As Long, res As Integer
''    StopReadPipe = False
'    Read_PipeResponse = 0
'    If StopReadPipe Then Exit Function
'    cbnCount = 0
'    res = ConnectNamedPipe(hResponsePipe, ByVal 0)
'    Do
'        DoEvents
'        res = ReadFile(hResponsePipe, aResponse, LenB(aResponse), cbnCount, ByVal 0)
'        If cbnCount > 0 Then
'
'            If Trim(WaitingResponseFrom) <> ClearFixedString(aResponse.Server_Machine) Then
'                cbnCount = 0
'            End If
'        End If
'    Loop Until cbnCount > 0 Or StopReadPipe
'    res = DisconnectNamedPipe(hResponsePipe)
'    If cbnCount > 0 Then
'        Read_PipeResponse = aResponse.Response_Code
'    Else
'        Read_PipeResponse = 0
'    End If
'End Function
'
'Public Sub Create_Response_Pipe()
'    Dim res As Integer
'    Dim dwOpenMode As Long, dwPipeMode As Long
'
'    pResponseSD = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
'    res = InitializeSecurityDescriptor(pResponseSD, SECURITY_DESCRIPTOR_REVISION)
'    res = SetSecurityDescriptorDacl(pResponseSD, -1, 0, 0)
'    ResponseSA.nLength = LenB(ResponseSA)
'    ResponseSA.lpSecurityDescriptor = pResponseSD
'    ResponseSA.bInheritHandle = True
'
'    dwOpenMode = PIPE_ACCESS_DUPLEX Or FILE_FLAG_WRITE_THROUGH
'    dwPipeMode = PIPE_NOWAIT Or PIPE_TYPE_MESSAGE Or PIPE_READMODE_MESSAGE
'
'    hResponsePipe = CreateNamedPipe(szClientPipeName, dwOpenMode, dwPipeMode, _
'                              10, 10000, 2000, 10000, ResponseSA)
'
'End Sub
'
'Public Sub Close_Response_Pipe()
'      StopWritePipe = True
'      StopReadPipe = True
'      DoEvents
'      CloseHandle hResponsePipe
'      GlobalFree (pResponseSD)
'End Sub
'
'Public Sub SendRequestOverPipe(inRequestCMD As Integer, inRequestListener As String)
'Dim szSrvPipeName As String
'Dim aRequest As RAuth_Request
'Dim res As Integer, cbnCount As Long
'Dim lPipe As Long, lSD As Long
'
'    szSrvPipeName = "\\" & inRequestListener & "\pipe\RAuthRequest"
'    aRequest.Request_Cmd = inRequestCMD
'    res = GetUserName(aRequest.Client_Name, MAX_USERNAME_LENGTH)
'    res = GetComputerName(aRequest.Client_Machine, MAX_COMPUTERNAME_LENGTH)
'    WaitingResponseFrom = inRequestListener
'    StopWritePipe = False
'    StopReadPipe = False
'    Do
'        lSD = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
'        res = InitializeSecurityDescriptor(lSD, SECURITY_DESCRIPTOR_REVISION)
'        lPipe = CreateFile(szSrvPipeName, ByVal GENERIC_WRITE, ByVal FILE_SHARE_WRITE, _
'            lSD, ByVal OPEN_EXISTING, ByVal 0, ByVal 0)
'
'        res = WriteFile(lPipe, aRequest, LenB(aRequest), cbnCount, ByVal 0)
'        CloseHandle lPipe
'        GlobalFree (lSD)
'
'        If res = 1 Then Exit Do
'        DoEvents
'    Loop Until StopWritePipe
'    If StopWritePipe Then WaitingResponseFrom = ""
'End Sub
'
'
