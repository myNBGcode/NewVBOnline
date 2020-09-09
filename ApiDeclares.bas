Attribute VB_Name = "APIDeclares"
Option Explicit

Type POINTAPI
    x As Long
    Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Const GWL_WNDPROC = -4

Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMPACTING = &H41
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_COPYDATA = &H4A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DEVICECHANGE = &H219
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_DROPFILES = &H233
Public Const WM_ENDSESSION = &H16
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_ERASEBKGND = &H14
Public Const WM_EXITMENULOOP = &H212
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_MOVING = &H216
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NOTIFY = &H4E
Public Const WM_OTHERWINDOWCREATED = &H42
Public Const WM_OTHERWINDOWDESTROYED = &H43
Public Const WM_PAINT = &HF
Public Const WM_PALETTECHANGED = &H311
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_POWER = &H48
Public Const WM_POWERBROADCAST = &H218
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_QUERYOPEN = &H13
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_SIZE = &H5
Public Const WM_SIZING = &H214
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMECHANGE = &H1E
Public Const WM_USERCHANGED = &H54
Public Const WM_VSCROLL = &H115
Public Const WM_WININICHANGE = &H1A

' these are used in mouse messages
Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4

' these are used in menu messages
Public Const MF_APPEND = &H100&
Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DEFAULT = &H1000&
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20& ' columns with a separator line
Public Const MF_MENUBREAK = &H40&    ' columns w/o a separator line
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
Public Const MF_RIGHTJUSTIFY = &H4000&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MFS_DEFAULT = &H1000&

Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20


' the possible return values for WM_MOUSEACTIVATE
Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4

' these are used by the WM_MOVING and WM_SIZING messages
Public Const WMSZ_BOTTOM = 6
Public Const WMSZ_BOTTOMLEFT = 7
Public Const WMSZ_BOTTOMRIGHT = 8
Public Const WMSZ_LEFT = 1
Public Const WMSZ_RIGHT = 2
Public Const WMSZ_TOP = 3
Public Const WMSZ_TOPLEFT = 4
Public Const WMSZ_TOPRIGHT = 5

' this are the values that can be returned by a WM_NCHITTEST message
Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTCLOSE = 20
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHELP = 21
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0
Public Const HTOBJECT = 19
Public Const HTREDUCE = HTMINBUTTON
Public Const HTRIGHT = 11
Public temp As Long, lpPrevWndProc As Long

Public Const HTSIZE = HTGROWBOX
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSYSMENU = 3
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTTRANSPARENT = (-1)
Public Const HTVSCROLL = 7
Public Const HTZOOM = HTMAXBUTTON

' these values are passed to a WM_SIZE message
Public Const SIZE_MAXHIDE = 4
Public Const SIZE_MAXIMIZED = 2
Public Const SIZE_MAXSHOW = 3
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_RESTORED = 0

' this values are related to the WM_SYSCOMMAND message
Public Const SC_ARRANGE = &HF110
Public Const SC_CLOSE = &HF060
Public Const SC_CONTEXTHELP = &HF180
Public Const SC_DEFAULT = &HF160
Public Const SC_HOTKEY = &HF150
Public Const SC_HSCROLL = &HF080
Public Const SC_KEYMENU = &HF100
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const SC_MONITORPOWER = &HF170
Public Const SC_MOUSEMENU = &HF090
Public Const SC_MOVE = &HF010
Public Const SC_NEXTWINDOW = &HF040
Public Const SC_PREVWINDOW = &HF050
Public Const SC_RESTORE = &HF120
Public Const SC_SCREENSAVE = &HF140
Public Const SC_SIZE = &HF000
Public Const SC_TASKLIST = &HF130
Public Const SC_VSCROLL = &HF070
Public Const SC_ZOOM = SC_MAXIMIZE

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
' used to implement drag-and-drop recipients
Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
' used within the WM_PAINT message
Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function AddFontResource Lib "GDI32" Alias "AddFontResourceA" (ByVal lpszFilename As String) As Long

Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long

' used for dealing with menus
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Any, ByVal lpNewItem As Any) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
  
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As _
    Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle _
    As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias _
    "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
    ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal _
    TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
    PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As _
    Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As _
    Long, ByVal uExitCode As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const HWND_BROADCAST = &HFFFF

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" _
(ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" _
(ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" _
(ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal command As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
(ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
(ByVal lpString1 As String, ByVal lpString2 As Any) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Const LOCALE_SDECIMAL              As Long = &HE     'decimal separator
Public Const LOCALE_STHOUSAND             As Long = &HF     'thousand separator

Public Declare Function GetThreadLocale Lib "kernel32" () As Long

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
   
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
Public Const TH32CS_SNAPPROCESS As Long = 2&
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long



Public Function GenerateGUID() As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim Retval As Long
    lLen = 40
    bGUID = String(lLen, 0)
    CoCreateGuid uGUID
    'Convert the structure into a displayable string
    Retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(Mid$(sGUID, Retval, 1)) = 0) Then Retval = Retval - 1
    GenerateGUID = Left$(sGUID, Retval)
End Function

Public Function GetGuid() As String
    Dim strGUID As String
    strGUID = GenerateGUID
    strGUID = Replace(strGUID, "{", "")
    strGUID = Replace(strGUID, "}", "")
    GetGuid = strGUID
End Function
Public Sub globalSelectPrinter(NewPrinter As String)
    Dim Prt As Printer
    
    For Each Prt In Printers
        If Prt.DeviceName = NewPrinter Then
            Set Printer = Prt
        Exit For
        End If
    Next
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_COMMAND Then Debug.Print "Message: "; hw, uMsg, wParam, lParam
'    Debug.Print "Message: "; hw, uMsg, wParam, lParam
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

Public Sub SetOnLineLocaleInfo()
    Dim LCID As Long
    Dim decimalSeparator As String
    Dim groupSeparator As String
    Dim Result As Long
    LCID = GetSystemDefaultLCID()
    decimalSeparator = GetUserLocaleInfo(LCID, LOCALE_SDECIMAL)
    groupSeparator = GetUserLocaleInfo(LCID, LOCALE_STHOUSAND)
    If (decimalSeparator <> ",") Then
        Result = SetLocaleInfo(LCID, LOCALE_SDECIMAL, ",")
    End If
    If (groupSeparator <> ".") Then
        Result = SetLocaleInfo(LCID, LOCALE_STHOUSAND, ".")
    End If
End Sub


Function KillProcess(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) _
    As Boolean
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
    
    ' Windows NT/2000 require a special treatment
    ' to ensure that the calling process has the
    ' privileges to shut down the system
    
    ' under NT the high-order bit (that is, the sign bit)
    ' of the value retured by GetVersion is cleared
    If GetVersion() >= 0 Then
        ' open the tokens for the current process
        ' exit if any error
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY, hToken) = 0 Then
            GoTo CleanUp
        End If
        
        ' retrieves the locally unique identifier (LUID) used
        ' to locally represent the specified privilege name
        ' (first argument = "" means the local system)
        ' Exit if any error
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            GoTo CleanUp
        End If
    
        ' complete the TOKEN_PRIVILEGES structure with the # of
        ' privileges and the desired attribute
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
    
        ' try to acquire debug privilege for this process
        ' exit if error
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, _
            ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
    End If
    
    ' now we can finally open the other process
    ' while having complete access on its attributes
    ' exit if any error
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        ' call was successful, so we can kill the application
        ' set return value for this function
        KillProcess = (TerminateProcess(hProcess, ExitCode) <> 0)
        ' close the process handle
        CloseHandle hProcess
    End If
    
    If GetVersion() >= 0 Then
        ' under NT restore original privileges
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
        
CleanUp:
        If hToken Then CloseHandle hToken
    End If
    
End Function

Public Function KillProcesses(processName) As String
    Dim i As Integer
    Dim hSnapshot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    Dim nom(1 To 100)
    Dim num(1 To 100)
    Dim nr As Integer
    nr = 0
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapshot, uProcess)
    Do While r
        nr = nr + 1
        nom(nr) = uProcess.szexeFile
        num(nr) = uProcess.th32ProcessID
        r = ProcessNext(hSnapshot, uProcess)
    Loop
    For i = 1 To nr
    If InStr(UCase(nom(i)), UCase(processName)) <> 0 Then
        KillProcess (num(i))
        Exit For
    End If
    Next i
End Function

Public Function SendWindowsMessage(DestinationWindow As Long, sourcewindow As Long, Message As String)

    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 33500) As Byte
    Dim i As Long
    
    Call CopyMemory(buf(1), ByVal Message, Len(Message))
    cds.dwData = 1
    cds.cbData = Len(Message)
    cds.lpData = VarPtr(buf(1))
    i = SendMessage(DestinationWindow, WM_COPYDATA, sourcewindow, cds)

End Function

Public Function SendWindowsMessageTimeOut(DestinationWindow As Long, sourcewindow As Long, Message As String)

    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 33500) As Byte
    Dim i As Long, res As Long
    
    Call CopyMemory(buf(1), ByVal Message, Len(Message))
    cds.dwData = 1
    cds.cbData = Len(Message)
    cds.lpData = VarPtr(buf(1))
    i = SendMessageTimeout(DestinationWindow, WM_COPYDATA, sourcewindow, cds, &H2, 100, res)

End Function

