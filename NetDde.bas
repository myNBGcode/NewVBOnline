Attribute VB_Name = "NetDde"
Option Compare Text
Option Explicit

' API error codes
Enum NDDE_ERRORS
    NDDE_NO_ERROR = 0
    NDDE_ACCESS_DENIED = 1
    NDDE_BUF_TOO_SMALL = 2
    NDDE_ERROR_MORE_DATA = 3
    NDDE_INVALID_SERVER = 4
    NDDE_INVALID_SHARE = 5
    NDDE_INVALID_PARAMETER = 6
    NDDE_INVALID_LEVEL = 7
    NDDE_INVALID_PASSWORD = 8
    NDDE_INVALID_ITEMNAME = 9
    NDDE_INVALID_TOPIC = 10
    NDDE_INTERNAL_ERROR = 11
    NDDE_OUT_OF_MEMORY = 12
    NDDE_INVALID_APPNAME = 13
    NDDE_NOT_IMPLEMENTED = 14
    NDDE_SHARE_ALREADY_EXIST = 15
    NDDE_SHARE_NOT_EXIST = 16
    NDDE_INVALID_FILENAME = 17
    NDDE_NOT_RUNNING = 18
    NDDE_INVALID_WINDOW = 19
    NDDE_INVALID_SESSION = 20
    NDDE_INVALID_ITEM_LIST = 21
    NDDE_SHARE_DATA_CORRUPTED = 22
    NDDE_REGISTRY_ERROR = 23
    NDDE_CANT_ACCESS_SERVER = 24
    NDDE_INVALID_SPECIAL_COMMAND = 25
    NDDE_INVALID_SECURITY_DESC = 26
    NDDE_TRUST_SHARE_FAIL = 27
End Enum

' string size constants
Const MAX_NDDESHARENAME = 256
Const MAX_DOMAINNAME = 15
Const MAX_USERNAME = 15
Const MAX_APPNAME = 255
Const MAX_TOPICNAME = 255
Const MAX_ITEMNAME = 255

' connectFlags bits for ndde service affix
Const NDDEF_NOPASSWORDPROMPT = &H1
Const NDDEF_NOCACHELOOKUP = &H2
Const NDDEF_STRIP_NDDE = &H4


' NDDESHAREINFO - contains information about a NDDE share
Private Type NDDESHAREINFO
    lRevision As Long
    lpszShareName As String
    lShareType As Long
    lpszAppTopicList As String
    fSharedFlag As Long
    fService As Long
    fStartAppFlag As Long
    nCmdShow As Long
    qModifyId(0 To 2) As Long
    cNumItems As Long
    lpszItemList As String
End Type

'  Trusted share options
Public Const NDDE_TRUST_SHARE_START = &H80000000     ' Start App Allowed
Public Const NDDE_TRUST_SHARE_INIT = &H40000000      ' Init Conv Allowed
Public Const NDDE_TRUST_SHARE_DEL = &H20000000       ' Delete Trusted Share (on set)
Public Const NDDE_TRUST_CMD_SHOW = &H10000000        ' Use supplied cmd show
Public Const NDDE_CMD_SHOW_MASK = &HFFFF&            ' Command Show mask

'  Share Types
Public Const SHARE_TYPE_OLD = &H1                     ' Excel|sheet1.xls
Public Const SHARE_TYPE_NEW = &H2                     ' ExcelWorksheet|sheet1.xls
Public Const SHARE_TYPE_STATIC = &H4                  ' ClipSrv|SalesData

Declare Function NDdeShareAdd Lib "NDDEAPI.DLL" Alias "NDdeShareAddA" (ByVal lpszServer As String, ByVal nLevel As Long, pSD As Any, lpBuffer As Any, cBufSize As Long) As NDDE_ERRORS
 
Declare Function NDdeShareDel Lib "NDDEAPI.DLL" Alias "NDdeShareDelA" (ByVal lpszServer As String, ByVal lpszShareName As String, ByVal wReserved As Long) As NDDE_ERRORS
 
Declare Function NDdeGetShareSecurity Lib "NDDEAPI.DLL" Alias "NDdeGetShareSecurityA" (ByVal lpszServer As String, ByVal lpszShareName, si As Any, pSD As Any, ByVal cbSD As Long, ByRef lpcbsdRequired As Long) As NDDE_ERRORS
 
Declare Function NDdeSetShareSecurity Lib "NDDEAPI.DLL" Alias "NDdeSetShareSecurityA" (ByVal lpszServer As String, ByVal lpszShareName, si As Any, pSD As Any) As NDDE_ERRORS
 
Declare Function NDdeShareEnum Lib "NDDEAPI.DLL" Alias "NDdeShareEnumA" (ByVal lpszServer As String, ByVal nLevel As Long, lpBuffer As Any, ByVal cBufSize As Long, ByRef lpnEntriesRead As Long, ByRef lpcbTotalAvailable As Long) As NDDE_ERRORS

Declare Function NDdeShareGetInfo Lib "NDDEAPI.DLL" Alias "NDdeShareGetInfoA" (ByVal lpszServer As String, ByVal lpszShareName As String, ByVal nLevel As Long, lpBuffer As Any, ByVal cBufSize As Long, ByRef lpnTotalAvailable As Long, ByRef lpnItems As Integer) As NDDE_ERRORS
 
Declare Function NDdeShareSetInfo Lib "NDDEAPI.DLL" Alias "NDdeShareSetInfoA" (ByVal lpszServer As String, ByVal lpszShareName As String, ByVal nLevel As Long, lpBuffer As Any, ByVal cBufSize As Long, ByVal sParmNum As Integer) As NDDE_ERRORS

Declare Function NDdeSetTrustedShare Lib "NDDEAPI.DLL" Alias "NDdeSetTrustedShareA" (ByVal lpszServer As String, ByVal lpszShareName As String, ByVal dwTrustOptions As Long) As NDDE_ERRORS

Declare Function NDdeGetTrustedShare Lib "NDDEAPI.DLL" Alias "NDdeGetTrustedShareA" (ByVal lpszServer As String, ByVal lpszShareName As String, ByRef lpdwTrustOptions As Long, ByRef lpdwShareModId0 As Long, ByRef lpdwShareModId1 As Long) As NDDE_ERRORS

Declare Function NDdeTrustedShareEnum Lib "NDDEAPI.DLL" Alias "NDdeTrustedShareEnumA" (ByVal lpszServer As String, ByVal nLevel As Long, lpBuffer As Any, ByVal cBufSize As Long, ByRef lpnEntriesRead As Long, ByRef lpcbTotalAvailable As Long) As NDDE_ERRORS

Declare Function NDdeGetErrorString Lib "NDDEAPI.DLL" Alias "NDdeGetErrorStringA" (ByVal uErrorCode As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As NDDE_ERRORS

Declare Function NDdeIsValidShareName Lib "NDDEAPI.DLL" Alias "NDdeIsValidShareNameA" (ByVal ShareName As String) As Long

Declare Function NDdeIsValidAppTopicList Lib "NDDEAPI.DLL" Alias "NDdeIsValidAppTopicListA" (ByVal targetTopic As String) As Long

