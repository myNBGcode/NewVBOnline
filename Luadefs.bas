Attribute VB_Name = "LUADEFS"

Declare Function VB4SLICONNECT _
        Lib "VB4SLI.DLL" _
        (pLUName As String, _
         pAppId As String, _
         ByVal pConvertIt As Long, _
         ByVal pTimeOut As Long, _
         pRet1 As Long, pRet2 As Long, pRet3 As Long, _
         ByVal pDebug As Long) As String

Declare Function VB4SLIDISCONNECT _
        Lib "VB4SLI.DLL" _
        (ByVal pTimeOut As Long, _
         pRet1 As Long, pRet2 As Long, pRet3 As Long, _
         ByVal pDebug As Long) As String

Declare Function VB4SLISEND _
        Lib "VB4SLI.DLL" _
        (pData As String, _
         ByVal pConvertIt As Long, _
         ByVal pTimeOut As Long, _
         pLen As Long, _
         pMsgType As Long, _
         pRet1 As Long, pRet2 As Long, pRet3 As Long, _
         ByVal pDebug As Long) As String

Declare Function VB4SLIRECEIVE _
        Lib "VB4SLI.DLL" _
        (ByVal pConvertIt As Long, _
         ByVal pTimeOut As Long, _
         pLen As Long, _
         pMsgType As Long, _
         pRet1 As Long, pRet2 As Long, pRet3 As Long, _
         ByVal pDebug As Long) As String

'Public Const BETB = 1
'Public Const SEND = 2
'Public Const RECV = 3
'Public pLUADirection As Long

'Global Connected As Integer
'Global Batch As Integer
'Global pBuff As String
