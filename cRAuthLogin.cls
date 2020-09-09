VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRAuthLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim adoc As New MSXML2.DOMDocument30

Dim username_ As String
Dim password_ As String
Dim userfullname_ As String
Dim connectiontimestamp_ As Date
Dim disconnectiontimestamp_ As Date
Dim computername_ As String
Dim branchcode_ As String
Dim branchindex_ As String
Dim ip_ As String
Dim ischief_ As String
Dim ismanager_ As String

Dim aWeblink As cXMLWebLink
Dim method As cXMLWebMethod

Private Sub Class_Initialize()
    Set adoc = Nothing
End Sub
Public Sub Initialize(filename As String)
   On Error GoTo CorruptedData

    Set adoc = XmlLoadFile(filename)
    If adoc Is Nothing Then GoTo CorruptedData
        
    Exit Sub
CorruptedData:
    
    NBG_MsgBox "–Ò¸‚ÎÁÏ· ÛÙÔ ‹ÌÔÈ„Ï· ÙÔı ·Ò˜ÂﬂÔı " & filename & "... (¡6)  " & error(), True, "À¡»œ”"
    Exit Sub
End Sub

Public Function Find() As Boolean
On Error GoTo 0

    Find = False
    
    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim findNode As IXMLDOMNode
    Set findNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/FindRAuthLogin")
    If findNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No FindRAuthLogin node"

    messagedoc.LoadXML findNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No FindRAuthLogin/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserName").Text = username_
    
    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then
        Dim userNode As IXMLDOMElement
        Set userNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//ODATA/RAuthLogins")
        If Not userNode Is Nothing Then
            username_ = GetXmlNodeIfPresent(userNode, "./_UserName").Text
            password_ = GetXmlNodeIfPresent(userNode, "./_Password").Text
            
            Dim coder As New Base64Class
            password_ = coder.DecodeToString(password_)
            password_ = SimpleDecrypt(password_, UCase(username_))
            
            Dim adate As String, bdate As String
            adate = GetXmlNodeIfPresent(userNode, "./_ConnectionTimestamp").Text
            bdate = GetXmlNodeIfPresent(userNode, "./_DisconnectionTimestamp").Text
            If adate <> "" Then
                connectiontimestamp_ = DateSerial(Left(adate, 4), Mid(adate, 6, 2), Mid(adate, 9, 2))
                connectiontimestamp_ = DateAdd("h", Mid(adate, 12, 2), connectiontimestamp_)
                connectiontimestamp_ = DateAdd("n", Mid(adate, 15, 2), connectiontimestamp_)
                connectiontimestamp_ = DateAdd("s", Mid(adate, 18, 2), connectiontimestamp_)
            End If
            If bdate <> "" Then
                disconnectiontimestamp_ = DateSerial(Left(bdate, 4), Mid(bdate, 6, 2), Mid(bdate, 9, 2))
                disconnectiontimestamp_ = DateAdd("h", Mid(bdate, 12, 2), disconnectiontimestamp_)
                disconnectiontimestamp_ = DateAdd("n", Mid(bdate, 15, 2), disconnectiontimestamp_)
                disconnectiontimestamp_ = DateAdd("s", Mid(bdate, 18, 2), disconnectiontimestamp_)
            End If
            
            userfullname_ = GetXmlNodeIfPresent(userNode, "./_UserFullName").Text
            computername_ = GetXmlNodeIfPresent(userNode, "./_ComputerName").Text
            branchcode_ = GetXmlNodeIfPresent(userNode, "./_BranchCode").Text
            branchindex_ = GetXmlNodeIfPresent(userNode, "./_BranchIndex").Text
            ip_ = GetXmlNodeIfPresent(userNode, "./_IP").Text
            ischief_ = GetXmlNodeIfPresent(userNode, "./_IsChief").Text
            ismanager_ = GetXmlNodeIfPresent(userNode, "./_IsManager").Text
            
            Find = True
        End If
    End If
    
    Set findNode = Nothing
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function

Public Function Update() As Boolean
On Error GoTo 0

    Update = False

    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim updateNode As IXMLDOMNode
    Set updateNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/UpdateRAuthLogin")
    If updateNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No UpdateRAuthLogin node"

    password_ = SimpleEncrypt(password_, UCase(username_))
    Dim coder As New Base64Class
    password_ = coder.EncodeString(password_)

    messagedoc.LoadXML updateNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No UpdateRAuthLogin/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserName").Text = username_
    message_.selectSingleNode("//OBJ/RAuthLogins/_Password").Text = password_
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserFullName").Text = userfullname_
    message_.selectSingleNode("//OBJ/RAuthLogins/_ComputerName").Text = computername_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchCode").Text = branchcode_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchIndex").Text = branchindex_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IP").Text = ip_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsChief").Text = ischief_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsManager").Text = ismanager_

    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then Update = True
    
    Set updateNode = Nothing
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function

Public Function ChangePassword() As Boolean
On Error GoTo 0

    ChangePassword = False

    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim changeNode As IXMLDOMNode
    Set changeNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/ChangePassRAuthLogin")
    If changeNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No ChangePassRAuthLogin node"

    password_ = SimpleEncrypt(password_, UCase(username_))
    Dim coder As New Base64Class
    password_ = coder.EncodeString(password_)

    messagedoc.LoadXML changeNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No ChangePassRAuthLogin/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserName").Text = username_
    message_.selectSingleNode("//OBJ/RAuthLogins/_Password").Text = password_
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserFullName").Text = userfullname_
    message_.selectSingleNode("//OBJ/RAuthLogins/_ComputerName").Text = computername_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchCode").Text = branchcode_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchIndex").Text = branchindex_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IP").Text = ip_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsChief").Text = ischief_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsManager").Text = ismanager_

    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then ChangePassword = True
    
    Set changeNode = Nothing
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function

Public Function Insert() As Boolean
On Error GoTo 0

    Insert = False

    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim insertNode As IXMLDOMNode
    Set insertNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/InsertRAuthLogin")
    If insertNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No InsertRAuthLogin node"

    password_ = SimpleEncrypt(password_, UCase(username_))
    Dim coder As New Base64Class
    password_ = coder.EncodeString(password_)
    
    messagedoc.LoadXML insertNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No InsertRAuthLogin/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserName").Text = username_
    message_.selectSingleNode("//OBJ/RAuthLogins/_Password").Text = password_
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserFullName").Text = userfullname_
    message_.selectSingleNode("//OBJ/RAuthLogins/_ComputerName").Text = computername_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchCode").Text = branchcode_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchIndex").Text = branchindex_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IP").Text = ip_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsChief").Text = ischief_
    message_.selectSingleNode("//OBJ/RAuthLogins/_IsManager").Text = ismanager_
    
    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then Insert = True
    
    Set insertNode = Nothing
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function
Public Function Disconnect() As Boolean
On Error GoTo 0

    Disconnect = False

    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim updateNode As IXMLDOMNode
    Set updateNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/DisconnectRAuthLogin")
    If updateNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No DisconnectRAuthLogin node"

    messagedoc.LoadXML updateNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No DisconnectRAuthLogin/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_UserName").Text = username_
    
    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then Disconnect = True
    
    Set updateNode = Nothing
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function

Public Function FindConnectedUsers() As Collection
On Error GoTo 0

    Set FindConnectedUsers = Nothing
    
    Dim messagedoc As New MSXML2.DOMDocument30
    Dim message_ As IXMLDOMElement

    Dim findNode As IXMLDOMNode
    Set findNode = GetXmlNodeIfPresent(adoc.documentElement, "//messages/GetConnectedRAuthLogins")
    If findNode Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No GetConnectedRAuthLogins node"

    messagedoc.LoadXML findNode.XML
    Set message_ = GetXmlNodeIfPresent(messagedoc.documentElement, "//MESSAGE")
    If message_ Is Nothing Then Err.Raise 9999, "cRAuthLogin", "No GetConnectedRAuthLogins/MESSAGE node"
    
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchCode").Text = branchcode_
    message_.selectSingleNode("//OBJ/RAuthLogins/_BranchIndex").Text = branchindex_
    message_.selectSingleNode("//OBJ/RAuthLogins/_ConnectionTimestamp").Text = format(connectiontimestamp_, "yyyy-mm-dd")
    
    Dim ares As String
    ares = method.LoadXmlNoTrnUpdate(message_.XML)
    Dim tempdoc As New MSXML2.DOMDocument30
    tempdoc.LoadXML ares
    
    Dim returnNode As IXMLDOMNode
    Set returnNode = GetXmlNodeIfPresent(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text <> "0" Then
        Set FindConnectedUsers = New Collection
        Dim loginNode As IXMLDOMNode
        For Each loginNode In tempdoc.documentElement.SelectNodes("//RAuthLogins")
            Dim aLogin As New cRAuthLogin
            
            aLogin.UserName = GetXmlNodeIfPresent(loginNode, "./_UserName").Text
            aLogin.UserFullName = GetXmlNodeIfPresent(loginNode, "./_UserFullName").Text
            aLogin.ComputerName = GetXmlNodeIfPresent(loginNode, "./_ComputerName").Text
            aLogin.BranchCode = GetXmlNodeIfPresent(loginNode, "./_BranchCode").Text
            aLogin.¬ranch…ndex = GetXmlNodeIfPresent(loginNode, "./_BranchIndex").Text
            aLogin.IP = GetXmlNodeIfPresent(loginNode, "./_IP").Text
            
            Dim adate As String, bdate As String
            adate = GetXmlNodeIfPresent(loginNode, "./_ConnectionTimestamp").Text
            bdate = GetXmlNodeIfPresent(loginNode, "./_DisconnectionTimestamp").Text
            If adate <> "" Then
                aLogin.ConnectionTimestamp = DateSerial(Left(adate, 4), Mid(adate, 6, 2), Mid(adate, 9, 2))
                aLogin.ConnectionTimestamp = DateAdd("h", Mid(adate, 12, 2), aLogin.ConnectionTimestamp)
                aLogin.ConnectionTimestamp = DateAdd("n", Mid(adate, 15, 2), aLogin.ConnectionTimestamp)
                aLogin.ConnectionTimestamp = DateAdd("s", Mid(adate, 18, 2), aLogin.ConnectionTimestamp)
            End If
            If bdate <> "" Then
                aLogin.DisConnectionTimestamp = DateSerial(Left(bdate, 4), Mid(bdate, 6, 2), Mid(bdate, 9, 2))
                aLogin.DisConnectionTimestamp = DateAdd("h", Mid(bdate, 12, 2), aLogin.DisConnectionTimestamp)
                aLogin.DisConnectionTimestamp = DateAdd("n", Mid(bdate, 15, 2), aLogin.DisConnectionTimestamp)
                aLogin.DisConnectionTimestamp = DateAdd("s", Mid(bdate, 18, 2), aLogin.DisConnectionTimestamp)
            End If
            
            aLogin.IsChief = GetXmlNodeIfPresent(loginNode, "./_IsChief").Text
            aLogin.…sManager = GetXmlNodeIfPresent(loginNode, "./_IsManager").Text

            FindConnectedUsers.add aLogin
            Set aLogin = Nothing
        Next
    Else
        Set FindConnectedUsers = Nothing
    End If
    
    Set message_ = Nothing
    Set messagedoc = Nothing
End Function

Public Property Let WebLink(value As cXMLWebLink)
    Set aWeblink = value
End Property
Public Property Let WebMethod(value As cXMLWebMethod)
    Set method = value
End Property

Public Property Get UserName() As String
    UserName = username_
End Property

Public Property Let UserName(value As String)
    username_ = value
End Property
Public Property Get Password() As String
    Password = password_
End Property

Public Property Let Password(value As String)
    password_ = value
End Property

Public Property Get UserFullName() As String
    UserFullName = userfullname_
End Property

Public Property Let UserFullName(value As String)
    userfullname_ = value
End Property

Public Property Get ConnectionTimestamp() As Date
    ConnectionTimestamp = connectiontimestamp_
End Property

Public Property Let ConnectionTimestamp(value As Date)
    connectiontimestamp_ = value
End Property

Public Property Get DisConnectionTimestamp() As Date
    DisConnectionTimestamp = disconnectiontimestamp_
End Property

Public Property Let DisConnectionTimestamp(value As Date)
    disconnectiontimestamp_ = value
End Property

Public Property Get ComputerName() As String
    ComputerName = computername_
End Property

Public Property Let ComputerName(value As String)
    computername_ = value
End Property

Public Property Get BranchCode() As String
    BranchCode = branchcode_
End Property

Public Property Let BranchCode(value As String)
    branchcode_ = value
End Property

Public Property Get ¬ranch…ndex() As String
    ¬ranch…ndex = branchindex_
End Property

Public Property Let ¬ranch…ndex(value As String)
    branchindex_ = value
End Property

Public Property Get IP() As String
    IP = ip_
End Property

Public Property Let IP(value As String)
    ip_ = value
End Property

Public Property Get IsChief() As String
    IsChief = ischief_
End Property

Public Property Let IsChief(value As String)
    ischief_ = value
End Property
Public Property Get …sManager() As String
    …sManager = ismanager_
End Property

Public Property Let …sManager(value As String)
    ismanager_ = value
End Property

Private Sub Class_Terminate()
    Set adoc = Nothing
End Sub
