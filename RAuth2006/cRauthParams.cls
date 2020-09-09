VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRauthParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Const RauthFileName = "RauthInfo.xml"
Dim RemoteClient As String

Dim rauthDoc As New MSXML2.DOMDocument30
Dim rauthNode As IXMLDOMElement
Dim computername_  As IXMLDOMElement
Dim date_  As IXMLDOMElement
Dim ip_ As IXMLDOMElement
Dim uname_  As IXMLDOMElement
Dim ufullname_  As IXMLDOMElement
Dim chief_  As IXMLDOMElement
Dim manager_ As IXMLDOMElement
Dim password_ As IXMLDOMElement
Dim connected_ As IXMLDOMElement
Dim location As String

Dim iFile As String

Private Sub Class_Initialize()
    Set rauthNode = Nothing
    RemoteClient = ""
End Sub



Public Function Initialize(defaultblock As String, filename As String) As cRauthParams
    
    location = RauthDir
   
    If ChkXmlFileExistNew(location, filename) Then
        Dim rauth As cRauthParams
        Set rauthDoc = LoadXmlFile(location, filename)
        Set rauth = New cRauthParams
        rauth.Document = rauthDoc
        rauth.InitializeNodes filename
        rauth.File = filename
        Set Initialize = rauth
        
    Else
        Dim adoc As MSXML2.DOMDocument30
        Set adoc = XmlLoadFile(defaultblock, "RauthParams", "Πρόβλημα στη διαδικασία Έναρξης...")
        If adoc Is Nothing Then Exit Function
        Set rauthNode = GetXmlNode(adoc.documentElement, "//storage/rauth", "rauth", "storage", "Πρόβλημα στη διαδικασία Έναρξης...")
        If rauthNode Is Nothing Then Exit Function
        rauthDoc.loadXML rauthNode.xml
        
        Set rauth = New cRauthParams
        rauth.Document = rauthDoc
        rauth.File = filename
        rauth.InitializeNodes defaultblock
        Set Initialize = rauth
    End If
  
End Function
Public Sub InitializeNodes(filename As String)
On Error GoTo CorruptedData
    Set rauthNode = GetXmlNode(rauthDoc.documentElement, "//rauth", "rauth", "rauthDocument", "Πρόβλημα στη διαδικασία Έναρξης...")
    If rauthNode Is Nothing Then Exit Sub

    Set computername_ = GetXmlNode(rauthNode, "./computername", "ComputerName", "rauth")
    Set date_ = GetXmlNode(rauthNode, "./date", "Date", "rauth")
    Set uname_ = GetXmlNode(rauthNode, "./uname", "UserName", "rauth")
    Set ufullname_ = GetXmlNode(rauthNode, "./ufullname", "FullName", "rauth")
    Set ip_ = GetXmlNode(rauthNode, "./ip", "IP", "rauth")
    Set chief_ = GetXmlNode(rauthNode, "./chief", "CHIEF", "rauth")
    Set manager_ = GetXmlNode(rauthNode, "./manager", "MANAGER", "rauth")
    Set password_ = GetXmlNode(rauthNode, "./password", "PASSWORD", "rauth")
    Set connected_ = GetXmlNode(rauthNode, "./connected", "CONNECTED", "rauth")
    
    If date_.Text = "" Then date_.Text = CStr(Date)
    Exit Sub
CorruptedData:
    
     NBG_MsgBox "Πρόβλημα στο άνοιγμα του αρχείου " & filename & "... (Α6)  " & Error(), True, "ΛΑΘΟΣ"
    Exit Sub
End Sub
Public Function InitializeFromFile(location As String, filename As String) As cRauthParams
    Dim rauth As cRauthParams
    Dim adoc As MSXML2.DOMDocument30
    location = RauthDir
   
    If ChkXmlFileExistNew(location, filename) Then
        Set rauth = New cRauthParams
        Set adoc = LoadXmlFile(location, filename)
        rauth.Document = adoc
        rauth.File = filename
        rauth.InitializeNodes filename
    End If
    Set InitializeFromFile = rauth
End Function
'Public Function InitializeRemote(remotehost As String) As cRauthParams
'    RemoteClient = remotehost
'    Dim rauth As cRauthParams
'    Dim adoc As MSXML2.DOMDocument30
'    If ChkXmlFileExistRemote(RauthFileName, RemoteClient) Then
'        Set rauth = New cRauthParams
'        Set adoc = LoadXmlFileRemote(RauthFileName, RemoteClient)
'        rauth.Document = adoc
'        rauth.InitializeNodes
'    End If
'    Set InitializeRemote = rauth
'End Function

Public Property Get File() As String
   File = iFile
End Property
Public Property Let File(value As String)
    iFile = value
End Property

Public Property Get Document() As MSXML2.DOMDocument30
   Set Document = rauthDoc
End Property
Public Property Let Document(value As MSXML2.DOMDocument30)
    Set rauthDoc = value
End Property

Public Sub Store()
    location = RauthDir
    SaveXmlFileNew location, File, rauthDoc
End Sub

Public Property Get Connected() As Boolean
    Connected = IIf(connected_.Text = "1", True, False)
End Property
Public Property Let Connected(value As Boolean)
    connected_.Text = IIf(value = True, "1", "0") 'CStr(value)
End Property
Public Property Get Manager() As Boolean
    Manager = IIf(chief_.Text = "1", True, False)
End Property
Public Property Let Manager(value As Boolean)
    manager_.Text = IIf(value = True, "1", "0") 'CStr(value)
End Property

Public Property Get Chief() As Boolean
    Chief = IIf(chief_.Text = "1", True, False)
End Property
Public Property Let Chief(value As Boolean)
    chief_.Text = IIf(value = True, "1", "0") 'CStr(value)
End Property

Public Property Get Password() As String
    Password = password_.Text
End Property

Public Property Let Password(value As String)
    password_.Text = CStr(value)
End Property

Public Property Get IP() As String
    IP = ip_.Text
End Property

Public Property Let IP(value As String)
    ip_.Text = CStr(value)
End Property

Public Property Get WorkDate() As Date
    WorkDate = CDate(date_.Text)
End Property

Public Property Let WorkDate(value As Date)
    date_.Text = CStr(value)
End Property

Public Property Get ComputerName() As String
    ComputerName = computername_.Text
End Property

Public Property Let ComputerName(value As String)
    computername_.Text = value
End Property

Public Property Get UName() As String
    UName = uname_.Text
End Property

Public Property Let UName(value As String)
    uname_.Text = value
End Property

Public Property Get UFullName() As String
    UFullName = ufullname_.Text
End Property

Public Property Let UFullName(value As String)
    ufullname_.Text = value
End Property

