VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkStationParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const WorkStationFileName = "WorkstationInfo.xml"
Dim RemoteClient As String

Dim workstationDoc As New MSXML2.DOMDocument30
Dim workstationNode As IXMLDOMElement
Dim computername_  As IXMLDOMElement
Dim date_  As IXMLDOMElement
Dim trnnum_  As IXMLDOMElement
Dim uname_  As IXMLDOMElement
Dim ufullname_  As IXMLDOMElement

Private Sub Class_Initialize()
    Set workstationNode = Nothing
     RemoteClient = ""
End Sub

'Public Sub Initialize(filename As String)
'    If ChkXmlFileExist(WorkStationFileName) Then
'        Set workstationDoc = LoadXmlFile(WorkStationFileName)
'    Else
'        Dim adoc As MSXML2.DOMDocument30
'        Set adoc = XmlLoadFile(filename, "WorkstationParams", "Πρόβλημα στη διαδικασία Έναρξης...")
'        If adoc Is Nothing Then Exit Sub
'
'        Set workstationNode = GetXmlNode(adoc.documentElement, "//storage/workstation", "workstation", "storage", "Πρόβλημα στη διαδικασία Έναρξης...")
'        If workstationNode Is Nothing Then Exit Sub
'        workstationDoc.LoadXml workstationNode.XML
'    End If
'
'   InitializeNodes
'
'End Sub

Public Function Initialize(filename As String) As cWorkStationParams
    Dim station As cWorkStationParams
    If ChkXmlFileExist(WorkStationFileName) Then
        Set workstationDoc = LoadXmlFile(NetworkHomeDir, WorkStationFileName)
        Set station = New cWorkStationParams
        station.Document = workstationDoc
        station.InitializeNodes
        Set Initialize = station
    Else
        Dim adoc As MSXML2.DOMDocument30
        Set adoc = XmlLoadFile(filename, "WorkstationParams", "Πρόβλημα στη διαδικασία Έναρξης...")
        If adoc Is Nothing Then Exit Function
        
        Set workstationNode = GetXmlNode(adoc.documentElement, "//storage/workstation", "workstation", "storage", "Πρόβλημα στη διαδικασία Έναρξης...")
        If workstationNode Is Nothing Then Exit Function
        workstationDoc.LoadXml workstationNode.XML
        
        Set station = New cWorkStationParams
        station.Document = workstationDoc
        station.InitializeNodes
        Set Initialize = station
    End If


End Function

Public Sub InitializeNodes()

On Error GoTo CorruptedData
    Set workstationNode = GetXmlNode(workstationDoc.documentElement, "//workstation", "workstation", "WorkstationDocument", "Πρόβλημα στη διαδικασία Έναρξης...")
    If workstationNode Is Nothing Then Exit Sub
    Set computername_ = GetXmlNode(workstationNode, "./computername", "ComputerName", "Workstation")
    Set date_ = GetXmlNode(workstationNode, "./date", "Date", "Workstation")
    Set trnnum_ = GetXmlNode(workstationNode, "./trnnum", "TransactionNumber", "Workstation")
    Set uname_ = GetXmlNode(workstationNode, "./uname", "UserName", "Workstation")
    Set ufullname_ = GetXmlNode(workstationNode, "./ufullname", "FullName", "Workstation")
    If date_.Text = "" Then date_.Text = CStr(Date)
    If trnnum_.Text = "" Then trnnum_.Text = "0"
    Exit Sub
CorruptedData:
    
    NBG_MsgBox "Πρόβλημα στο άνοιγμα του αρχείου " & WorkStationFileName & "... (Α6)  " & error(), True, "ΛΑΘΟΣ"
    Exit Sub
End Sub
Public Function InitializeRemote(remotehost As String) As cWorkStationParams
    RemoteClient = remotehost
    Dim station As cWorkStationParams
    Dim adoc As MSXML2.DOMDocument30
    If ChkXmlFileExistRemote(WorkStationFileName, RemoteClient) Then
        Set station = New cWorkStationParams
        Set adoc = LoadXmlFileRemote(WorkStationFileName, RemoteClient)
        station.Document = adoc
        station.InitializeNodes
    End If
    Set InitializeRemote = station
       
End Function

Public Property Get Document() As MSXML2.DOMDocument30
   Set Document = workstationDoc
End Property
Public Property Let Document(value As MSXML2.DOMDocument30)
    Set workstationDoc = value
End Property


Public Sub Store()
    SaveXmlFile WorkStationFileName, workstationDoc
End Sub

Public Property Get TrnNum() As Integer
    TrnNum = CInt(trnnum_.Text)
End Property

Public Property Let TrnNum(value As Integer)
    trnnum_.Text = CStr(value)
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

