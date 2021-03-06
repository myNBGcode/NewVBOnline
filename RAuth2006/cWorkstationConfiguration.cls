VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWorkstationConfigurationMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim messagedoc As New MSXML2.DOMDocument30
Dim objNode As IXMLDOMElement
Dim messageNode As IXMLDOMElement
Dim computername_  As IXMLDOMElement
Dim message_  As IXMLDOMElement

Private Sub Class_Initialize()
    Set objNode = Nothing
End Sub

Public Function Initialize(filename As String) As cWorkstationConfigurationMessage
    Dim station As cWorkstationConfigurationMessage
   
        Dim adoc As MSXML2.DOMDocument30
        Set adoc = XmlLoadFile(filename, "GetComputerParams", "�������� ��� ���������� �������...")
        If adoc Is Nothing Then Exit Function
        
        Set objNode = GetXmlNode(adoc.documentElement, "//messages/GetComputerParams", "GetComputerParams", "messages", "�������� ��� ���������� �������...")
        If objNode Is Nothing Then Exit Function
        messagedoc.LoadXml objNode.XML
        
        Set station = New cWorkstationConfigurationMessage
        station.Document = messagedoc
        station.InitializeNodes
        Set Initialize = station
   


End Function

Public Sub InitializeNodes()

    Set messageNode = GetXmlNode(messagedoc.documentElement, "//MESSAGE", "MESSAGE", "MessageDocument", "�������� ��� ���������� �������...")
    If messageNode Is Nothing Then Exit Sub
    Set objNode = GetXmlNode(messagedoc.documentElement, "//OBJ", "OBJ", "MessageDocument", "�������� ��� ���������� �������...")
    If objNode Is Nothing Then Exit Sub
    Set computername_ = GetXmlNode(objNode, "./Computers/_COMPUTERNAME ", "ComputerName", "OBJ")
    Set message_ = GetXmlNode(messageNode, ".", "MESSAGE", "GetComputerParams")
    Exit Sub
End Sub

Public Property Get Document() As MSXML2.DOMDocument30
   Set Document = messagedoc
End Property
Public Property Let Document(value As MSXML2.DOMDocument30)
    Set messagedoc = value
End Property

Public Property Get ComputerName() As String
    ComputerName = computername_.Text
End Property

Public Property Let ComputerName(value As String)
    computername_.Text = value
End Property

Public Property Get Message() As String
    Message = message_.XML
End Property

'Public Property Let Message(value As String)
'    message_.XML = value
'End Property

Private Sub Class_Terminate()

Set objNode = Nothing
Set messageNode = Nothing
Set computername_ = Nothing
Set message_ = Nothing
Set messagedoc = Nothing

End Sub

