VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExecRuleWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WebServiceUrl As String

Public OpClass As String
Public OpCode As String
Public OpDescription As String

Dim soapfile As String

Private Sub Class_Initialize()
    soapfile = ReadDir & "\WebServices\ExecRuleB64CompressedWithLog_SOAPMessage.xml"
End Sub

Private Function ExecMethod(soapDoc As MSXML2.DOMDocument30) As MSXML2.DOMDocument30
    Dim sender As New MSXML2.ServerXMLHTTP30
    Dim soapResponse As New MSXML2.DOMDocument30
    Set soapResponse = CreateObject("Msxml2.DOMDocument.4.0")
    Set sender = New MSXML2.ServerXMLHTTP
    If UseActiveDirectory Then
        sender.open "POST", WebServiceUrl, False
    Else
        sender.open "POST", WebServiceUrl, False, "CENTRAL\webuser", "!*WebUser*!"
    End If
    sender.setRequestHeader "Content-Type", "text/xml"
    sender.setRequestHeader "SOAPAction", "http://tempuri.org/ExecRuleB64CompressedWithLog"
    
    'soapDoc.save "C:\request.xml"
    SaveXmlFile "Request.xml", soapDoc
    
    On Error GoTo Send_Error
    sender.setTimeouts 180000, 180000, 180000, 180000
    eJournalWrite "cExecRuleWrapper.ExecMethod Send : " & format(Time, "HH:MM:SS")
    sender.SEND soapDoc.XML
    eJournalWrite "cExecRuleWrapper.ExecMethod Receive : " & format(Time, "HH:MM:SS")
    soapResponse.LoadXML sender.responseXML.XML
    
    If sender.responseXML.XML = "" And sender.responseText <> "" Then
        'Open "c:\responsetext.txt" For Output As #1
        Open NetworkHomeDir & "\responsetext.txt" For Output As #1
        Print #1, sender.responseText
        Close #1
        LogMsgbox "Καταγραφή προβλήματος στο ResponseText.txt" & vbCrLf & _
            sender.responseText, vbOKOnly, "Πρόβλημα Επικοινωνίας"
    End If
    
'    MsgBox "step2"
'    MsgBox "ResponseXml:" & sender.responseXML.XML
'    MsgBox "ResponseText:" & sender.responseText
    Set ExecMethod = soapResponse
    
'    MsgBox "step3"
    Dim aErr As IXMLDOMElement
    Set aErr = soapResponse.selectSingleNode("//*[local-name()='Fault']")
    If Not (aErr Is Nothing) Then
        If Not aErr.childNodes Is Nothing Then
            If aErr.childNodes.length > 0 Then
                Dim aMsg As String
                Dim aErrLine As IXMLDOMElement
                For Each aErrLine In aErr.childNodes
                    aMsg = aMsg & "( " & aErrLine.nodename & " ):" & aErrLine.Text & vbCrLf
                Next
                'soapResponse.save "C:\SoapResponseError.xml"
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Err.Raise 999, "cExecRuleWrapper", aMsg
                Exit Function
            End If
        End If
    End If
'    MsgBox "step4"
    'soapResponse.save "C:\response.xml"
    SaveXmlFile "Response.xml", soapResponse
    Exit Function
Send_Error:
    Err.Raise 999, "cExecRuleWrapper.ExecMethod", "Πρόβλημα Κλήσης Web Service: " & Err.number & ", " & Err.description
End Function

Private Function SendData(slength As Long, Data As String, rlength As Long) As String
    Dim adoc As New MSXML2.DOMDocument30
    Dim ExecRuleResult As MSXML2.IXMLDOMElement
    
    adoc.Load soapfile
    Dim SendLength As MSXML2.IXMLDOMNode
    Dim sendbuffer As MSXML2.IXMLDOMNode
    Dim receivelength As MSXML2.IXMLDOMNode
    Dim branchcode_ As MSXML2.IXMLDOMNode
    Dim computername_ As MSXML2.IXMLDOMNode
    Dim username_ As MSXML2.IXMLDOMNode
    Dim opclass_ As MSXML2.IXMLDOMNode
    Dim opcode_ As MSXML2.IXMLDOMNode
    Dim opdescription_ As MSXML2.IXMLDOMNode
    Dim systemid_ As MSXML2.IXMLDOMNode
    
    Set SendLength = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/SendLength")
    Set sendbuffer = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/SendData")
    Set receivelength = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/ReceiveLength")
    
    Set branchcode_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/BranchCode")
    Set computername_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/ComputerName")
    Set username_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/UserName")
    Set opclass_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/OpClass")
    Set opcode_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/OpCode")
    Set opdescription_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/OpDescription")
    Set systemid_ = adoc.selectSingleNode("//ExecRuleB64CompressedWithLog/SystemID")
    SendLength.Text = slength
    sendbuffer.Text = Data
    receivelength.Text = rlength
    
    branchcode_.Text = cBRANCH
    computername_.Text = MachineName
    username_.Text = cUserName
    opclass_.Text = OpClass
    opcode_.Text = OpCode
    opdescription_.Text = OpDescription
    systemid_.Text = "SHINE"
    
    Set adoc = ExecMethod(adoc)
    Set ExecRuleResult = adoc.selectSingleNode("//*[local-name()='ExecRuleB64CompressedWithLogResponse']/*[local-name()='ExecRuleB64CompressedWithLogResult']")
    SendData = ExecRuleResult.Text
End Function

Public Function SendString(SendLength As Long, Data As String, receivelength As Long) As String
    Dim coder As New Base64Class
    Dim astr As String
    Dim Result As Long
    Dim bytes() As Byte
    'bytes = StrConv(Data, vbFromUnicode)
    
    'Result = GenWorkForm.ZControl.CompressData(bytes)
    Result = GenWorkForm.ZControl.CompressString(Data)
    
    astr = coder.EncodeString(Data)
    astr = SendData(SendLength, astr, receivelength)
    
    
    If ExecRuleB64ReceiveFile <> "" Then
        Open ExecRuleB64ReceiveFile For Output As #3
        Print #3, astr
        Close #3
    End If
    
    SendString = coder.DecodeToString(astr)
    bytes = StrConv(SendString, vbFromUnicode)
    Result = GenWorkForm.ZControl.DecompressData(bytes(), receivelength)
    SendString = StrConv(bytes, vbUnicode)
    Result = 0
End Function

'Public Function SendByteArray(data() As Byte, receivelength As Long) As Byte()
'    Dim coder As New Base64Class
'    Dim astr As String
'    astr = coder.EncodeByteArray(data)
'    astr = SendData(UBound(data) - LBound(data) + 1, astr, receivelength)
'    SendByteArray = coder.DecodeToByteArray(astr)
'End Function
'
