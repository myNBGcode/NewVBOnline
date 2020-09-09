VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExecTRNWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WebServiceUrl As String

Public OpClass As String
Public OpCode As String
Public OpDescription As String
Public AuthUser As String
Public MessageID As String

Private times As Integer
Dim soapfile As String

Private Sub Class_Initialize()
    'soapfile = ReadDir & "\WebServices\ExecTRNB64CompressedWithLogV2_SOAPMessage.xml"
    'soapfile = ReadDir & "\WebServices\ExecTB64CLV2_SOAPMessage.xml"
    soapfile = ReadDir & "\WebServices\ExecTB64CLV3_SOAPMessage.xml"
    times = 0
End Sub

Private Function ExecMethod(soapDoc As MSXML2.DOMDocument30, Optional method) As MSXML2.DOMDocument30
   
    Dim Line As String
    Dim sender As New MSXML2.ServerXMLHTTP30
    Dim soapResponse As New MSXML2.DOMDocument30
    Line = "1000"
    Set soapResponse = CreateObject("Msxml2.DOMDocument.6.0")
    Set sender = New MSXML2.ServerXMLHTTP
    Line = "1002"
    If UseActiveDirectory Then
        sender.open "POST", WebServiceUrl, False
    Else
        sender.open "POST", WebServiceUrl, False, "CENTRAL\webuser", "!*WebUser*!"
    End If
    sender.setRequestHeader "Content-Type", "text/xml"
    If IsMissing(method) Then
        'sender.setRequestHeader "SOAPAction", "http://tempuri.org/ExecTB64CLV2"
        sender.setRequestHeader "SOAPAction", "http://tempuri.org/ExecTB64CLV3"
    Else
        sender.setRequestHeader "SOAPAction", "http://tempuri.org/" & method
    End If
    SaveXmlFile "Request.xml", soapDoc
    Line = "1005"
    On Error GoTo Send_Error
    sender.setTimeouts 180000, 180000, 180000, 180000
'    If (times >= 3) Then
'        sender.setTimeouts 180000, 180000, 180000, 180000
'    Else
'        sender.setTimeouts 5000, 5000, 5000, 5000
'    End If
    Line = "1010"
    'eJournalWrite "cExecTRNWrapper.ExecMethod Send : " & format(Time, "HH:MM:SS")
    sender.SEND soapDoc.xml
    'eJournalWrite "cExecTRNWrapper.ExecMethod Receive : " & format(Time, "HH:MM:SS")
    Line = "1015"
    soapResponse.LoadXML sender.responseXML.xml
    
    Line = "1020"
    If sender.responseXML.xml = "" And sender.responseText <> "" Then
        Open NetworkHomeDir & "\responsetext.txt" For Output As #1
        Print #1, sender.responseText
        Close #1
        LogMsgbox "Καταγραφή προβλήματος στο ResponseText.txt" & vbCrLf & _
            sender.responseText, vbOKOnly, "Πρόβλημα Επικοινωνίας"
'        MsgBox "Καταγραφή προβλήματος στο ResponseText.txt" & vbCrLf & _
'            sender.responseText, vbOKOnly, "Πρόβλημα Επικοινωνίας"
    End If
    Line = "1025"
'    MsgBox "step2"
'    MsgBox "ResponseXml:" & sender.responseXML.XML
'    MsgBox "ResponseText:" & sender.responseText
    Set ExecMethod = soapResponse
    Line = "1030"
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
                Line = "1035"
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Err.Raise 999, "cExecTRNWrapper", aMsg
                Exit Function
            End If
        End If
    End If
'    MsgBox "step4"
    SaveXmlFile "Response.xml", soapResponse
    Exit Function
Send_Error:
    Dim error
    Set error = Err
    eJournalWrite "Πρόβλημα Κλήσης Web Service: Line=" & Line & " " & " Λάθος:" & error.number & " " & error.description & " " & error.source
     
    times = times + 1
    ShowStatusMessage "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ.ΠΑΡΑΚΑΛΩ ΠΕΡΙΜΕΝΕΤΕ..."
    If (times < 4) Then
        LogMsgbox "Πρόβλημα επικοινωνίας. Πατήστε ΟΚ για ανάκτηση στοιχείων.", vbOKOnly, "ΠΡΟΣΟΧΗ!"
        Set ExecMethod = AskData(soapDoc)
    Else
        Err.Raise 999, "cExecTRNWrapper.ExecMethod", "Πρόβλημα Κλήσης Web Service: " & Err.number & ", " & Err.description
    End If
End Function


Private Function AskData(soapDoc As MSXML2.DOMDocument30) As MSXML2.DOMDocument30
    Dim ExecRuleResult As MSXML2.IXMLDOMElement
    Dim adoc As New MSXML2.DOMDocument30
    SaveXmlFile "soap.xml", soapDoc
    Dim oldnode, tempnode, newnode   As IXMLDOMNode
    Set oldnode = soapDoc.selectSingleNode("//ExecTB64CLV3")
    If Not (oldnode Is Nothing) Then
        Set newnode = soapDoc.createNode(NODE_ELEMENT, "AskTB64CLV2", "http://tempuri.org/")
        
        Dim nodelist As IXMLDOMNodeList
        Set nodelist = oldnode.childNodes
        For Each tempnode In nodelist
            newnode.appendChild tempnode
        Next
        
        soapDoc.documentElement.selectSingleNode("//soap:Body").replaceChild newnode, oldnode
    End If
   
    SaveXmlFile "ask.xml", soapDoc

    Dim StartTime
    StartTime = Time
    Dim StartTickCount
    StartTickCount = GetTickCount
    
    eJournalWrite "ASK: " & format(Time, "HH:MM:SS")
    Set adoc = ExecMethod(soapDoc, "AskTB64CLV2")
    
    Dim oldresultnode, newresultnode As IXMLDOMNode
    
    Set oldnode = adoc.selectSingleNode("//*[local-name()='AskTB64CLV2Response']")
    If (oldnode Is Nothing) Then
        Set AskData = adoc
        Exit Function
    End If
    Set oldresultnode = adoc.selectSingleNode("//*[local-name()='AskTB64CLV2Result']")
    
    Set newnode = adoc.createNode(NODE_ELEMENT, "ExecTB64CLV3Response", "http://tempuri.org/")
    Set newresultnode = adoc.createNode(NODE_ELEMENT, "ExecTB64CLV3Result", "http://tempuri.org/")
    newnode.appendChild newresultnode
    
    Set nodelist = oldresultnode.childNodes
    For Each tempnode In nodelist
       newresultnode.appendChild tempnode
    Next
    adoc.documentElement.selectSingleNode("*[local-name()='Body']").replaceChild newnode, oldnode
  
    Set AskData = adoc
End Function

Private Function SendData(slength As Long, Data As String) As cExecTrnResult
    Dim adoc As New MSXML2.DOMDocument30
    Dim ExecRuleResult As MSXML2.IXMLDOMElement
    
    adoc.Load soapfile
    Dim SendLength As MSXML2.IXMLDOMNode
    Dim sendbuffer As MSXML2.IXMLDOMNode
    Dim branchcode_ As MSXML2.IXMLDOMNode
    Dim computername_ As MSXML2.IXMLDOMNode
    Dim username_ As MSXML2.IXMLDOMNode
    Dim authuser_ As MSXML2.IXMLDOMNode
    Dim opclass_ As MSXML2.IXMLDOMNode
    Dim opcode_ As MSXML2.IXMLDOMNode
    Dim opdescription_ As MSXML2.IXMLDOMNode
    Dim systemid_ As MSXML2.IXMLDOMNode
    Dim lastmessage_ As MSXML2.IXMLDOMNode
    Dim lastmessageexecutiontime_ As MSXML2.IXMLDOMNode
    Dim messageid_ As MSXML2.IXMLDOMNode
    Dim trnumber_ As MSXML2.IXMLDOMNode
    Dim uicode_ As MSXML2.IXMLDOMNode
    Dim comment_ As MSXML2.IXMLDOMNode
        
    Set SendLength = adoc.selectSingleNode("//ExecTB64CLV3/SendLength")
    Set sendbuffer = adoc.selectSingleNode("//ExecTB64CLV3/SendData")
    
    Set branchcode_ = adoc.selectSingleNode("//ExecTB64CLV3/BranchCode")
    Set computername_ = adoc.selectSingleNode("//ExecTB64CLV3/ComputerName")
    Set username_ = adoc.selectSingleNode("//ExecTB64CLV3/UserName")
    Set authuser_ = adoc.selectSingleNode("//ExecTB64CLV3/AuthUser")
    Set opclass_ = adoc.selectSingleNode("//ExecTB64CLV3/OpClass")
    Set opcode_ = adoc.selectSingleNode("//ExecTB64CLV3/OpCode")
    Set opdescription_ = adoc.selectSingleNode("//ExecTB64CLV3/OpDescription")
    Set systemid_ = adoc.selectSingleNode("//ExecTB64CLV3/SystemID")
    Set lastmessage_ = adoc.selectSingleNode("//ExecTB64CLV3/LastMessage")
    Set lastmessageexecutiontime_ = adoc.selectSingleNode("//ExecTB64CLV3/LastMessageExecutionTime")
    Set messageid_ = adoc.selectSingleNode("//ExecTB64CLV3/MessageID")
    Set trnumber_ = adoc.selectSingleNode("//ExecTB64CLV3/TrNumber")
    Set uicode_ = adoc.selectSingleNode("//ExecTB64CLV3/UICode")
    Set comment_ = adoc.selectSingleNode("//ExecTB64CLV3/Comment")
    
    SendLength.Text = slength
    sendbuffer.Text = Data
    
    branchcode_.Text = cBRANCH
    computername_.Text = MachineName
    username_.Text = cUserName
    authuser_.Text = AuthUser
    opclass_.Text = OpClass
    opcode_.Text = OpCode
    opdescription_.Text = OpDescription
    systemid_.Text = "SHINE"
    If (ExecutionResults.Count > 0) Then
        lastmessage_.Text = ExecutionResults(ExecutionResults.Count).MessageID
        lastmessageexecutiontime_.Text = ExecutionResults(ExecutionResults.Count).ExecutionTime
    Else
        lastmessage_.Text = ""
        lastmessageexecutiontime_.Text = "0"
    End If
    
    'Dim aguid As String
    'aguid = GetGuid
    
    messageid_.Text = MessageID
    trnumber_.Text = cTRNNum
    uicode_.Text = cTRNCode
    
    eJournalWrite "Ταυτότητα Συν/γής: " & MessageID
    
    Dim StartTime
    StartTime = Time
    Dim StartTickCount
    StartTickCount = GetTickCount
    
    Set adoc = ExecMethod(adoc)

    Dim EndTime
    Dim EndTickCount
    EndTime = Time
    EndTickCount = GetTickCount
    
    Set ExecRuleResult = adoc.selectSingleNode("//*[local-name()='ExecTB64CLV3Response']/*[local-name()='ExecTB64CLV3Result']")
    If ExecRuleResult Is Nothing Then
        Set SendData = Nothing
    Else
        Set SendData = New cExecTrnResult
        SendData.TranslateNode ExecRuleResult
        SendData.ExecutionTime = EndTickCount - StartTickCount
        
        ExecutionResults.add SendData
    End If
End Function

Public Function SendString(SendLength As Long, Data As String) As String
    Dim coder As New Base64Class
    Dim astr As String
    Dim Result As Long
    Dim bytes() As Byte
    bytes = StrConv(Data, vbFromUnicode)
    
    Result = GenWorkForm.ZControl.CompressString(Data)
    
    bytes = StrConv(Data, vbFromUnicode)
    
    astr = coder.EncodeString(Data)
    bytes = StrConv(astr, vbFromUnicode)
    
    Call PushMessage(astr, "Customer Marketing Listener")
    
    Dim SendResult As cExecTrnResult
    Set SendResult = SendData(SendLength, astr)
    
    
    If ExecRuleB64ReceiveFile <> "" Then
        Open ExecRuleB64ReceiveFile For Output As #3
        Print #3, astr
        Close #3
    End If
    
    If astr <> "" Then
        Dim length As String
        length = SendResult.ResultLength 'Left(astr, 7)
        If (length = 0 And SendResult.ExceptionData <> "") Then
            LogMsgbox SendResult.ExceptionData, vbOKOnly, "ΠΡΟΣΟΧΗ! cExecTRNWrapper.SendString"
             SendString = ""
        Else
            astr = SendResult.ResultData 'Right(astr, Len(astr) - 7)
            SendString = coder.DecodeToString(astr)
            bytes = StrConv(SendString, vbFromUnicode)
            Result = GenWorkForm.ZControl.DecompressData(bytes(), CInt(length))
            SendString = StrConv(bytes, vbUnicode)
        End If
        
    Else
        SendString = ""
    End If
    Result = 0
    
    
End Function

Public Sub PushMessage(base64str As String, windowHeaderDest As String)
    
    Dim xmldoc As New MSXML2.DOMDocument60
    Dim s As IXMLDOMElement
    Dim n As IXMLDOMNode
    Dim a As IXMLDOMAttribute
    Set s = xmldoc.createElement("Message")
    
    Set a = s.ownerDocument.createAttribute("opclass")
    a.value = OpClass
    s.Attributes.setNamedItem a
    
    Set a = s.ownerDocument.createAttribute("opcode")
    a.value = OpCode
    s.Attributes.setNamedItem a
    
    Set a = s.ownerDocument.createAttribute("opdescription")
    a.value = OpDescription
    s.Attributes.setNamedItem a
    
    Set a = s.ownerDocument.createAttribute("uicode")
    a.value = cTRNCode
    s.Attributes.setNamedItem a
    
    Set a = s.ownerDocument.createAttribute("journalid")
    a.value = MessageID
    s.Attributes.setNamedItem a
    
    Set xmldoc.documentElement = s

    Set n = xmldoc.createNode(NODE_ELEMENT, "Data", "")
    n.Text = base64str
    s.appendChild n

    If Not Screen.activeform Is Nothing Then
        Dim destwindow As Long, sourcewindow As Long
        destwindow = FindWindow(vbNullString, windowHeaderDest)
        sourcewindow = Screen.activeform.hWnd
        If destwindow <> 0 And Left(Right(WorkEnvironment_, 8), 4) = "PROD" Then
'            Call SendWindowsMessage(destwindow, sourcewindow, xmldoc.XML)
            Call SendWindowsMessageTimeOut(destwindow, sourcewindow, xmldoc.xml)
        End If
    End If

End Sub

