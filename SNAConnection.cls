VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSNAConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sData As String
Private slength As Long
Private rdata As String
Private rlength As Long

Private SendResult As Integer
Private ReceiveResult As Integer

Public StartTime, EndTime
Public StartTickCount
Public EndTickCount

Public SenseCodeMessage As String
Public ErrorMessage As String

Public OpClass As String
Public OpCode As String
Public OpDescription As String
Public AuthUser As String


Private Function SEND(inputStr As String) As Integer
    Dim pRetString As String
    Dim res As Integer
    Dim Bytelist()  As Byte
    Dim i As Integer
    
    Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
        pConvert As Long, pTimeOut As Long, pDebug As Long
    
    sData = inputStr: slength = Len(inputStr)
    rdata = "": rlength = 0
    cb.LUADirection = 1:
    cb.send_str = sData & Chr$(0)
    'cb.send_str_length = slength
    cb.receive_str = ""
    cb.receive_str_length = 0
    
    On Error GoTo JErrorHandler
    
    If Len(inputStr) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function
    SEND = SEND_OK
    pMsgType = 1: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0
    pTimeOut = cb.TimeOut: pDebug = cb.com_debug
    pConvert = 0: pRetString = String$(512, 0)
    On Error GoTo CErrorHandler
    If slength <= online_MAX_RU_SIZE Then
        pRetString = VB4SLISEND(sData & Chr$(0), pConvert, pTimeOut, slength, _
            pMsgType, pRet1, pRet2, pRetCode, pDebug)
        cb.LUADirection = pRetCode
    End If
    On Error GoTo JErrorHandler
    If pRet1 <> 0 Or pRet2 <> 0 Then
        SEND = SEND_FAILED
        Exit Function
    End If
    GenWorkForm.ComTimer.Enabled = False
    Exit Function
JErrorHandler:
    Call Runtime_error("Journal", Err.number, Err.description)
    SEND = SEND_RUNTIME_ERROR

CErrorHandler:
    Call Runtime_error("Send", Err.number, Err.description)
    SEND = SEND_RUNTIME_ERROR
    
End Function

Private Function Receive() As Integer
    On Error GoTo ErrorHandler
    Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
        pConvert As Long, pTimeOut As Long, pDebug As Long
    Dim pData As String, pLen As Long
    
    Receive = RECEIVE_OK
    pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0
    pTimeOut = cb.TimeOut: pDebug = cb.com_debug
    pData = "": pLen = 0
    pRetCode = cb.LUADirection
    pData = String$(online_MAX_RU_SIZE + 1, 0)
    
    pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)
    rdata = pData
    rlength = Len(rdata)
    cb.receive_str = pData
    cb.receive_str_length = Len(pData)
    cb.LUADirection = pRetCode
    
    If Len(pData) = 4 Then 'Sense Code
        Receive = RECEIVE_FAILED
        TranslateSenseCode pData
        Exit Function
    End If
    If pRet1 <> 0 Or pRet2 <> 0 Then
        Receive = RECEIVE_FAILED
        Exit Function
    End If
    GenWorkForm.ComTimer.Enabled = False
    Exit Function
ErrorHandler:
    Call Runtime_error("Receive", Err.number, Err.description)
    Receive = RECEIVE_RUNTIME_ERROR
End Function

Private Sub TranslateSenseCode(StrInput As String)
    Dim DFH As Integer
    Dim DFHhex As String
    Dim SenseCode As String

    SenseCode = StrPad_(Hex(Asc(Mid(StrInput, 1, 1))) & Hex(Asc(Mid(StrInput, 2, 1))), 4, "0", "L")
    DFHhex = "&H" & StrPad_(Hex(Asc(Mid(StrInput, 3, 1))) & Hex(Asc(Mid(StrInput, 4, 1))), 4, "0", "L")
    DFH = DFHhex
    SenseCodeMessage = "SENSE CODE:" & SenseCode & "  DFH:" & Str(DFH)
    If SenseCode = "008F" Then
       SenseCodeMessage = SenseCodeMessage & " δεν έχετε πρόσβαση στη συναλλαγή"
    ElseIf SenseCode = "0103" Then
        SenseCodeMessage = SenseCodeMessage & " Αγνωστη Συναλλαγή"
    ElseIf SenseCode = "0824" Then
        SenseCodeMessage = SenseCodeMessage & " ABEND Πρόβλημα Συναλλαγής"
    End If
End Sub

Private Function GenSNAPoolLink(SendLength As Long, ByVal data As String) As cSNAResult
   
    Set GenSNAPoolLink = New cSNAResult
    
    GenSNAPoolLink.ErrCode = 0: GenSNAPoolLink.ErrMessage = "": GenSNAPoolLink.SenseCodeMessage = ""
    SendResult = SEND_OK: ReceiveResult = RECEIVE_OK
    
    sData = data: slength = Len(sData)
    rdata = "": rlength = 0
    
    Dim linkurl As String
    linkurl = "http://w000010035:8001/SNAPoolLink/SNALink.asmx"
    linkurl = "http://localhost:8002/SNALinkWS/SNALink.asmx"
    Dim linkName As String
    linkName = "SNAPoolLink"
    Dim debugsnaurl As String
    
    On Error Resume Next
    If DebugSNAPoolLink Then debugsnaurl = WebLinks(UCase("DebugSNAPoolLink"))
    On Error GoTo 0
    
    If debugsnaurl = "" Then
        If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
            linkurl = WebLinks(UCase("EDUC" & linkName))
        ElseIf Left(Right(WorkEnvironment_, 8), 4) = "PROD" Then
            linkurl = WebLinks(UCase("PROD" & linkName))
        Else
            linkurl = ""
        End If
    Else
        linkurl = debugsnaurl
    End If
    
    Dim w As New cExecTRNWrapper
    w.WebServiceUrl = linkurl
    w.OpClass = OpClass
    w.OpCode = OpCode
    w.OpDescription = OpDescription
    w.MessageID = GetGuid
    GenSNAPoolLink.MessageID = w.MessageID
    w.AuthUser = AuthUser
    rdata = w.SendString(SendLength, data)
    rlength = Len(rdata)
End Function

Public Sub HandleError(owner)
    If SendResult = SEND_FAILED Then
        If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
        Report_ComError owner
    End If
    If ReceiveResult = RECEIVE_FAILED Then
        If Len(rdata) = 4 Then
            Dim SenseCodeMessage As String
            SenseCodeMessage = DecodeSenseCode(rdata)
            If Not (owner Is Nothing) Then owner.sbWriteStatusMessage SenseCodeMessage
            eJournalWrite "Err:" & SenseCodeMessage & "ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ"
        Else
            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
            Report_ComError owner
        End If
    End If
    
End Sub

Public Function Exec(OwnerForm, data As String) As cSNAResult
    On Error GoTo sHappens
    If Not (OwnerForm Is Nothing) Then OwnerForm.Enabled = False
    Set Exec = SimpleExec(data)
sHappens:
    If Not (OwnerForm Is Nothing) Then OwnerForm.Enabled = True
End Function

Public Function SimpleExec(data As String) As cSNAResult
    Set SimpleExec = New cSNAResult
    SimpleExec.ErrCode = 0
    If LogIrisCom Then sbWriteLogFileNew "SNAConnection.Send", data
        
    Dim tempdata As String
    Set SimpleExec = GenSNAPoolLink(Len(data), data)
           
    HandleResult SimpleExec
End Function

Public Property Get SendOk() As Boolean
    SendOk = (SendResult = SEND_OK)
End Property

Public Property Get ReceiveOk() As Boolean
    ReceiveOk = (ReceiveResult = RECEIVE_OK)
End Property

Public Property Get SendData() As String
    SendData = sData
End Property

Public Property Get ReceiveData() As String
    ReceiveData = rdata
End Property

Private Sub HandleResult(Result As cSNAResult)
    Result.SenseCodeMessage = SenseCodeMessage
    If SendOk And ReceiveOk And SenseCodeMessage = "" Then
        If ReceiveData = "" Then
            Result.ErrMessage = "Δεν επέστρεψαν δεδομένα.": Result.ErrCode = GENERIC_COM_ERROR: Exit Sub
        Else
            If OpClass <> "LEGACY" And Len(ReceiveData) <> Len(SendData) Then
                If Len(ReceiveData) = 160 Then
                    Dim adoc As MSXML2.DOMDocument30
                    Dim Node As IXMLDOMElement
                    Dim datanode As IXMLDOMElement
                    Dim datadoc As MSXML2.DOMDocument30
                    
                    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
                    If adoc Is Nothing Then
                        Result.ErrMessage = "Invalid XmlBlocks.xml file": Result.ErrCode = GENERIC_COM_ERROR: Exit Sub
                    End If
                    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='HPSERROR']", "HPSERROR", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
                    If Node Is Nothing Then
                        Result.ErrMessage = "HPSError not found in XmlBlocks.xml file": Result.ErrCode = GENERIC_COM_ERROR: Exit Sub
                    End If
                    
                    Dim errComArea As New cXmlComArea
                    Dim errBuffers As New Buffers
                    Set errComArea.content = Node
                    Set errComArea.Container = errBuffers
                    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "HPSERROR", "Πρόβλημα στη διαδικασία Σύνδεσης...")
                    If datanode Is Nothing Then
                        errBuffers.ClearAll
                        Result.ErrMessage = "HPSError/data/comarea not found in XmlBlocks.xml file": Result.ErrCode = GENERIC_COM_ERROR: Exit Sub
                    End If
                    Set datadoc = XmlLoadString(datanode.XML, "DataDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
                    If datadoc Is Nothing Then
                        errBuffers.ClearAll
                        Result.ErrMessage = "Invalid Xml Data: [" & datadoc & "].": Result.ErrCode = GENERIC_COM_ERROR: Exit Sub
                    End If
                    errComArea.LoadXML datadoc.XML
                    errComArea.Buffer.data = ReceiveData
                    
                    Result.ErrMessage = errComArea.Buffer.v2Value("MESSAGE"): Result.ErrCode = GENERIC_COM_ERROR:
                    errBuffers.ClearAll
                    Exit Sub
                Else
                    Result.ErrMessage = "Λάθος μέγεθος μηνύματος. Ζητήθηκαν: " & Len(SendData) & " Στάλθηκαν: " & Len(ReceiveData) & " bytes."
                    Result.ErrCode = GENERIC_COM_ERROR
                    Exit Sub
                End If
            Else
            End If
        End If
    Else
        Result.ErrCode = GENERIC_COM_ERROR
        If Not SendOk Then
            Result.ErrMessage = "ΑΠΕΤΥΧΕ Η ΑΠΟΣΤΟΛΗ ΣΤΟΙΧΕΙΩΝ...."
        ElseIf Not ReceiveOk And Result.SenseCodeMessage = "" Then
            Result.ErrMessage = "ΠΡΟΣΟΧΗ! ΑΠΕΤΥΧΕ Η ΛΗΨΗ ΑΠΑΝΤΗΣΗΣ.ΚΑΝΤΕ ΕΛΕΓΧΟ ΓΙΑ ΤΗΝ ΤΥΧΗ ΤΗΣ ΣΥΝΑΛΛΑΓΗΣ."
        Else
            Result.ErrMessage = Result.SenseCodeMessage
        End If
    End If
End Sub
