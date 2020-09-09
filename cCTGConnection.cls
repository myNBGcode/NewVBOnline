VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCTGConnection"
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

Private Function GenCTGLink(SendLength As Long, ByVal Data As String) As cSNAResult
    
    Set GenCTGLink = New cSNAResult
    
    GenCTGLink.ErrCode = 0: GenCTGLink.ErrMessage = "": GenCTGLink.SenseCodeMessage = ""
    SendResult = SEND_OK: ReceiveResult = RECEIVE_OK
    
    sData = Data: slength = Len(sData)
    rdata = "": rlength = 0
    
    Dim linkurl As String
    Dim linkName As String
    linkName = "CTGLink"
    Dim debugsnaurl As String
    
    On Error Resume Next
    If DebugSNAPoolLink Then debugsnaurl = WebLinks(UCase("DebugSNAPoolLink"))
    On Error GoTo 0
    
    'debugsnaurl = "http://localhost:31405/CTGLink.asmx"
    
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
    
    Dim w As New cExecCTGTRNWrapper
    w.WebServiceUrl = linkurl
    w.OpClass = OpClass
    w.OpCode = OpCode
    w.OpDescription = OpDescription
    w.AuthUser = AuthUser
    rdata = w.SendString(SendLength, Data)
    rlength = Len(rdata)
End Function

Public Function SimpleExec(Data As String) As cSNAResult
    Set SimpleExec = New cSNAResult
    SimpleExec.ErrCode = 0
    If LogIrisCom Then sbWriteLogFileNew "CTGConnection.Send", Data
    Set SimpleExec = GenCTGLink(Len(Data), Data)
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
                    errComArea.Buffer.Data = ReceiveData
                    
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

