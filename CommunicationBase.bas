Attribute VB_Name = "CommunicationBase"
Option Explicit

Public Const IRIS_MAX_RU_SIZE = 16384
Public Const IRIS_MAX_RU_SIZE_new = 32768
Public Const IRIS_OFFSET = 82
Public Const IRIS_OUTPUT_VIEW_POS = 67

Public Const online_MAX_RU_SIZE = 32768

Public Const COM_BASE = 100
Public Const CONNECT_BASE = 200
Public Const SEND_BASE = 300
Public Const RECEIVE_BASE = 400
Public Const PARSE_BASE = 500
Public Const DISCONNECT_BASE = 600
Public Const LOGON_BASE = 700

Public Const GENERIC_COM_ERROR = 999

'translation table type
Public Const UCS_OLD = 1
Public Const UCS_NEW = 2

'communicate status type
Public Const COM_OK = COM_BASE
Public Const COM_FAILED = COM_BASE + 1
Public Const COM_RUNTIME_ERROR = COM_BASE + 2
Public Const COM_USER_TERMINATED = COM_BASE + 3

'connect status type
Public Const CONNECT_OK = CONNECT_BASE
Public Const CONNECT_FAILED = CONNECT_BASE + 1
Public Const CONNECT_RUNTIME_ERROR = CONNECT_BASE + 2
Public Const CONNECT_ALREADY_CONNECTED = CONNECT_BASE + 3

'send status type
Public Const SEND_OK = SEND_BASE
Public Const SEND_NO_CONNECTION = SEND_BASE + 1
Public Const SEND_NO_DATA = SEND_BASE + 2
Public Const SEND_FAILED = SEND_BASE + 3
Public Const SEND_RUNTIME_ERROR = SEND_BASE + 4

'send status type
Public Const RECEIVE_OK = RECEIVE_BASE
Public Const RECEIVE_NO_CONNECTION = RECEIVE_BASE + 1
Public Const RECEIVE_FAILED = RECEIVE_BASE + 2
Public Const RECEIVE_RUNTIME_ERROR = RECEIVE_BASE + 3

Public Const RECEIVE_AUTH_FAILED = RECEIVE_BASE + 4

'parse status type
Public Const PARSE_OK = PARSE_BASE
Public Const PARSE_READ_AGAIN = PARSE_BASE + 1
Public Const PARSE_ANSWER_REQUIRED = PARSE_BASE + 2
Public Const PARSE_BAD_RECEIVED_DATA = PARSE_BASE + 3
Public Const PARSE_CHIEF_TELLER_REQUIRED = PARSE_BASE + 4
Public Const PARSE_CANCEL = PARSE_BASE + 5
Public Const PARSE_TRANSACTION_COMPLETED = PARSE_BASE + 6
Public Const PARSE_RUNTIME_ERROR = PARSE_BASE + 7
Public Const PARSE_SENSE_CODE = PARSE_BASE + 8
Public Const PARSE_HOST_REJECTION = PARSE_BASE + 9
Public Const PARSE_ANSWER_REQUIRED_DATA = PARSE_BASE + 10
Public Const PARSE_SEND_AGAIN = PARSE_BASE + 11
Public Const PARSE_MANAGER_REQUIRED = PARSE_BASE + 12

Public Const DISCONNECT_OK = DISCONNECT_BASE
Public Const DISCONNECT_FAILED = DISCONNECT_BASE + 1
Public Const DISCONNECT_RUNTIME_ERROR = DISCONNECT_BASE + 2

Public Const LOGON_OK = LOGON_BASE
Public Const LOGON_SEND_FAILED = LOGON_BASE + 1
Public Const LOGON_RECEIVE_FAILED = LOGON_BASE + 2

Public tranlation_type As Integer

Public ContinueCommunication As Boolean
Public ResetKey As Boolean

Public SenseCodeMessage As String
Public SenseCode As String

Public Function DecodeSenseCode(StrInput As String) As String
    Dim astr As String
    Dim DFH As Integer
    Dim DFHhex As String
    SenseCode = StrPad_(Hex(Asc(Mid(StrInput, 1, 1))) & Hex(Asc(Mid(StrInput, 2, 1))), 4, "0", "L")
    DFHhex = "&H" & StrPad_(Hex(Asc(Mid(StrInput, 3, 1))) & Hex(Asc(Mid(StrInput, 4, 1))), 4, "0", "L")
    DFH = DFHhex
    astr = "SENSE CODE:" & SenseCode & "  DFH:" & Str(DFH)
    If SenseCode = "008F" Then
        astr = astr & " δεν έχετε πρόσβαση στη συναλλαγή"
    ElseIf SenseCode = "0103" Then
        astr = astr & " Αγνωστη Συναλλαγή"
    ElseIf SenseCode = "0824" Then
        astr = astr & " ABEND Πρόβλημα Συναλλαγής"
    End If
    DecodeSenseCode = astr
End Function

Public Function CTGComAreaCom_(Modulename As String, InputView, InputViewName As String, OutputViewName As String, _
    Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount, Optional UpdateTrnCountFlag As Boolean) As cSNAResult
    
    Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
    
    Set CTGComAreaCom_ = New cSNAResult
    CTGComAreaCom_.ErrCode = 0
    
    If Not Flag610 Then
        eJournalWrite "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ " & cTRNCode
        CTGComAreaCom_.ErrMessage = "Δεν έχει γίνει σύνδεση (0610)": CTGComAreaCom_.ErrCode = GENERIC_COM_ERROR: Exit Function
    End If
    
    If IsMissing(UpdateTrnCountFlag) Then UpdateTrnCountFlag = True
    If UpdateTrnCountFlag Then UpdateTrnNum_
    
    InputView.v2Value("TRANID") = "XXXX"
    If Not IsMissing(Appltran) Then InputView.v2Value("TRANID") = Appltran
    InputView.v2Value("REQ_TYPE") = "COMM"
    InputView.v2Value("APPL_PGM") = UCase(Modulename)
    
    InputView.v2Value("USER_ID") = UCase(cIRISUserName)
    InputView.v2Value("WS_ID") = UCase(cIRISComputerName)
    InputView.v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
    
    InputView.v2Value("TRAN_SCD") = cHEAD
    
    If Trim(InputView.v2Value("TRAN_KEY")) = "" Then
        InputView.v2Value("TRAN_KEY") = "TELL"
    End If
    
    If Trim(InputView.v2Value("AUTH_USER")) = "" Then
        If Left(Trim(InputView.v2Value("TRAN_KEY")), 4) = "CHIE" Then
            InputView.v2Value("AUTH_USER") = UCase(cCHIEFUserName)
        ElseIf Left(Trim(InputView.v2Value("TRAN_KEY")), 4) = "MANA" Then
            InputView.v2Value("AUTH_USER") = UCase(cMANAGERUserName)
        End If
    Else
        InputView.v2Value("AUTH_USER") = UCase(InputView.v2Value("AUTH_USER"))
    End If
    
    If Trim(InputView.v2Value("AUTH_TRANS")) = "" Then
        Dim codtx As String
        codtx = GetCodTx(Modulename)
        InputView.v2Value("AUTH_TRANS") = codtx
    End If
    
    InputView.v2Value("IDFLEN") = Len(InputView.ByName(InputViewName).Data)
    If OutputViewName <> "" Then
        InputView.v2Value("ODFLEN") = Len(InputView.ByName(OutputViewName).Data)
    Else
        InputView.v2Value("ODFLEN") = 0
    End If
    
    InputView.v2Value("ATERM_ID") = Encode_Greek_(cTERMINALID)
    InputView.v2Value("BRANCH") = UCase(cBRANCH)
    
    astr = astr & InputView.Data
    eJournalWrite "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ " & cTRNCode
    
    Dim connector As New cCTGConnection
    connector.OpClass = "COMAREA_CTG"
    Dim adescription As String
        
    adescription = InputView.name
    
    If Not IsMissing(Appltran) Then
        connector.OpCode = Appltran
    Else
       connector.OpCode = Modulename
    End If
    connector.OpDescription = adescription
    If Not IsMissing(AuthUser) Then
        connector.AuthUser = Trim(AuthUser)
    Else
        connector.AuthUser = Trim(InputView.v2Value("AUTH_USER"))
    End If
    
    Set CTGComAreaCom_ = connector.SimpleExec(astr)
    If (CTGComAreaCom_.ErrCode = 0 Or CTGComAreaCom_.ErrCode = COM_OK) And CTGComAreaCom_.SenseCodeMessage = "" Then
        InputView.Data = Right(connector.ReceiveData, Len(InputView.Data))
    End If

End Function

Public Function ComAreaCom_(Modulename As String, InputView, InputViewName As String, OutputViewName As String, _
    Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount, Optional UpdateTrnCountFlag As Boolean) As cSNAResult
Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
Dim onlineAuthError As String
    Set ComAreaCom_ = New cSNAResult
    ComAreaCom_.ErrCode = 0
   
    If Not Flag610 Then
        eJournalWrite "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ " & cTRNCode
        ComAreaCom_.ErrMessage = "Δεν έχει γίνει σύνδεση (0610)": ComAreaCom_.ErrCode = GENERIC_COM_ERROR: Exit Function
    End If

    If IsMissing(UpdateTrnCountFlag) Then UpdateTrnCountFlag = True
    If UpdateTrnCountFlag Then UpdateTrnNum_
    
    onlineAuthError = ""
    InputView.v2Value("TRANID") = "XXXX"
    If Not IsMissing(Appltran) Then InputView.v2Value("TRANID") = Appltran
    InputView.v2Value("TRAN_SCD") = cHEAD
    
    If Trim(InputView.v2Value("TRAN_KEY")) = "" Then InputView.v2Value("TRAN_KEY") = "TELLER"
    
    If Trim(InputView.v2Value("AUTH_USER")) = "" Then
        If Left(Trim(InputView.v2Value("TRAN_KEY")), 4) = "CHIE" Then
            InputView.v2Value("AUTH_USER") = UCase(cCHIEFUserName)
        ElseIf Left(Trim(InputView.v2Value("TRAN_KEY")), 4) = "MANA" Then
            InputView.v2Value("AUTH_USER") = UCase(cMANAGERUserName)
        End If
    Else
        InputView.v2Value("AUTH_USER") = UCase(InputView.v2Value("AUTH_USER"))
    End If
    
    If Trim(InputView.v2Value("AUTH_TRANS")) = "" Then
        Dim codtx As String
        codtx = GetCodTx(Modulename)
        InputView.v2Value("AUTH_TRANS") = codtx
    End If
    
    '***********************************************
    InputView.v2Value("USER_ID") = UCase(cIRISUserName)
    InputView.v2Value("WS_ID") = UCase(cIRISComputerName)
    InputView.v2Value("ATERM_ID") = Encode_Greek_(cTERMINALID)
    InputView.v2Value("YPHRESIA") = cDepartment
    InputView.v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
    InputView.v2Value("APPL_PGM") = UCase(Modulename)
    InputView.v2Value("I_LEN") = Len(InputView.ByName(InputViewName).Data)
    If OutputViewName <> "" Then
        InputView.v2Value("O_LEN") = Len(InputView.ByName(OutputViewName).Data)
    Else
        InputView.v2Value("O_LEN") = 0
    End If
    
    If Left(Modulename, 1) = "P" Then
    
    ElseIf Left(Modulename, 1) = "S" Then
        InputView.v2Value("BRANCH") = UCase(cBRANCH)
    End If
    
    astr = astr & InputView.Data
    
     eJournalWrite "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ " & cTRNCode
'----------------------------------------------------------------------
    Dim connector As New cSNAConnection
    
    connector.OpClass = "COMAREA"
    Dim adescription As String
        
    adescription = InputView.name
    If Len(adescription) > 2 Then
        If Right(adescription, 2) = "_I" Then adescription = Left(adescription, Len(adescription) - 2)
    End If
    If Not IsMissing(Appltran) Then
        connector.OpCode = Appltran 'InputView.StructID 'v2Value("TRANID")
    Else
       connector.OpCode = Modulename
    End If
    connector.OpDescription = adescription
    If Not IsMissing(AuthUser) Then
        connector.AuthUser = Trim(AuthUser)
    Else
        connector.AuthUser = Trim(InputView.v2Value("AUTH_USER"))
    End If
    
    Set ComAreaCom_ = connector.SimpleExec(astr)
    If (ComAreaCom_.ErrCode = 0 Or ComAreaCom_.ErrCode = COM_OK) And ComAreaCom_.SenseCodeMessage = "" Then
        InputView.Data = Right(connector.ReceiveData, Len(InputView.Data))
    End If
    
End Function

Public Function onlineComErrorHandler(owner As Form, statusmessage As StatusBar, ComResult As Integer, Data As String)
    If ComResult = SEND_FAILED Then
        If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
        Report_ComError owner
    End If
    If ComResult = RECEIVE_FAILED Then
        If Len(Data) = 4 Then
            Dim SenseCodeMessage As String
            SenseCodeMessage = DecodeSenseCode(Data)
            If Not (owner Is Nothing) Then owner.sbWriteStatusMessage SenseCodeMessage
            eJournalWrite "Err:" & SenseCodeMessage & "ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ"
        Else
            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
            Report_ComError owner
        End If
    End If
    
End Function
