Attribute VB_Name = "ComTest"
Option Explicit
Private Const COM_BASE = 100
Private Const CONNECT_BASE = 200
Private Const SEND_BASE = 300
Private Const RECEIVE_BASE = 400
Private Const PARSE_BASE = 500
Private Const DISCONNECT_BASE = 600
Private Const LOGON_BASE = 700

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
Private Old_Key As String * 1

Public Function AsciiToEbcdic_(inputStr As String) As String
Dim InputAscii As String, OutputStr As String
Dim i As Integer

InputAscii = inputStr & Chr$(0)

ASCII_CP_STRING = ""
    ' ο χαρακτηρας είναι 255 έχει γραφεί με το alt και 255 από το αριθμητικό πληκτρολογιο
    For i = 1 To 31
        ASCII_CP_STRING = ASCII_CP_STRING & " "
    Next
    ASCII_CP_STRING = ASCII_CP_STRING & " !" & Chr$(34) & "#$%&'()*+,-./0123456789:;<=>?" & _
    "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €" & _
    "‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©ª«¬­®―°±²³΄µ¶·ΈΉΊ»Ό½ΎΏΐΑ" & _
    "ΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡÒΣΤΥΦΧΨΩΪΫάέήίΰαβγδεζηθικλμνξοπρςστυφχψωϊϋόύώÿ" & Chr$(0)
    EBCDIC_CP_STRING = ""
    ' ο χαρακτηρας είναι 255 έχει γραφεί με το alt και 255 από το αριθμητικό πληκτρολογιο
    For i = 1 To 63
        EBCDIC_CP_STRING = EBCDIC_CP_STRING & " "
    Next
    EBCDIC_CP_STRING = EBCDIC_CP_STRING & " ΑΒΓΔΕΖΗΘΙ[.<(+!&ΚΛΜΝΞΟΠΡΣ]$*);^-/ΤΥΦΧΨΩΪΫ|,%_>?" & _
    "¨ΆΈΉ ΊΌΎΏ`:#@'=" & Chr$(34) & "΅abcdefghiαβγδεζ°jklmnop" & _
    "qrηθικλμ`~stuvwxyzνξοπρσ£άέήϊίόύϋώςτυφχψ{ABCDEFGHI­ωΐΰ‘-}JKLMNOPQ" & _
    "R±½ ·’|\ STUVWXYZ²§  «¬0123456789³©  » " & Chr$(0)

'MsgBox InputAscii
'MsgBox EBCDIC_CP_STRING
'MsgBox ASCII_CP_STRING


OutputStr = GKTranslate(InputAscii, EBCDIC_CP_STRING, ASCII_CP_STRING)
'MsgBox "end ascci2"
AsciiToEbcdic_ = Left(OutputStr, Len(inputStr))

End Function

Public Function EbcdicToAscii_(inputStr As String) As String
Dim InputAscii As String, OutputStr As String

InputAscii = inputStr & Chr$(0)
OutputStr = GKTranslate(InputAscii, ASCII_CP_STRING, EBCDIC_CP_STRING)
EbcdicToAscii_ = Left(OutputStr, Len(inputStr))
End Function

Public Function IntToHps_(ByVal InputInt As Long) As String
    Dim i As Integer, StrOUT As String, Step As Long, Resd As Long
    StrOUT = ""
    Step = InputInt
    For i = 1 To 4
        Resd = Step Mod 256
        Step = Step \ 256
        StrOUT = Chr$(Resd) & StrOUT
    Next i
    IntToHps_ = StrOUT
End Function

Public Function HpsToInt_(ByVal InputInt As String) As Long
    Dim i As Integer, ValOut As Long
    ValOut = 0
    For i = 1 To 4
        ValOut = ValOut * 256 + Asc(Mid(InputInt, i, 1))
    Next i
    HpsToInt_ = ValOut
End Function

Public Function initialize_cb() As Boolean

'Dim connect_status As Integer
Dim OK As Boolean
OK = True
cb.TimeOut = StrPad_(cTimeOut, 4, "0", "L")  '20 read from file
cb.Ret1 = 0
cb.Ret2 = 0
cb.RetCode = 0
cb.LUADirection = 0
cLUName = StrPad_(cLUName, 8, " ", "R") & Chr(0) ' "T4001003" + Chr$(0) 'read from file

'biks
'cb.send_convert = 1
cb.send_convert = 0
'biks
cb.receive_convert = 1
cb.DecodeGreek = 1
cb.EncodeGreek = 1
cb.TimeOut = 180

'connect_status = CONNECT()
'If (connect_status <> CONNECT_OK And connect_status <> CONNECT_ALREADY_CONNECTED) Then
'    Call NBG_MsgBox("ΔΕΝ ΕΠΕΤΕΥΧΘΗ ΕΠΙΚΟΙΝΩΝΙΑ " & Str(connect_status))
'    Call NBG_error("initialize_cb", connect_status)
'    OK = False
'End If

initialize_cb = OK
End Function

Public Function communicate(owner As Form) As Integer
 
On Error GoTo ErrorHandler

Dim Send_status As Integer
Dim receive_status As Integer
Dim parse_status As Integer
Dim Result As Boolean
Dim WrongFieldLabel As String
Dim iCount As Integer
Dim CurrentIndex As Integer
Dim i As Integer

Do   'communicate loop

communicate = COM_OK


parse_status = PARSE_READ_AGAIN
Screen.ActiveForm.sbWriteStatusMessage "ΔΙΑΒΙΒΑΣΗ ΔΕΔΟΜΕΝΩΝ. ΠΕΡΙΜΕΝΕΤΕ..."

Send_status = SEND(owner)
If Send_status <> SEND_OK Then
    communicate = COM_FAILED

    Exit Function
End If

Do   'receive loop
    cb.TransTerminating = False
    cb.read_again = False
    cb.receive_str = ""
    cb.receive_str_length = 0
    receive_status = RECEIVE(owner)
    If receive_status <> RECEIVE_OK Then
        communicate = COM_FAILED
        Screen.ActiveForm.sbWriteStatusMessage "ΛΑΘΟΣ ΑΝΑΚΤΗΣΗΣ !!! " & Str(receive_status)
        Exit Function
    End If
    
    If Mid(cb.receive_str, 5, 3) = "DFH" Then
        Screen.ActiveForm.sbWriteStatusMessage Right(cb.receive_str, cb.receive_str_length - 4) 'Print data to screen
        Select Case Mid(cb.receive_str, 8, 5)
            Case "3506I"
                cb.read_again = True
                cb.TransTerminating = True
            Case "3512I"
                cb.TimeOut = 1
                cb.send_str = "&H08240000"
                cb.TransTerminating = True
                parse_status = PARSE_SEND_AGAIN
                cb.BoolTransOk = False
            Case "3500I"
                cb.BoolTransOk = False
                Exit Do
            Case Else
                cb.BoolTransOk = True
                Exit Do
        End Select
    End If
    
    If cb.TransTerminating Then GoTo ParsingEnd

    parse_status = parse_phase_1(owner)
    
    Select Case parse_status
        Case PARSE_SENSE_CODE
            Call read_sense_code(cb.receive_str)  'owner,
            cb.receive_convert = 1
            cb.BoolTransOk = False
            GoTo TerminateRead
        Case PARSE_BAD_RECEIVED_DATA
            communicate = COM_FAILED
            Screen.ActiveForm.sbWriteStatusMessage "ΛΑΘΟΣ ΣΤΟΙΧΕΙΑ !!! " & Str(parse_status)
            GoTo TerminateRead
        Case PARSE_OK
            'do nothing
        Case PARSE_CHIEF_TELLER_REQUIRED
            ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False
            
'            If Not isChiefTeller Then
'                'Set SelKeyFrm.Owner = TRNFrm
'                Set SelKeyFrm.owner = owner
'                ChiefRequest = True
'                'SelKeyFrm.Show vbModal, TRNFrm
'
'                SelKeyFrm.Show vbModal, owner
'                'Ειδικός χειρισμός για τις καταθέσεις συναλλαγματος
'                'πρέπει να πάρει οποσδήποτε έγκριση
'                While (cTRNCode = 5000 Or cTRNCode = 5001 Or cTRNCode = 5002 _
'                Or cTRNCode = 5100 Or cTRNCode = 5101 Or cTRNCode = 5102) And Not KeyAccepted
'                    SelKeyFrm.Show vbModal, owner
'                Wend
'                If Not KeyAccepted Then
'                    If cTRNCode <> 5000 And cTRNCode <> 5001 And cTRNCode <> 5002 _
'                    And cTRNCode <> 5100 And cTRNCode <> 5101 And cTRNCode <> 5102 Then
'                        Call TerminateTransaction(owner, Send_status)
'                        communicate = COM_USER_TERMINATED
'                    End If
'                Else
'                    cb.trn_key = cCHIEFKEY
'                End If
'            Else
'                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'
'                While (cTRNCode = 5000 Or cTRNCode = 5001 Or cTRNCode = 5002 _
'                Or cTRNCode = 5100 Or cTRNCode = 5101 Or cTRNCode = 5102) And Not KeyAccepted
'                    KeyWarning.Show vbModal, owner
'                Wend
'                If Not KeyAccepted Then
'                    If cTRNCode <> 5000 And cTRNCode <> 5001 And cTRNCode <> 5002 _
'                    And cTRNCode <> 5100 And cTRNCode <> 5101 And cTRNCode <> 5102 Then
'                        Call TerminateTransaction(owner, Send_status)
'                        communicate = COM_USER_TERMINATED
'                    End If
'                Else
'                    cb.trn_key = cCHIEFKEY
'                End If
'
'            End If
        Case PARSE_MANAGER_REQUIRED
            ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False
'            If Not isManager Then
'                'Set SelKeyFrm.Owner = TRNFrm
'                Set SelKeyFrm.owner = owner
'                ManagerRequest = True
'                'SelKeyFrm.Show vbModal, TRNFrm
'                SelKeyFrm.Show vbModal, owner
'                If Not KeyAccepted Then
'                    Call TerminateTransaction(owner, Send_status)
'                    communicate = COM_USER_TERMINATED
'                Else
'                    cb.trn_key = cTELLERMANAGERKEY
'                End If
'            Else
'                ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'
'                If Not KeyAccepted Then
'                    Call TerminateTransaction(owner, Send_status)
'                    communicate = COM_USER_TERMINATED
'                Else
'                    cb.trn_key = cCHIEFKEY
'                End If
'            End If
    End Select
    'MsgBox "Parse Phase 1 completed"
  
    If cb.TransTerminating Then GoTo ParsingEnd
    parse_status = parse_phase_2(owner)
    
    Select Case parse_status
        Case PARSE_OK
            'do nothing
    End Select
    'MsgBox "Parse Phase 2 completed"

    If cb.TransTerminating Then GoTo ParsingEnd
    parse_status = parse_phase_3()
    
    Select Case parse_status
        Case PARSE_READ_AGAIN
            'do nothing - receive loop will be repeated
        Case PARSE_ANSWER_REQUIRED
            'εμφάνιση οθόνης με ΟΚ και ΝΟ
'            If Not owner.SkipCommConfirmation Then
'                frmContinue.Show 1
'                If ContinueCommunication Then
'                    Call ContinueTransaction(owner, Send_status)
'                Else
'                    Call TerminateTransaction(owner, Send_status)
'                    communicate = COM_USER_TERMINATED
'                End If
'            Else
                Call ContinueTransaction(owner, Send_status)
'            End If
        
        Case PARSE_ANSWER_REQUIRED_DATA
            'εμφάνιση οθόνης με ΟΚ και ΝΟ
'            Load frmContinueData
'            Set frmContinueData.owner = owner
'            frmContinueData.Show 1
'            If Not ContinueCommunication Then
'                Call TerminateTransaction(owner, Send_status)
'                    communicate = COM_USER_TERMINATED
'            End If
        Case PARSE_HOST_REJECTION
            cb.BoolTransOk = False
            WrongFieldLabel = Mid(cb.receive_str, 3, 2)
            
            'να προστεθει κωδικας για τον εντοπισμο του πεδίου που εχει
            'λαθος τιμη, να ανοιγει φορμα που θα δημιουργεί το πεδίο αυτόματα
            'και θα στέλνει τη σωστή τιμή οπως η frmcontinuedata μονο που
            'θα κάνει format την τιμη πρώτα
            'προβλημα: Στην 1001, 1101, 1231 τα πεδία ποσό και υπόλοιπο θα πρέπει να φευγουν με
            'πρόσημο
'BIKSBIKS
'BIKSBIKS
'            CurrentIndex = Screen.ActiveForm.ActiveControl.Index
'            For iCount = 0 To cb.num_of_fields
'                If Strpin(iCount, 0) = WrongFieldLabel Then Exit For
'            Next iCount
'            If CurrentIndex <> iCount Then
'                Call Screen.ActiveForm.txtInput_LostFocus(CurrentIndex)
'            End If
'            Call FocusWrongInputField(Screen.ActiveForm, _
                                iCount, _
                                "Λανθασμένη Τιμή Πεδίου " & WrongFieldLabel)
'biks
            
'biks 28/11/99
'            frmContinueData.Show 1
'            If ContinueCommunication Then
'                Call ContinueTransaction(Send_status)
'            Else
'                Call TerminateTransaction(Send_status)
'            End If
'biks 28/11/99
            
            TerminateTransaction owner, Send_status
            communicate = COM_USER_TERMINATED
            Screen.ActiveForm.sbWriteStatusMessage " ΛΑΘΟΣ ΤΙΜΗ ΠΕΔΙΟΥ " & Mid(cb.receive_str, 3, 2)
            cb.receive_str = " ΛΑΘΟΣ ΤΙΜΗ ΠΕΔΙΟΥ " & Mid(cb.receive_str, 3, 2)
            
        Case PARSE_CANCEL
            cb.BoolTransOk = False
        Case PARSE_TRANSACTION_COMPLETED
            cb.BoolTransOk = True
            Screen.ActiveForm.sbWriteStatusMessage "Η ΣΥΝΑΛΛΑΓΗ ΟΛΟΚΛΗΡΩΘΗΚΕ."
        Case PARSE_SEND_AGAIN   ' ΤΟ ΠΡΟΣΘΕΣΑΜΕ
             'do nothing
    End Select
    'MsgBox "Parse Phase 3 completed"

ParsingEnd:
    If parse_status <> PARSE_READ_AGAIN And cb.inttime <> 0 Then cb.inttime = 0
'    GenWorkForm.vStatus.Panels(1).Text = ""
Loop While (parse_status = PARSE_READ_AGAIN Or cb.read_again = True)

Loop While (parse_status = PARSE_SEND_AGAIN)


TerminateRead:
' ΣΤΟ RECEIVE() ΓΙΝΕΤΑΙ ASSIGN H PRETCODE
' ΣΤΗΝ LUADIRECTION
Do While cb.LUADirection = 3
    cb.receive_str = ""
    cb.receive_str_length = 0
    receive_status = RECEIVE(owner)
    If receive_status <> RECEIVE_OK Then
        communicate = COM_FAILED
'biks
        Screen.ActiveForm.sbWriteStatusMessage "ΛΑΘΟΣ ΑΝΑΚΤΗΣΗΣ !!! " & Str(receive_status)
'        eJournalWriteAll owner, "ΛΑΘΟΣ ΑΝΑΚΤΗΣΗΣ !!! " & Str(receive_status) ', CStr(cTRNCode), cTRNNum
'biks
        Exit Function
    End If
Loop

If cb.inttime <> 0 Then
'   Call NBG_MsgBox("«ENTER» για συνέχεια", True, " ")
   cb.inttime = 0
End If

Exit Function

ErrorHandler:
'    Call Runtime_error("Communicate", Err.Number, Err.Description)
    communicate = COM_RUNTIME_ERROR

End Function

Public Function CONNECT() As Integer
Dim pRetString As String, res As Long

On Error GoTo ErrorHandler

'CONNECT = CONNECTEx
'Exit Function

'CLEARBUFFER
CONNECT = CONNECT_OK
'GenWorkForm.vStatus.Panels(2).Visible = True
'GenWorkForm.vStatus.Panels(3).Visible = False

pRetString = String$(512, 0)
'cb.LUName = "W12      " & Chr(0)
'cb.ApplId = "ABGDEZHU" & Chr(0)
cb.send_convert = 1

'pRetString = VB4SLICONNECT(cLUName, _
                        cb.ApplId, _
                        cb.send_convert, _
                        cb.TimeOut, _
                        cb.Ret1, _
                        cb.Ret2, _
                        cb.RetCode, _
                        cb.com_debug)
'cb.Ret1 = 256
'cb.Ret2 = 33554432
'cb.RetCode = 3

'res = VB4SLIReset(5)

'cb.com_debug = 2
pRetString = VB4SLICONNECT(cLUName, _
                        cb.ApplId, _
                        cb.send_convert, _
                        cb.TimeOut, _
                        cb.Ret1, _
                        cb.Ret2, _
                        cb.RetCode, _
                        cb.com_debug)

'If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
'    res = VB4SLIWAIT(10)
'End If

'If pRetString = "Timeout on connect" Then
'    res = VB4SLIWait(5)
'End If
'res = VB4SLIWait(5)

If cb.Ret1 = -1 And cb.Ret2 = -1 Then
    CONNECT = CONNECT_ALREADY_CONNECTED
    If EventLogWrite Then _
    Call EventLog(2, "VB Application :ATTEMPT TO CONNECT BUT ALREADY CONNECTED")
    Exit Function
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    CONNECT = CONNECT_FAILED
    If EventLogWrite Then _
    Call EventLog(1, "VB Application :CONNECT FAILED")
'    GenWorkForm.vStatus.Panels(2).Visible = False
'    GenWorkForm.vStatus.Panels(3).Visible = True
    Exit Function
End If

cb.receive_str = ""
cb.receive_str_length = 0

'GenWorkForm.ComTimer.Enabled = False
Exit Function

ErrorHandler:
Dim astr
    astr = Err.Description
'    Call Runtime_error("Connect", Err.Number, Err.Description)
    CONNECT = CONNECT_RUNTIME_ERROR

End Function

Public Function CONNECTEx() As Integer
Dim pRetString As String, res As Long

On Error GoTo ErrorHandler

'CLEARBUFFER
CONNECTEx = CONNECT_OK
'GenWorkForm.vStatus.Panels(2).Visible = True
'GenWorkForm.vStatus.Panels(3).Visible = False

pRetString = String$(512, 0)
'cb.LUName = "W12      " & Chr(0)
'cb.ApplId = "ABGDEZHU" & Chr(0)
cb.send_convert = 1

Dim aEvent As Long
aEvent = CreateEvent(0, 1, 0, "")
pRetString = VB4SLICONNECTEX(cLUName, _
                        cb.ApplId, _
                        cb.send_convert, _
                        10, _
                        cb.Ret1, _
                        cb.Ret2, _
                        cb.RetCode, _
                        2, aEvent)

res = WaitForSingleObject(aEvent, INFINITE)

'If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
'    res = VB4SLIWAIT(30)
'End If

'If pRetString = "Timeout on connect" Then
'    res = VB4SLIWait(5)
'End If
'res = VB4SLIWait(5)

If cb.Ret1 = -1 And cb.Ret2 = -1 Then
    CONNECTEx = CONNECT_ALREADY_CONNECTED
    If EventLogWrite Then _
    Call EventLog(2, "VB Application :ATTEMPT TO CONNECT BUT ALREADY CONNECTED")
    Exit Function
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    CONNECTEx = CONNECT_FAILED
    If EventLogWrite Then _
    Call EventLog(1, "VB Application :CONNECT FAILED")
'    GenWorkForm.vStatus.Panels(2).Visible = False
'    GenWorkForm.vStatus.Panels(3).Visible = True
    Exit Function
End If

cb.receive_str = ""
cb.receive_str_length = 0

'GenWorkForm.ComTimer.Enabled = False
Exit Function

ErrorHandler:
Dim astr
    astr = Err.Description
'    Call Runtime_error("Connect", Err.Number, Err.Description)
    CONNECTEx = CONNECT_RUNTIME_ERROR

End Function

Public Function DISCONNECT_() As Integer
Dim pRetString As String
Dim printer_status As Long

On Error GoTo ErrorHandler

DISCONNECT_ = DISCONNECT_OK
cb.Ret1 = 0
cb.Ret2 = 0
cb.RetCode = 0

pRetString = String$(512, 0)

'pRetString = VB4SLIDISCONNECT(cb.TimeOut, _
                           cb.Ret1, _
                           cb.Ret2, _
                           cb.RetCode, _
                           cb.com_debug)
                           
pRetString = VB4SLIDISCONNECT(180, _
                           cb.Ret1, _
                           cb.Ret2, _
                           cb.RetCode, _
                           cb.com_debug)
'cb.sid = 0

If cb.RetCode > 0 Then
'   Call NBG_MsgBox("LU Disconnected", True)
   cb.LUADirection = 0
Else
    DISCONNECT_ = DISCONNECT_FAILED
    If EventLogWrite Then _
    Call EventLog(1, "VB Application :DISCONNECT FAILED")
    Exit Function
End If
cb.receive_str = ""
cb.receive_str_length = 0

Exit Function

ErrorHandler:
'    Call Runtime_error("Disconnect", Err.Number, Err.Description)
    DISCONNECT_ = DISCONNECT_RUNTIME_ERROR

End Function


Public Function CONNECT_FAST() As Integer
Dim pRetString As String, res As Long

On Error GoTo ErrorHandler

CONNECT_FAST = CONNECT_OK

pRetString = String$(512, 0)
cb.send_convert = 1

pRetString = VB4SLICONNECT(cLUName, _
                        cb.ApplId, _
                        cb.send_convert, _
                        5, _
                        cb.Ret1, _
                        cb.Ret2, _
                        cb.RetCode, _
                        cb.com_debug)


If cb.Ret1 = -1 And cb.Ret2 = -1 Then
    CONNECT_FAST = CONNECT_ALREADY_CONNECTED
    If EventLogWrite Then _
    Call EventLog(2, "VB Application :ATTEMPT TO CONNECT BUT ALREADY CONNECTED")
    Exit Function
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    CONNECT_FAST = CONNECT_FAILED
    If EventLogWrite Then _
    Call EventLog(1, "VB Application :CONNECT FAILED")
    Exit Function
End If

cb.receive_str = ""
cb.receive_str_length = 0

Exit Function

ErrorHandler:
Dim astr
    astr = Err.Description
'    Call Runtime_error("Connect", Err.Number, Err.Description)
    CONNECT_FAST = CONNECT_RUNTIME_ERROR

End Function

Public Function DISCONNECT_ABEND() As Integer
Dim pRetString As String
Dim printer_status As Long

On Error GoTo ErrorHandler

DISCONNECT_ABEND = DISCONNECT_OK
cb.Ret1 = 1 ' 1:για abend - 0:για κανονικό disconnect
cb.Ret2 = 0
cb.RetCode = 0

pRetString = String$(512, 0)

'pRetString = VB4SLIDISCONNECT(cb.TimeOut, _
                           cb.Ret1, _
                           cb.Ret2, _
                           cb.RetCode, _
                           cb.com_debug)
                           
pRetString = VB4SLIDISCONNECT(180, _
                           cb.Ret1, _
                           cb.Ret2, _
                           cb.RetCode, _
                           cb.com_debug)
'cb.sid = 0

If cb.RetCode > 0 Then
'   Call NBG_MsgBox("LU Disconnected", True)
   cb.LUADirection = 0
Else
    DISCONNECT_ABEND = DISCONNECT_FAILED
    If EventLogWrite Then _
    Call EventLog(1, "VB Application :DISCONNECT FAILED")
    Exit Function
End If
cb.receive_str = ""
cb.receive_str_length = 0

Exit Function

ErrorHandler:
'    Call Runtime_error("Disconnect", Err.Number, Err.Description)
    DISCONNECT_ABEND = DISCONNECT_RUNTIME_ERROR

End Function

Public Sub Report_ComError() 'owner As Form
    
'    If Not GenWorkForm.ComTimer.Enabled Then
'        ComTimerCounter = 0: ComTimerCycle = 0
'        GenWorkForm.ComTimer.Interval = 6553
'        owner.sbShowCommStatus (False)
'        GenWorkForm.sbShowCommStatus (False)
'        GenWorkForm.ComTimer.Enabled = True
'    End If

End Sub

Public Function Restore_Connection() As Integer
Dim res As Long
    Restore_Connection = 0
    res = DISCONNECT_
    res = CONNECT_FAST
    If res = CONNECT_OK Then
        Restore_Connection = 1
    Else
        res = DISCONNECT_ABEND
    End If
End Function

Public Function RECEIVE(owner As Form) As Integer

On Error GoTo ErrorHandler

Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
    pConvert As Long, pTimeOut As Long, pDebug As Long
Dim pData As String, pData1 As String, pLen As Long, pDataTotal As String, _
    pLengthTotal As Long


RECEIVE = RECEIVE_OK
pDataTotal = ""
pLengthTotal = 0


pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0

pTimeOut = cb.TimeOut: pDebug = cb.com_debug

pData = "": pLen = 0

pRetCode = cb.LUADirection
pData = String$(4097, 0)

pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)

pData1 = Left$(pData, pLen)

cb.LUADirection = pRetCode

If pRet1 <> 0 Or pRet2 <> 0 Then
    RECEIVE = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError 'owner
    Exit Function
End If

If pLen > 0 And pMsgType <> 4 Then
    pDataTotal = pDataTotal & pData1
    pLengthTotal = pLengthTotal + pLen
End If

If pLengthTotal <= 0 Then
    pDataTotal = ""
End If

cb.receive_str = pDataTotal
cb.receive_str_length = pLengthTotal


'''''''''''''''''''''''''''''''''''
' Convert EBCDIC to ASCII

If cb.receive_convert = 1 Then
    If pLengthTotal <> 4 Then
        Call EbcdicToAscii
    End If
    If cb.CodePage = UCS_OLD And cb.DecodeGreek = 1 Then
        Call Decode_Greek
    End If
    ReceivedData.Add (cb.receive_str)
    
    If EventLogWrite Then _
    Call EventLog(8, " Translated RECEIVE :" & cb.receive_str)
'    If ReceiveJournalWrite Then eJournalWrite owner, "R:" & cb.receive_str
'    If Len(cb.receive_str) > 0 Then _
'        If Left(cb.receive_str, 1) = "0" Then eJournalWrite owner, cb.receive_str
End If
''''''''''''''''''''''''''''''''''''

cb.LUADirection = pRetCode

'GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
'    Call Runtime_error("Receive", Err.Number, Err.Description)
    RECEIVE = RECEIVE_RUNTIME_ERROR
    
End Function

Public Function SEND(owner As Form) As Integer
Dim pRetString As String
Dim pConvert As Integer, res As Integer

On Error GoTo JErrorHandler

SEND = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
'If Len(cb.send_str) < 1 Then
'    Call NBG_MsgBox("No data to send!!!", True):
'Exit Function
    
cb.RetCode = cb.LUADirection

cb.MsgType = 0
cb.send_str_length = Len(cb.send_str)
cb.send_str = cb.send_str & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0

''''''''''''''''''''''''''''''''
'Convert ASCII to ABCDIC

If cb.send_convert = 1 Then
    cb.initsend_str = cb.send_str
    If EventLogWrite Then Call EventLog(8, "Untranslated SEND :" & cb.send_str)
'    If SendJournalWrite Then eJournalWrite owner, "S:" & cb.send_str
    If cb.CodePage = UCS_OLD And cb.EncodeGreek = 1 Then Call Encode_Greek
    Call AsciiToEbcdic
End If
pConvert = 0
''''''''''''''''''''''''''''''''

pRetString = String$(512, 0)

On Error GoTo CErrorHandler
pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, cb.send_str_length, _
        cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, cb.com_debug)

cb.LUADirection = cb.RetCode

On Error GoTo JErrorHandler
If ResetKey Then
    cb.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    SEND = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    Report_ComError 'owner
    Exit Function
End If
    
'GenWorkForm.ComTimer.Enabled = False
Exit Function

JErrorHandler:
'    Call Runtime_error("Journal", Err.Number, Err.Description)
    SEND = SEND_RUNTIME_ERROR

CErrorHandler:
'    Call Runtime_error("Send", Err.Number, Err.Description)
    SEND = SEND_RUNTIME_ERROR
    
End Function

Private Sub Decode_Greek()
Dim pData As String, _
    HostLetters As String, _
    ASCILetters As String
'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh΅tyzsuvw" & Chr$(0)
'ASCILetters = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟ/ΠΡΣΤΥΦΧΨΩ" & "CDFGHJLQRSUVW" & Chr$(0)
HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfg~΅tyzsuvw" & Chr$(0)
ASCILetters = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟ/ΠΡΣΤΥΦΧΨΩ" & "CDFGJ@LQRSUVW" & Chr$(0)
pData = cb.receive_str & Chr$(0)
cb.receive_str = GKTranslate(pData, _
                         HostLetters, _
                         ASCILetters)
cb.receive_str = Left$(cb.receive_str, cb.receive_str_length)

End Sub
Private Sub Encode_Greek()
Dim pData As String, _
    HostLetters As String, _
    ASCILetters As String
'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh΅tyzsuvw" & Chr$(0)
'ASCILetters = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟ/ΠΡΣΤΥΦΧΨΩ" & "CDFGHJLQRSUVW" & Chr$(0)
HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgG~΅tyzsuvwMJLRVF" & Chr$(0)
ASCILetters = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟ/ΠΡΣΤΥΦΧΨΩ" & "CDFGHJ@LQRSUVWNKMPYZ" & Chr$(0)
pData = cb.send_str & Chr$(0)
cb.send_str = GKTranslate(pData, _
                         ASCILetters, _
                         HostLetters)
cb.send_str = Left$(cb.send_str, cb.send_str_length)

End Sub


Private Function parse_phase_1(owner As Form)
    Dim F_byte As String
    Dim G_byte As String
    Dim MF_Byte As String
    Dim inti As Integer
    
    
    parse_phase_1 = PARSE_OK
    
    If Len(cb.receive_str) < 5 Then
    
        If Len(cb.receive_str) = 4 Then
            parse_phase_1 = PARSE_SENSE_CODE
        Else
            parse_phase_1 = PARSE_BAD_RECEIVED_DATA
'            Screen.ActiveForm.sbWriteStatusMessage "ΛΑΘΟΣ ΣΤΟΙΧΕΙΑ !!! "
'            eJournalWriteAll owner, "ΛΑΘΟΣ ΣΤΟΙΧΕΙΑ !!! " ', CStr(cTRNCode), cTRNNum
            
'            Call NBG_MsgBox("ΛΑΘΟΣ ΣΤΟΙΧΕΙΑ: *" & cb.receive_str & "*", True)
        End If
        Exit Function
    End If
    
    
    F_byte = Mid(cb.receive_str, 5, 1)
    G_byte = Mid(cb.receive_str, 1, 1)
    MF_Byte = Mid(cb.receive_str, 2, 1)
    
    
    If Len(cb.receive_str) >= 6 Then
        cb.received_data = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
    Else
        cb.received_data = ""
    End If
    
    Select Case F_byte
        Case "0" 'Do nothing
        Case "1"
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
            If G_byte = "4" Then Screen.ActiveForm.sbWriteStatusMessage cb.received_data
        Case "2"
            If G_byte = "0" And MF_Byte = "2" Then
                G0Data.Add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                Screen.ActiveForm.sbWriteStatusMessage cb.received_data
            Else
                Screen.ActiveForm.sbWriteStatusMessage cb.received_data
                Comm_printJrn owner
            End If
        Case "3"
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
            
            If G_byte = "0" And MF_Byte = "2" Then
                G0Data.Add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                Screen.ActiveForm.sbWriteStatusMessage cb.received_data
            Else
                Screen.ActiveForm.sbWriteStatusMessage cb.received_data
                'Comm_printJrn owner
            End If
        Case "4"
'            If Not isChiefTeller Then
'                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            Else
'                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            End If
        Case "5"
'            If Not isChiefTeller Then
'                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            Else
'                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            End If
            
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
        Case "6"
'            If Not isChiefTeller Then
'                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            Else
'                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            End If
'biks
            Screen.ActiveForm.sbWriteStatusMessage cb.received_data
'            eJournalWriteAll owner, cb.received_data ', CStr(cTRNCode), cTRNNum
'biks
        Case "7"
'            If Not isChiefTeller Then
'                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
'            Else
'                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'                If KeyAccepted Then
'                    ResetKey = True
'                    Old_Key = cb.trn_key
'                    cb.trn_key = cCHIEFKEY
'                End If
'            End If
            
            
            'print_to_journal (received_data)
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
'biks
            Screen.ActiveForm.sbWriteStatusMessage cb.received_data
            'Comm_printJrn owner
'biks
        Case "8", "9", "A", "B", "C", "D", "E", "F", "Α", "Β", "Γ", "Δ", "Ε", "Ζ"
'            If Not isManager Then
'                parse_phase_1 = PARSE_MANAGER_REQUIRED
'            Else
'                ManagerRequest = True: ChiefRequest = False: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
'                If KeyAccepted Then
'                    ResetKey = True
'                    Old_Key = cb.trn_key
'                    cb.trn_key = cTELLERMANAGERKEY
'                End If
'            End If

            'print_to_journal (received_data)
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
    End Select

End Function

Private Function parse_phase_2(owner As Form)
    Dim G_byte As String
    Dim MF_Byte As String
    Dim MN_Bytes As String

    Dim OK As Boolean

    parse_phase_2 = PARSE_OK
    G_byte = Mid(cb.receive_str, 1, 1)
    MF_Byte = Mid(cb.receive_str, 2, 1)
    MN_Bytes = Mid(cb.receive_str, 3, 2)

    Select Case MF_Byte
        Case "3"
            Select Case MN_Bytes
                Case "00"
                    'do something
                Case "01"
                    ' Α Κ Α Τ Α Χ Ω Ρ Ι Σ Τ Ε Σ    ΕΓΓΡΑΦΕΣ
                    'OK = Print_PASSBOOK()
                Case "02"
                    'ΕΚΤΥΠΩΣΗ ΣΤΟ ΒΙΒΛΙΑΡΙΟ, ΥΠΟΔΟΧΗ
                    ' εαν mid(cb.received_data,5,1) = "1" eject βιβλιάριο
                Case "03"
                    ' ΣΥΓΚΡΙΣΗ ΠΟΣΟΥ  ΔΕΝ ΧΡΗΣΙΜΟΠΟΙΕΙΤΑΙ
                    ' ΣΥΓΚΡΙΣΗ ΥΠΟΛΟΙΠΟΥ ΓΙΑ PASSWORD ΠΡΟΪΣΤΑΜΕΝΟΥ
                    If cb.curr_transaction = "1131" Then
'                        If Val(Mid(cb.receive_str, 6, 12)) > cb.posolimit Then
'                            Do
'                                cb.TransTerminating = False
'biks
'                                Call ChangeChiefTellerState(frm1131)
'biks
'                            Loop While cb.TransTerminating = True
'                        End If
                    End If
                Case "04"
'                    eJournalWriteAll owner, cPOSTDATE & " ΗΜ/ΝΙΑ ΤΕΡ/ΚΟΥ" ', CStr(cTRNCode), cTRNNum)
                Case "05"
                    'ΠΑΙΡΝΩ ΑΠΟ ΤΟ ΚΜ ΤΟ STRING ΓΙΑ ΤΗΝ 0620
                    cHEAD = Mid(cb.receive_str, Len(cb.receive_str) - 4, 4)
                Case "07"
                    'ΔΗΜΙΟΥΡΓΩ ΤΟ STRING
                    cb.send_str = StrPad_(CStr(cTRNNum), 3, "0", "L") & cNextDateFlag & "00000000000000000000"
                Case "09"
                    'αθροιστης και ποσό από HOST
'BIKS
'                    cb.stringfor3130 = cb.received_data
'                    Call PrintJrnl(Mid(cb.received_data, 1, 1) & " " & _
                            StrPad(format_num(Val(Mid(cb.received_data, 2, 12))), 17) & _
                            IIf(Mid(cb.received_data, 13, 1) = "+", "ΠΧ", "ΠΠ"), False)
'BIKS
                Case "16"
                    If G_byte = "5" Then
'biks
'                        Call Print_Parastatiko_16
'biks
                    Else
'biks
'                        Call sbInsertFile16
'biks
                    End If
                Case "23"
                    G0Data.Add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                    owner.sbWriteStatusMessage cb.receive_str
            End Select
            
    End Select

End Function

Private Function parse_phase_3()

    Dim G_byte As String
    Dim MN_Bytes As String

    parse_phase_3 = PARSE_OK
    
    G_byte = Mid(cb.receive_str, 1, 1)
    MN_Bytes = Mid(cb.receive_str, 3, 2)

            
    Select Case G_byte
        Case "0"
            parse_phase_3 = PARSE_READ_AGAIN
        Case "1"
            parse_phase_3 = PARSE_ANSWER_REQUIRED
        Case "2"
            parse_phase_3 = PARSE_ANSWER_REQUIRED_DATA
        Case "3"
            parse_phase_3 = PARSE_ANSWER_REQUIRED_DATA
            
'            parse_phase_3 = PARSE_HOST_REJECTION
        Case "4"
            parse_phase_3 = PARSE_CANCEL
        Case "5"
            parse_phase_3 = PARSE_TRANSACTION_COMPLETED
        Case "7"
            If MN_Bytes = "16" Then
               parse_phase_3 = PARSE_READ_AGAIN
            Else
                Select Case Mid(cb.curr_transaction, 1, 1)
                    Case "0", "1", "2", "3"
                        parse_phase_3 = PARSE_SEND_AGAIN
                        
                        cb.send_str = "9999" & cHEAD & cb.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
                        
                        cb.send_str_length = Len(cb.send_str)
                    Case Else
                        parse_phase_3 = PARSE_READ_AGAIN
                End Select
            End If
        Case "8"
            parse_phase_3 = PARSE_SEND_AGAIN
            cb.send_str = "0002" & cHEAD & cb.trn_key & cb.send_str
            cb.send_str_length = Len(cb.send_str)
    End Select

End Function

Public Sub TerminateTransaction(owner As Form, ByRef Send_status As Integer)
    cb.read_again = True
    cb.send_str = "0000" & cHEAD & cb.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
    cb.send_str_length = Len(cb.send_str)
    Send_status = SEND(owner)
    If Send_status <> SEND_OK Then
        cb.read_again = False
'        Call NBG_error("TerminateTransaction", Send_status)
    End If
'    cb.receive_convert = 0
    'Call PrintJrnl("ΟΧΙ", False)
End Sub

Public Sub ContinueTransaction(owner As Form, ByRef Send_status As Integer)
    cb.read_again = True
'    cTRNNum = cTRNNum + 1: UpdateParams
    cb.send_str = "9999" & cHEAD & cb.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
    cb.send_str_length = Len(cb.send_str)
    Send_status = SEND(owner)
    If Send_status <> SEND_OK Then
        cb.read_again = False
'        Call NBG_error("Continue_YES_Click", Send_status)
    End If

'    Call PrintJrnl("ΝΑΙ", False)
End Sub

Public Sub read_sense_code(StrInput As String)  'owner As Form,
    Dim SenseCode As String
    Dim astr As String
    Dim DFH As Integer
    Dim DFHhex As String

    SenseCode = StrPad_(Hex(Asc(Mid(StrInput, 1, 1))) & Hex(Asc(Mid(StrInput, 2, 1))), 4, "0", "L")
    MsgBox SenseCode
    DFHhex = "&H" & StrPad_(Hex(Asc(Mid(StrInput, 3, 1))) & Hex(Asc(Mid(StrInput, 4, 1))), 4, "0", "L")
    DFH = DFHhex
    astr = "SENSE CODE:" & SenseCode & "  DFH:" & Str(DFH)
    'Screen.ActiveForm.sbWriteStatusMessage astr
    MsgBox astr
'    eJournalWrite owner, "Err:" & astr
'    eJournalWrite owner, "ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ"
End Sub

Private Sub AsciiToEbcdic()
Dim InputAscii As String

InputAscii = cb.send_str & Chr$(0)
cb.send_str = GKTranslate(InputAscii, EBCDIC_CP_STRING, ASCII_CP_STRING)
cb.send_str = Left$(cb.send_str, cb.send_str_length)

End Sub

Private Sub EbcdicToAscii()
Dim InputEbcdic As String
Dim iCount As Integer


InputEbcdic = ""
For iCount = 1 To cb.receive_str_length 'replace all chr$(0) in cb.receive_str
    If Mid$(cb.receive_str, iCount, 1) = Chr$(0) Then
        InputEbcdic = InputEbcdic & "."
    Else
        InputEbcdic = InputEbcdic & Mid$(cb.receive_str, iCount, 1)
    End If
Next iCount

InputEbcdic = InputEbcdic & Chr$(0)

cb.receive_str = GKTranslate(InputEbcdic, _
                         ASCII_CP_STRING, _
                         EBCDIC_CP_STRING)

cb.receive_str = Left$(cb.receive_str, cb.receive_str_length)


End Sub

Private Function Comm_printJrn(owner As Form)
'    eJournalWriteAll owner, cb.received_data ', CStr(cTRNCode), cTRNNum
End Function

Public Function HPSSEND_(inputStr As String) As Integer  'owner As Form,
Dim pRetString As String
Dim pConvert As Integer, res As Integer
Dim Bytelist()  As Byte

Dim i As Integer

On Error GoTo JErrorHandler

HPSSEND_ = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
'If Len(inputStr) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function
    
cb.LUADirection = 1: cb.send_convert = 1
cb.RetCode = 1

cb.MsgType = 1
cb.send_str_length = Len(inputStr) + 1
cb.send_str = inputStr & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0
pConvert = 0

pRetString = String$(512, 0)

On Error GoTo CErrorHandler

If cb.send_str_length <= 4096 Then
    pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, cb.send_str_length, _
        cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)

        cb.LUADirection = cb.RetCode

Else
Dim ahead As String, atotal As String, apart As String, firstflag As Boolean
    firstflag = True
    ahead = Left(cb.send_str, 59)
    atotal = Right(cb.send_str, Len(cb.send_str) - 59)
    While atotal <> ""
        If Len(atotal) > 4037 Then
            cb.send_str = Left(ahead, 4) & AsciiToEbcdic_(IIf(firstflag, "F", "M")) & Right(ahead, 54) & Left(atotal, 4037)
            
            atotal = Right(atotal, Len(atotal) - 4037)
            firstflag = False
        Else
            cb.send_str = Left(ahead, 4) & AsciiToEbcdic_("L") & Right(ahead, 54) & atotal
            atotal = ""
        End If
        
        cb.send_str_length = Len(cb.send_str)
        
        pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, cb.send_str_length, _
            cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)
            
        
        cb.LUADirection = cb.RetCode

        res = HPSRECEIVE_() 'owner
        If res <> RECEIVE_OK Then
            HPSSEND_ = SEND_FAILED
            Exit Function
        End If
        cb.RetCode = 1
    Wend
End If


On Error GoTo JErrorHandler
If ResetKey Then
    cb.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    HPSSEND_ = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    'Report_ComError owner
    Exit Function
End If
    
'GenWorkForm.ComTimer.Enabled = False
Exit Function

JErrorHandler:
   MsgBox "SEND_RUNTIME_ERROR J"
'    Call Runtime_error("Journal", Err.Number, Err.Description)
    HPSSEND_ = SEND_RUNTIME_ERROR

CErrorHandler:
   MsgBox "SEND_RUNTIME_ERROR C"
'    Call Runtime_error("Send", Err.Number, Err.Description)
    HPSSEND_ = SEND_RUNTIME_ERROR
    
End Function


Public Function HPSRECEIVE_() As Integer 'owner As Form

On Error GoTo ErrorHandler

Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
    pConvert As Long, pTimeOut As Long, pDebug As Long
Dim pData As String, pData1 As String, pLen As Long, pDataTotal As String, _
    pLengthTotal As Long

HPSRECEIVE_ = RECEIVE_OK
pDataTotal = ""
pLengthTotal = 0


pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0

pTimeOut = cb.TimeOut: pDebug = cb.com_debug

pData = "": pLen = 0

pRetCode = cb.LUADirection
pData = String$(4097, 0)

pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)

If Len(pData) = 4 Then 'Sense Code
    HPSRECEIVE_ = RECEIVE_FAILED
    read_sense_code pData  'owner,
    
    Exit Function
End If

Dim ahead As String, alldata As String, res As Long
alldata = ""
While Len(pData) >= 59
    ahead = Left(pData, 59)
    alldata = alldata & Right(pData, Len(pData) - 59)
    pData = ""

    If EbcdicToAscii_(Mid(ahead, 5, 1)) <> "O" And EbcdicToAscii_(Mid(ahead, 5, 1)) <> "L" Then
        
        
        res = HPSSEND_(Left(ahead, 4) & AsciiToEbcdic_("R") & Right(ahead, 54)) 'owner,
        If res <> SEND_OK Then
            HPSRECEIVE_ = res
            Exit Function
        End If
        
RepeatReceive:
        HPSRECEIVE_ = RECEIVE_OK
        pDataTotal = "": pLengthTotal = 0
        pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0: pTimeOut = cb.TimeOut: pDebug = cb.com_debug: pData = "": pLen = 0
        pRetCode = cb.LUADirection
        pData = String$(4097, 0)
        pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)
        
        If Len(pData) = 4 Then 'Sense Code
            HPSRECEIVE_ = RECEIVE_FAILED
            read_sense_code pData  'owner,
            Exit Function
        End If
        
        If pRet1 <> 0 Or pRet2 <> 0 Then
            HPSRECEIVE_ = RECEIVE_FAILED
            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
            Report_ComError 'owner
            Exit Function
        End If
        If pData = "" Then GoTo RepeatReceive
        
'        HPSRECEIVE_ = RECEIVE_OK
'        pDataTotal = "": pLengthTotal = 0
'        pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0: pTimeOut = cb.TimeOut: pDebug = cb.com_debug: pData = "": pLen = 0
'        'pRetCode = 3
'        pData = String$(4097, 0)
'
'        pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)
'
'        If pRet1 <> 0 Or pRet2 <> 0 Then
'            HPSRECEIVE_ = RECEIVE_FAILED
'            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
'            Report_ComError owner
'            Exit Function
'        End If
    End If
Wend
pData1 = Left$(pData, pLen)

cb.LUADirection = pRetCode

If pRet1 <> 0 Or pRet2 <> 0 Then
    HPSRECEIVE_ = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError 'owner
    Exit Function
End If

If pLen > 0 And pMsgType <> 4 Then
    pDataTotal = pDataTotal & pData1
    pLengthTotal = pLengthTotal + pLen
End If

If pLengthTotal <= 0 Then
    pDataTotal = ""
End If

'cb.receive_str = pDataTotal
'cb.receive_str_length = pLengthTotal
cb.receive_str = alldata
cb.receive_str_length = Len(alldata)

'GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
'    Call Runtime_error("Receive", Err.Number, Err.Description)
    HPSRECEIVE_ = RECEIVE_RUNTIME_ERROR
    
End Function





