Attribute VB_Name = "Communication"
Option Explicit

Private cbcomarea_ctg As cXmlComArea

Private Old_Key As String * 1
Private receivelength As Long

Public Function newReadSenseCode(StrInput As String) As String
    Dim DFH As Integer
    
    newReadSenseCode = StrPad_(Hex(Asc(Mid(StrInput, 1, 1))) & Hex(Asc(Mid(StrInput, 2, 1))), 4, "0", "L")
    DFH = "&H" & StrPad_(Hex(Asc(Mid(StrInput, 3, 1))) & Hex(Asc(Mid(StrInput, 4, 1))), 4, "0", "L")
    newReadSenseCode = "SENSE CODE:" & newReadSenseCode & "  DFH:" & Str(DFH)
End Function

Public Function initialize_cb() As Boolean

Dim OK As Boolean
OK = True
'cb.TimeOut = StrPad_(cTimeOut, 4, "0", "L")  '20 read from file
cb.Ret1 = 0
cb.Ret2 = 0
cb.RetCode = 0
cb.LUADirection = 0

'cb.ApplId = cApplID & Chr$(0)
cb.send_convert = 0
cb.receive_convert = 1
cb.DecodeGreek = 1
cb.encodegreek = 1

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

Dim StartTickCount As Long
Dim EndTickCount As Long
cTRNTime = 0
StartTickCount = GetTickCount

Do   'communicate loop
    communicate = COM_OK
    parse_status = PARSE_READ_AGAIN
    Screen.activeform.sbWriteStatusMessage "ƒ…¡¬…¬¡”« ƒ≈ƒœÃ≈ÕŸÕ. –≈—…Ã≈Õ≈‘≈..."
    Send_status = SEND(owner)
    If Send_status <> SEND_OK Then
        communicate = COM_FAILED
    
        Screen.activeform.sbWriteStatusMessage "À¡»œ” ƒ…¡¬…¬¡”«” !!! " & Str(Send_status)
        eJournalWriteAll owner, "ƒ…¡ œ–«= –—œ¬À«Ã¡ ≈–… œ…ÕŸÕ…¡”  " & Str(Send_status) ', CStr(cTRNCode), cTRNNum
        Exit Function
    End If
    GenWorkForm.vStatus.Panels(1).Text = "¡Õ¡ ‘«”« ƒ≈ƒœÃ≈ÕŸÕ. –≈—…Ã≈Õ≈‘≈..."
    
    Do   'receive loop
        cb.TransTerminating = False
        cb.read_again = False
        cb.receive_str = ""
        cb.receive_str_length = 0
        receive_status = Receive(owner)
        If receive_status <> RECEIVE_OK Then
            communicate = COM_FAILED
            Screen.activeform.sbWriteStatusMessage "À¡»œ” ¡Õ¡ ‘«”«” !!! " & Str(receive_status)
            eJournalWriteAll owner, "À¡»œ” ¡Õ¡ ‘«”«” !!! " & Str(receive_status)  ', CStr(cTRNCode), cTRNNum
            Exit Function
        End If
        EndTickCount = GetTickCount
        cTRNTime = cTRNTime + EndTickCount - StartTickCount
        
        If Mid(cb.receive_str, 5, 3) = "DFH" Then
            Screen.activeform.sbWriteStatusMessage Right(cb.receive_str, cb.receive_str_length - 4) 'Print data to screen
            eJournalWriteAll owner, Right(cb.receive_str, cb.receive_str_length - 4) ', CStr(cTRNCode), cTRNNum
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
                Call read_sense_code(owner, cb.receive_str)
                cb.receive_convert = 1
                cb.BoolTransOk = False
                GoTo TerminateRead
            Case PARSE_BAD_RECEIVED_DATA
                communicate = COM_FAILED
                Screen.activeform.sbWriteStatusMessage "À¡»œ” ”‘œ…◊≈…¡ !!! " & Str(parse_status)
                eJournalWriteAll owner, "À¡»œ” ”‘œ…◊≈…¡ !!! " & Str(parse_status) ', CStr(cTRNCode), cTRNNum
                GoTo TerminateRead
            Case PARSE_OK
                'do nothing
            Case PARSE_CHIEF_TELLER_REQUIRED
                ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False
                If Not isChiefTeller Then
                    'Set SelKeyFrm.Owner = TRNFrm
                    Set SelKeyFrm.owner = owner
                    ChiefRequest = True
                    'SelKeyFrm.Show vbModal, TRNFrm
                    
                    SelKeyFrm.Show vbModal, owner
                    '≈È‰ÈÍ¸Ú ˜ÂÈÒÈÛÏ¸Ú „È· ÙÈÚ Í·Ù·Ë›ÛÂÈÚ ÛıÌ·ÎÎ·„Ï·ÙÔÚ
                    'Ò›ÂÈ Ì· ‹ÒÂÈ ÔÔÛ‰ﬁÔÙÂ ›„ÍÒÈÛÁ
                    While (cTRNCode = 5000 Or cTRNCode = 5001 Or cTRNCode = 5002 _
                    Or cTRNCode = 5100 Or cTRNCode = 5101 Or cTRNCode = 5102) And Not KeyAccepted
                        SelKeyFrm.Show vbModal, owner
                    Wend
                    If Not KeyAccepted Then
                        If cTRNCode <> 5000 And cTRNCode <> 5001 And cTRNCode <> 5002 _
                        And cTRNCode <> 5100 And cTRNCode <> 5101 And cTRNCode <> 5102 Then
                            Call TerminateTransaction(owner, Send_status)
                            communicate = COM_USER_TERMINATED
                        End If
                    Else
                        owner.trn_key = cCHIEFKEY
                    End If
                Else
                    ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                    
                    While (cTRNCode = 5000 Or cTRNCode = 5001 Or cTRNCode = 5002 _
                    Or cTRNCode = 5100 Or cTRNCode = 5101 Or cTRNCode = 5102) And Not KeyAccepted
                        KeyWarning.Show vbModal, owner
                    Wend
                    If Not KeyAccepted Then
                        If cTRNCode <> 5000 And cTRNCode <> 5001 And cTRNCode <> 5002 _
                        And cTRNCode <> 5100 And cTRNCode <> 5101 And cTRNCode <> 5102 Then
                            Call TerminateTransaction(owner, Send_status)
                            communicate = COM_USER_TERMINATED
                        End If
                    Else
                        owner.trn_key = cCHIEFKEY
                    End If
                    
                End If
            Case PARSE_MANAGER_REQUIRED
                ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False
                If Not isManager Then
                    'Set SelKeyFrm.Owner = TRNFrm
                    Set SelKeyFrm.owner = owner
                    ManagerRequest = True
                    'SelKeyFrm.Show vbModal, TRNFrm
                    SelKeyFrm.Show vbModal, owner
                    If Not KeyAccepted Then
                        Call TerminateTransaction(owner, Send_status)
                        communicate = COM_USER_TERMINATED
                    Else
                        owner.trn_key = cTELLERMANAGERKEY
                    End If
                Else
                    ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                    
                    If Not KeyAccepted Then
                        Call TerminateTransaction(owner, Send_status)
                        communicate = COM_USER_TERMINATED
                    Else
                        owner.trn_key = cCHIEFKEY
                    End If
                End If
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
        parse_status = parse_phase_3(owner)
        
        Select Case parse_status
            Case PARSE_READ_AGAIN
                'do nothing - receive loop will be repeated
            Case PARSE_ANSWER_REQUIRED
                'ÂÏˆ‹ÌÈÛÁ ÔË¸ÌÁÚ ÏÂ œ  Í·È Õœ
                If Not owner.SkipCommConfirmation Then
                    Load frmContinue
                    Set frmContinue.aOwner = owner
                    frmContinue.Show 1
                    If ContinueCommunication Then
                        Call ContinueTransaction(owner, Send_status)
                    Else
                        Call TerminateTransaction(owner, Send_status)
                        communicate = COM_USER_TERMINATED
                    End If
                Else
                    Call ContinueTransaction(owner, Send_status)
                End If
            
            Case PARSE_ANSWER_REQUIRED_DATA
                'ÂÏˆ‹ÌÈÛÁ ÔË¸ÌÁÚ ÏÂ œ  Í·È Õœ
                Load frmContinueData
                Set frmContinueData.owner = owner
                frmContinueData.Show 1
                If Not ContinueCommunication Then
                    Call TerminateTransaction(owner, Send_status)
                        communicate = COM_USER_TERMINATED
                End If
            Case PARSE_HOST_REJECTION
                cb.BoolTransOk = False
                WrongFieldLabel = Mid(cb.receive_str, 3, 2)
                
                'Ì· ÒÔÛÙÂËÂÈ Í˘‰ÈÍ·Ú „È· ÙÔÌ ÂÌÙÔÈÛÏÔ ÙÔı Â‰ﬂÔı Ôı Â˜ÂÈ
                'Î·ËÔÚ ÙÈÏÁ, Ì· ·ÌÔÈ„ÂÈ ˆÔÒÏ· Ôı Ë· ‰ÁÏÈÔıÒ„Âﬂ ÙÔ Â‰ﬂÔ ·ıÙ¸Ï·Ù·
                'Í·È Ë· ÛÙ›ÎÌÂÈ ÙÁ Û˘ÛÙﬁ ÙÈÏﬁ Ô˘Ú Á frmcontinuedata ÏÔÌÔ Ôı
                'Ë· Í‹ÌÂÈ format ÙÁÌ ÙÈÏÁ Ò˛Ù·
                'ÒÔ‚ÎÁÏ·: ”ÙÁÌ 1001, 1101, 1231 Ù· Â‰ﬂ· ÔÛ¸ Í·È ı¸ÎÔÈÔ Ë· Ò›ÂÈ Ì· ˆÂı„ÔıÌ ÏÂ
                'Ò¸ÛÁÏÔ
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
                                    "À·ÌË·ÛÏ›ÌÁ ‘ÈÏﬁ –Â‰ﬂÔı " & WrongFieldLabel)
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
                Screen.activeform.sbWriteStatusMessage " À¡»œ” ‘…Ã« –≈ƒ…œ’ " & Mid(cb.receive_str, 3, 2)
                cb.receive_str = " À¡»œ” ‘…Ã« –≈ƒ…œ’ " & Mid(cb.receive_str, 3, 2)
                
            Case PARSE_CANCEL
                cb.BoolTransOk = False
            Case PARSE_TRANSACTION_COMPLETED
                cb.BoolTransOk = True
                Screen.activeform.sbWriteStatusMessage "« ”’Õ¡ÀÀ¡√« œÀœ À«—Ÿ»« ≈."
            Case PARSE_SEND_AGAIN   ' ‘œ –—œ”»≈”¡Ã≈
                 'do nothing
        End Select
        'MsgBox "Parse Phase 3 completed"
    
ParsingEnd:
        If parse_status <> PARSE_READ_AGAIN And cb.inttime <> 0 Then cb.inttime = 0
        GenWorkForm.vStatus.Panels(1).Text = ""
        
        StartTickCount = GetTickCount
    Loop While (parse_status = PARSE_READ_AGAIN Or cb.read_again = True)
Loop While (parse_status = PARSE_SEND_AGAIN)


TerminateRead:
' ”‘œ RECEIVE() √…Õ≈‘¡… ASSIGN H PRETCODE
' ”‘«Õ LUADIRECTION
Do While cb.LUADirection = 3
    cb.receive_str = ""
    cb.receive_str_length = 0
    receive_status = Receive(owner)
    If receive_status <> RECEIVE_OK Then
        communicate = COM_FAILED
'biks
        Screen.activeform.sbWriteStatusMessage "À¡»œ” ¡Õ¡ ‘«”«” !!! " & Str(receive_status)
        eJournalWriteAll owner, "À¡»œ” ¡Õ¡ ‘«”«” !!! " & Str(receive_status) ', CStr(cTRNCode), cTRNNum
'biks
        Exit Function
    End If
Loop

If cb.inttime <> 0 Then
   Call NBG_MsgBox("´ENTERª „È· ÛıÌ›˜ÂÈ·", True, " ")
   cb.inttime = 0
End If

Exit Function

ErrorHandler:
    Call Runtime_error("Communicate", Err.number, Err.description)
    communicate = COM_RUNTIME_ERROR

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
pRetString = VB4SLIDISCONNECT(180, _
                           cb.Ret1, _
                           cb.Ret2, _
                           cb.RetCode, _
                           cb.com_debug)
If cb.RetCode > 0 Then
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
    Call Runtime_error("Disconnect", Err.number, Err.description)
    DISCONNECT_ = DISCONNECT_RUNTIME_ERROR

End Function

Public Function DISCONNECT_ABEND() As Integer
Dim pRetString As String
Dim printer_status As Long

On Error GoTo ErrorHandler

DISCONNECT_ABEND = DISCONNECT_OK
cb.Ret1 = 1 ' 1:„È· abend - 0:„È· Í·ÌÔÌÈÍ¸ disconnect
cb.Ret2 = 0
cb.RetCode = 0

pRetString = String$(512, 0)
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
    Call Runtime_error("Disconnect", Err.number, Err.description)
    DISCONNECT_ABEND = DISCONNECT_RUNTIME_ERROR

End Function

Public Sub Report_ComError(owner)
    
    If Not GenWorkForm.ComTimer.Enabled Then
        ComTimerCounter = 0: ComTimerCycle = 0
        GenWorkForm.ComTimer.interval = 6553
        On Error Resume Next
        
        owner.sbShowCommStatus False
        GenWorkForm.sbShowCommStatus False
        GenWorkForm.ComTimer.Enabled = True
        On Error GoTo 0
    End If

End Sub

Public Function Restore_Connection() As Integer
    Restore_Connection = 1
End Function

Public Function Receive(owner As Form) As Integer

On Error GoTo ErrorHandler

Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
    pConvert As Long, pTimeOut As Long, pDebug As Long
Dim pData As String, pData1 As String, pLen As Long, pDataTotal As String, _
    pLengthTotal As Long


Receive = RECEIVE_OK
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
    Receive = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError owner
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
    ReceivedData.add (cb.receive_str)
    
    If EventLogWrite Then _
    Call EventLog(8, " Translated RECEIVE :" & cb.receive_str)
    If ReceiveJournalWrite Then eJournalWrite "R:" & cb.receive_str
'    If Len(cb.receive_str) > 0 Then _
'        If Left(cb.receive_str, 1) = "0" Then eJournalWrite owner, cb.receive_str
End If
''''''''''''''''''''''''''''''''''''

cb.LUADirection = pRetCode

GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
    Call Runtime_error("Receive", Err.number, Err.description)
    Receive = RECEIVE_RUNTIME_ERROR
    
End Function

Public Function SEND(owner As Form) As Integer
Dim pRetString As String
Dim pConvert As Integer, res As Integer

On Error GoTo JErrorHandler

SEND = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
If Len(cb.send_str) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function
    
cb.RetCode = cb.LUADirection

cb.MsgType = 0
'cb.send_str_length = Len(cb.send_str)
cb.send_str = cb.send_str & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0

''''''''''''''''''''''''''''''''
'Convert ASCII to ABCDIC

If cb.send_convert = 1 Then
    'cb.initsend_str = cb.send_str
    If EventLogWrite Then Call EventLog(8, "Untranslated SEND :" & cb.send_str)
    If SendJournalWrite Then eJournalWrite "S:" & cb.send_str
    If cb.CodePage = UCS_OLD And cb.encodegreek = 1 Then Call Encode_Greek
    cb.send_str = AsciiToEbcdic_(cb.send_str)
    
    'Call AsciiToEbcdic
End If
pConvert = 0
''''''''''''''''''''''''''''''''

pRetString = String$(512, 0)

On Error GoTo CErrorHandler
pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
        cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, cb.com_debug)

cb.LUADirection = cb.RetCode

On Error GoTo JErrorHandler
If ResetKey Then
    owner.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    SEND = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    Report_ComError owner
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
Public Function DecodeGreek(inputStr As String)
    Dim pData As String, _
    HostLetters As String, _
    ASCILetters As String, _
    returnstr As String
    'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh°tyzsuvw" & Chr$(0)
    'ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJLQRSUVW" & Chr$(0)
    HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfg~°tyzsuvw" & Chr$(0)
    ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGJ@LQRSUVW" & Chr$(0)
    pData = inputStr & Chr$(0)
    returnstr = GKTranslate(pData, _
                         HostLetters, _
                         ASCILetters)
    returnstr = Left$(returnstr, Len(inputStr))
    DecodeGreek = returnstr
End Function
Public Function encodegreek(inputStr As String)
    Dim pData As String, _
    HostLetters As String, _
    ASCILetters As String, _
    returnstr As String
    'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh°tyzsuvw" & Chr$(0)
    'ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJLQRSUVW" & Chr$(0)
    HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgG~°tyzsuvwMJLRVF" & Chr$(0)
    ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJ@LQRSUVWNKMPYZ" & Chr$(0)
    pData = inputStr & Chr$(0)
    returnstr = GKTranslate(pData, _
                         ASCILetters, _
                         HostLetters)
    returnstr = Left$(returnstr, Len(inputStr))
    encodegreek = returnstr
End Function


Private Sub Decode_Greek()
Dim pData As String, _
    HostLetters As String, _
    ASCILetters As String
'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh°tyzsuvw" & Chr$(0)
'ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJLQRSUVW" & Chr$(0)
HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfg~°tyzsuvw" & Chr$(0)
ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGJ@LQRSUVW" & Chr$(0)
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
'HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgh°tyzsuvw" & Chr$(0)
'ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJLQRSUVW" & Chr$(0)
HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgG~°tyzsuvwMJLRVF" & Chr$(0)
ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJ@LQRSUVWNKMPYZ" & Chr$(0)
pData = cb.send_str & Chr$(0)
cb.send_str = GKTranslate(pData, _
                         ASCILetters, _
                         HostLetters)
cb.send_str = Left$(cb.send_str, Len(cb.send_str))

End Sub

Public Function Decode_Greek_(inputdata As String) As String
Dim pData As String, HostLetters As String, ASCILetters As String
    HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfg~°tyzsuvw" & Chr$(0)
    ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGJ@LQRSUVW" & Chr$(0)
    pData = inputdata & Chr$(0)
    Decode_Greek_ = GKTranslate(pData, HostLetters, ASCILetters)
    Decode_Greek_ = Left$(Decode_Greek_, Len(inputdata))
End Function

Public Function Encode_Greek_(inputdata As String) As String
Dim pData As String, HostLetters As String, ASCILetters As String
    HostLetters = "ABCDEFGHIJKLMNOPQRSUVWXYZ" & "cdfgG~°tyzsuvwMJLRVF" & Chr$(0)
    ASCILetters = "¡¬√ƒ≈∆«»… ÀÃÕŒœ/–—”‘’÷◊ÿŸ" & "CDFGHJ@LQRSUVWNKMPYZ" & Chr$(0)
    pData = inputdata & Chr$(0)
    Encode_Greek_ = GKTranslate(pData, ASCILetters, HostLetters)
    Encode_Greek_ = Left$(Encode_Greek_, Len(inputdata))
End Function


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
'            Screen.ActiveForm.sbWriteStatusMessage "À¡»œ” ”‘œ…◊≈…¡ !!! "
'            eJournalWriteAll owner, "À¡»œ” ”‘œ…◊≈…¡ !!! " ', CStr(cTRNCode), cTRNNum
            
'            Call NBG_MsgBox("À¡»œ” ”‘œ…◊≈…¡: *" & cb.receive_str & "*", True)
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
            If G_byte = "4" Then Screen.activeform.sbWriteStatusMessage cb.received_data
        Case "2"
            If G_byte = "0" And MF_Byte = "2" Then
                G0Data.add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                Screen.activeform.sbWriteStatusMessage cb.received_data
            Else
                Screen.activeform.sbWriteStatusMessage cb.received_data
                Comm_printJrn owner
            End If
        Case "3"
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
            
            If G_byte = "0" And MF_Byte = "2" Then
                G0Data.add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                Screen.activeform.sbWriteStatusMessage cb.received_data
            Else
                Screen.activeform.sbWriteStatusMessage cb.received_data
                'Comm_printJrn owner
            End If
        Case "4"
            If Not isChiefTeller Then
                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            Else
                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            End If
        Case "5"
            If Not isChiefTeller Then
                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            Else
                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            End If
            
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
        Case "6"
            If Not isChiefTeller Then
                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            Else
                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                If Not KeyAccepted Then parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            End If
'biks
            Screen.activeform.sbWriteStatusMessage cb.received_data
            eJournalWriteAll owner, cb.received_data ', CStr(cTRNCode), cTRNNum
'biks
        Case "7"
            If Not isChiefTeller Then
                parse_phase_1 = PARSE_CHIEF_TELLER_REQUIRED
            Else
                ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                If KeyAccepted Then
                    ResetKey = True
                    Old_Key = owner.trn_key
                    owner.trn_key = cCHIEFKEY
                End If
            End If
            
            
            'print_to_journal (received_data)
            If Len(cb.received_data) > 1 Then Comm_printJrn owner
'biks
            Screen.activeform.sbWriteStatusMessage cb.received_data
            'Comm_printJrn owner
'biks
        Case "8", "9", "A", "B", "C", "D", "E", "F", "¡", "¬", "√", "ƒ", "≈", "∆"
            If Not isManager Then
                parse_phase_1 = PARSE_MANAGER_REQUIRED
            Else
                ManagerRequest = True: ChiefRequest = False: Load KeyWarning: Set KeyWarning.owner = owner: KeyWarning.Show vbModal, owner
                If KeyAccepted Then
                    ResetKey = True
                    Old_Key = owner.trn_key
                    owner.trn_key = cTELLERMANAGERKEY
                End If
            End If

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
                    ' ¡   ¡ ‘ ¡ ◊ Ÿ — … ” ‘ ≈ ”    ≈√√—¡÷≈”
                    'OK = Print_PASSBOOK()
                Case "02"
                    '≈ ‘’–Ÿ”« ”‘œ ¬…¬À…¡—…œ, ’–œƒœ◊«
                    ' Â·Ì mid(cb.received_data,5,1) = "1" eject ‚È‚ÎÈ‹ÒÈÔ
                Case "03"
                    ' ”’√ —…”« –œ”œ’  ƒ≈Õ ◊—«”…Ãœ–œ…≈…‘¡…
                    ' ”’√ —…”« ’–œÀœ…–œ’ √…¡ PASSWORD –—œ⁄”‘¡Ã≈Õœ’
'                    If cb.curr_transaction = "1131" Then
'                        If Val(Mid(cb.receive_str, 6, 12)) > cb.posolimit Then
'                            Do
'                                cb.TransTerminating = False
''biks
''                                Call ChangeChiefTellerState(frm1131)
''biks
'                            Loop While cb.TransTerminating = True
'                        End If
'                    End If
                Case "04"
                    eJournalWriteAll owner, cPOSTDATE & " «Ã/Õ…¡ ‘≈—/ œ’" ', CStr(cTRNCode), cTRNNum)
                Case "05"
                    '–¡…—ÕŸ ¡–œ ‘œ  Ã ‘œ STRING √…¡ ‘«Õ 0620
                    cHEAD = Mid(cb.receive_str, Len(cb.receive_str) - 4, 4)
                    UpdatexmlEnvironment "SESSIONCD", CStr(cHEAD)
                Case "07"
                    'ƒ«Ã…œ’—√Ÿ ‘œ STRING
                    cb.send_str = StrPad_(CStr(cTRNNum), 3, "0", "L") & cNextDateFlag & "00000000000000000000"
                Case "09"
                    '·ËÒÔÈÛÙÁÚ Í·È ÔÛ¸ ·¸ HOST
'BIKS
'                    cb.stringfor3130 = cb.received_data
'                    Call PrintJrnl(Mid(cb.received_data, 1, 1) & " " & _
                            StrPad(format_num(Val(Mid(cb.received_data, 2, 12))), 17) & _
                            IIf(Mid(cb.received_data, 13, 1) = "+", "–◊", "––"), False)
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
                    G0Data.add Mid(cb.receive_str, 6, Len(cb.receive_str) - 1)
                    owner.sbWriteStatusMessage cb.receive_str
            End Select
            
    End Select

End Function

Private Function parse_phase_3(owner As Form)

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
                        
                        cb.send_str = "9999" & cHEAD & owner.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
                        
                        'cb.send_str_length = Len(cb.send_str)
                    Case Else
                        parse_phase_3 = PARSE_READ_AGAIN
                End Select
            End If
        Case "8"
            parse_phase_3 = PARSE_SEND_AGAIN
            cb.send_str = "0002" & cHEAD & owner.trn_key & cb.send_str
            'cb.send_str_length = Len(cb.send_str)
    End Select

End Function

Public Sub TerminateTransaction(owner As Form, ByRef Send_status As Integer)
    cb.read_again = True
    cb.send_str = "0000" & cHEAD & owner.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
    'cb.send_str_length = Len(cb.send_str)
    Send_status = SEND(owner)
    If Send_status <> SEND_OK Then
        cb.read_again = False
        Call NBG_error("TerminateTransaction", Send_status)
    End If
'    cb.receive_convert = 0
    'Call PrintJrnl("œ◊…", False)
End Sub

Public Sub ContinueTransaction(owner As Form, ByRef Send_status As Integer)
    cb.read_again = True
    cb.send_str = "9999" & cHEAD & owner.trn_key & StrPad_(CStr(cTRNNum), 3, "0", "L")
    Send_status = SEND(owner)
    If Send_status <> SEND_OK Then
        cb.read_again = False
        Call NBG_error("Continue_YES_Click", Send_status)
    End If
End Sub

Public Sub read_sense_code(owner As Form, StrInput As String)
    SenseCodeMessage = DecodeSenseCode(StrInput)
    If Not (owner Is Nothing) Then owner.sbWriteStatusMessage SenseCodeMessage
    eJournalWrite "Err:" & SenseCodeMessage & "ƒ…¡ œ–« ”’Õ¡ÀÀ¡√«”"
End Sub

'Private Sub AsciiToEbcdic()
'Dim InputAscii As String

'InputAscii = cb.send_str & Chr$(0)
'cb.send_str = GKTranslate(InputAscii, EBCDIC_CP_STRING, ASCII_CP_STRING)
'cb.send_str = Left$(cb.send_str, Len(cb.send_str))
'End Sub


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
    eJournalWriteAll owner, cb.received_data ', CStr(cTRNCode), cTRNNum
End Function

Public Function HPSSEND_(owner As Form, inputStr As String) As Integer
Dim pRetString As String
Dim pConvert As Integer, res As Integer
Dim Bytelist()  As Byte

Dim i As Integer

On Error GoTo JErrorHandler

HPSSEND_ = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
If Len(inputStr) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function

cb.LUADirection = 1: cb.send_convert = 1
cb.RetCode = 1

cb.MsgType = 1
'cb.send_str_length = Len(inputStr) + 1
cb.send_str = inputStr & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0
pConvert = 0

pRetString = String$(512, 0)

On Error GoTo CErrorHandler

If Len(cb.send_str) <= 4096 Then

    pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
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

        'cb.send_str_length = Len(cb.send_str)

        pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
            cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)


        cb.LUADirection = cb.RetCode

        res = HPSRECEIVE_(owner)
        If res <> RECEIVE_OK Then
            HPSSEND_ = SEND_FAILED
            Exit Function
        End If
        cb.RetCode = 1
    Wend
End If


On Error GoTo JErrorHandler
If ResetKey Then
    owner.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    HPSSEND_ = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    Report_ComError owner
    Exit Function
End If

GenWorkForm.ComTimer.Enabled = False
Exit Function

JErrorHandler:
    Call Runtime_error("Journal", Err.number, Err.description)
    HPSSEND_ = SEND_RUNTIME_ERROR

CErrorHandler:
    Call Runtime_error("Send", Err.number, Err.description)
    HPSSEND_ = SEND_RUNTIME_ERROR

End Function


Public Function HPSRECEIVE_(owner As Form) As Integer

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
    read_sense_code owner, pData
    Exit Function
End If

Dim ahead As String, alldata As String, res As Long
alldata = ""
While Len(pData) >= 59
    ahead = Left(pData, 59)
    alldata = alldata & Right(pData, Len(pData) - 59)
    pData = ""

    If EbcdicToAscii_(Mid(ahead, 5, 1)) <> "O" And EbcdicToAscii_(Mid(ahead, 5, 1)) <> "L" Then


        res = HPSSEND_(owner, Left(ahead, 4) & AsciiToEbcdic_("R") & Right(ahead, 54))
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
            read_sense_code owner, pData
            Exit Function
        End If

        If pRet1 <> 0 Or pRet2 <> 0 Then
            HPSRECEIVE_ = RECEIVE_FAILED
            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
            Report_ComError owner
            Exit Function
        End If
        If pData = "" Then GoTo RepeatReceive

    End If
Wend
pData1 = Left$(pData, pLen)

cb.LUADirection = pRetCode

If pRet1 <> 0 Or pRet2 <> 0 Then
    HPSRECEIVE_ = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError owner
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

GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
    Call Runtime_error("Receive", Err.number, Err.description)
    HPSRECEIVE_ = RECEIVE_RUNTIME_ERROR

End Function

Public Function IRISSEND_(owner As Form, inputStr As String) As Integer
Dim pRetString As String
Dim pConvert As Integer, res As Integer
Dim Bytelist()  As Byte

Dim i As Integer

If LogIrisCom Then
    Dim logfilename As String
    If tmpSendViewName = "" Then logfilename = "inputview.txt" Else logfilename = tmpSendViewName & ".txt"
    
'    On Error Resume Next
    SaveTextFile logfilename, inputStr
'
'    Close #3
'
'    On Error GoTo 0
'    'Open "c:\" & logfilename For Output As #3
'
'    Open App.path & "\" & logfilename For Output As #3
'
'    Print #3, inputStr
'
'    Close #3
    
End If
On Error GoTo JErrorHandler

IRISSEND_ = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
If Len(inputStr) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function

cb.LUADirection = 1: cb.send_convert = 1
cb.RetCode = 1

cb.MsgType = 1
'cb.send_str_length = Len(inputStr) + 1
cb.send_str = inputStr & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0
pConvert = 0

pRetString = String$(512, 0)

On Error GoTo CErrorHandler

If Len(cb.send_str) <= IRIS_MAX_RU_SIZE Then

    pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
        cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)

        cb.LUADirection = cb.RetCode
Else

Dim ahead As String, atotal As String, apart As String, firstflag As Boolean
    firstflag = True
    ahead = Left(cb.send_str, IRIS_OFFSET)
    atotal = Right(cb.send_str, Len(cb.send_str) - IRIS_OFFSET)
    While atotal <> ""
    
        If Len(atotal) > IRIS_MAX_RU_SIZE - IRIS_OFFSET Then
            cb.send_str = Left(ahead, 4) & AsciiToEbcdic_(IIf(firstflag, "F", "M")) & Right(ahead, IRIS_OFFSET - 5) & Left(atotal, IRIS_MAX_RU_SIZE - IRIS_OFFSET)
            
            atotal = Right(atotal, Len(atotal) - (IRIS_MAX_RU_SIZE - IRIS_OFFSET))
            firstflag = False
        Else
            cb.send_str = Left(ahead, 4) & AsciiToEbcdic_("L") & Right(ahead, IRIS_OFFSET - 5) & atotal
            atotal = ""
        End If
        
        'cb.send_str_length = Len(cb.send_str)
    
        pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
            cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)
            
        
        cb.LUADirection = cb.RetCode

        res = IRISRECEIVE_(owner, EbcdicToAscii_(Mid(ahead, 60, 8)), 60)

        If res <> RECEIVE_OK Then
            IRISSEND_ = SEND_FAILED
            Exit Function
        End If
        cb.RetCode = 1
    Wend
End If


On Error GoTo JErrorHandler
If ResetKey Then
    owner.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then

    IRISSEND_ = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    Report_ComError owner
    Exit Function
End If

GenWorkForm.ComTimer.Enabled = False
Exit Function

JErrorHandler:
    Call Runtime_error("Journal", Err.number, Err.description)
    IRISSEND_ = SEND_RUNTIME_ERROR

CErrorHandler:
    Call Runtime_error("Send", Err.number, Err.description)
    IRISSEND_ = SEND_RUNTIME_ERROR
    
End Function


Public Function IRISRECEIVE_(owner As Form, OutputName As String, outputPos As Integer) As Integer

On Error GoTo ErrorHandler

Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
    pConvert As Long, pTimeOut As Long, pDebug As Long
Dim pData As String, pData1 As String, pLen As Long, pDataTotal As String, _
    pLengthTotal As Long


IRISRECEIVE_ = RECEIVE_OK
pDataTotal = ""
pLengthTotal = 0


pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0

pTimeOut = cb.TimeOut: pDebug = cb.com_debug

pData = "": pLen = 0

pRetCode = cb.LUADirection
pData = String$(IRIS_MAX_RU_SIZE + 1, 0)

pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)

If Len(pData) = 4 Then 'Sense Code
    IRISRECEIVE_ = RECEIVE_FAILED
    read_sense_code owner, pData
    Exit Function
End If

Dim ahead As String, alldata As String, res As Long
alldata = ""
While Len(pData) >= IRIS_OFFSET
    ahead = Left(pData, IRIS_OFFSET)
    If EbcdicToAscii_(Mid(ahead, outputPos, 8)) <> OutputName Then
        alldata = alldata & Right(pData, Len(pData) - IRIS_OFFSET)
        cb.receive_str = alldata
        cb.receive_str_length = Len(alldata)

        IRISRECEIVE_ = RECEIVE_AUTH_FAILED
        If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
        Exit Function
    End If
    
    alldata = alldata & Right(pData, Len(pData) - IRIS_OFFSET)
    If LogIrisCom Then
        Dim logfilename As String
        If tmpReceiveViewName = "" Then logfilename = "outputview.txt" Else logfilename = tmpReceiveViewName & ".txt"
        'Open "c:\" & logfilename For Output As #3
        Open NetworkHomeDir & "\" & logfilename For Output As #3
        Print #3, ahead + alldata
        Close #3
    End If
    
    pData = ""

    If EbcdicToAscii_(Mid(ahead, 5, 1)) <> "O" And EbcdicToAscii_(Mid(ahead, 5, 1)) <> "L" Then
        
        
        res = HPSSEND_(owner, Left(ahead, 4) & AsciiToEbcdic_("R") & Right(ahead, IRIS_OFFSET - 5))
        If res <> SEND_OK Then
            IRISRECEIVE_ = res
            Exit Function
        End If
        
RepeatReceive:
        IRISRECEIVE_ = RECEIVE_OK
        pDataTotal = "": pLengthTotal = 0
        pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0: pTimeOut = cb.TimeOut: pDebug = cb.com_debug: pData = "": pLen = 0
        pRetCode = cb.LUADirection
        pData = String$(4097, 0)
        pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)
        
        If Len(pData) = 4 Then 'Sense Code
            IRISRECEIVE_ = RECEIVE_FAILED
            read_sense_code owner, pData
            Exit Function
        End If
        
        If pRet1 <> 0 Or pRet2 <> 0 Then
            IRISRECEIVE_ = RECEIVE_FAILED
            If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
            Report_ComError owner
            Exit Function
        End If
        If pData = "" Then GoTo RepeatReceive
        
    End If
Wend
pData1 = Left$(pData, pLen)


'RECEIVE_AUTH_FAILED
cb.LUADirection = pRetCode

If pRet1 <> 0 Or pRet2 <> 0 Then
    IRISRECEIVE_ = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError owner
    Exit Function
End If

If pLen > 0 And pMsgType <> 4 Then
    pDataTotal = pDataTotal & pData1
    pLengthTotal = pLengthTotal + pLen
End If

If pLengthTotal <= 0 Then
    pDataTotal = ""
End If

cb.receive_str = alldata
cb.receive_str_length = Len(alldata)

GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
    Call Runtime_error("Receive", Err.number, Err.description)
    IRISRECEIVE_ = RECEIVE_RUNTIME_ERROR
    
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
Public Function onlineSEND_(owner As Form, inputStr As String) As Integer
Dim pRetString As String
Dim pConvert As Integer, res As Integer
Dim Bytelist()  As Byte

Dim i As Integer

On Error GoTo JErrorHandler

onlineSEND_ = SEND_OK
cb.Ret1 = 0
cb.Ret2 = 0
If Len(inputStr) < 1 Then Call NBG_MsgBox("No data to send!!!", True): Exit Function
    
cb.LUADirection = 1: cb.send_convert = 1
cb.RetCode = 1

cb.MsgType = 1
'cb.send_str_length = Len(inputStr)
cb.send_str = inputStr & Chr$(0)
cb.receive_str = ""
cb.receive_str_length = 0
pConvert = 0

pRetString = String$(512, 0)

On Error GoTo CErrorHandler

If Len(cb.send_str) <= online_MAX_RU_SIZE Then
    
    pRetString = VB4SLISEND(cb.send_str, pConvert, cb.TimeOut, Len(cb.send_str), _
        cb.MsgType, cb.Ret1, cb.Ret2, cb.RetCode, 0)

        cb.LUADirection = cb.RetCode
Else

End If


On Error GoTo JErrorHandler
If ResetKey Then
    owner.trn_key = Old_Key
    ResetKey = False
End If

If cb.Ret1 <> 0 Or cb.Ret2 <> 0 Then
    onlineSEND_ = SEND_FAILED
    If EventLogWrite Then Call EventLog(1, "VB Application :SEND FAILED")
    Report_ComError owner
    Exit Function
End If
    
GenWorkForm.ComTimer.Enabled = False
Exit Function

JErrorHandler:
    Call Runtime_error("Journal", Err.number, Err.description)
    onlineSEND_ = SEND_RUNTIME_ERROR

CErrorHandler:
    Call Runtime_error("Send", Err.number, Err.description)
    onlineSEND_ = SEND_RUNTIME_ERROR
    
End Function

Public Function onlineRECEIVE_(owner As Form) As Integer

On Error GoTo ErrorHandler

Dim pMsgType As Long, pRet1 As Long, pRet2 As Long, pRetCode As Long, _
    pConvert As Long, pTimeOut As Long, pDebug As Long
Dim pData As String, pData1 As String, pLen As Long, pDataTotal As String, _
    pLengthTotal As Long


onlineRECEIVE_ = RECEIVE_OK
pDataTotal = ""
pLengthTotal = 0


pMsgType = 0: pRet1 = 0: pRet2 = 0: pRetCode = 0: pConvert = 0

pTimeOut = cb.TimeOut: pDebug = cb.com_debug

pData = "": pLen = 0

pRetCode = cb.LUADirection
pData = String$(online_MAX_RU_SIZE + 1, 0)

pData = VB4SLIRECEIVE(pConvert, pTimeOut, pLen, pMsgType, pRet1, pRet2, pRetCode, pDebug)

If Len(pData) = 4 Then 'Sense Code
    onlineRECEIVE_ = RECEIVE_FAILED
    read_sense_code owner, pData
    Exit Function
End If


'RECEIVE_AUTH_FAILED
cb.LUADirection = pRetCode

If pRet1 <> 0 Or pRet2 <> 0 Then
    onlineRECEIVE_ = RECEIVE_FAILED
    If EventLogWrite Then EventLog 1, "VB Application :RECEIVE FAILED"
    Report_ComError owner
    Exit Function
End If

If pLen > 0 And pMsgType <> 4 Then
    pDataTotal = pData
    pLengthTotal = pLen
End If

If pLengthTotal <= 0 Then
    pDataTotal = ""
End If

cb.receive_str = pData
cb.receive_str_length = Len(pData)

GenWorkForm.ComTimer.Enabled = False

Exit Function

ErrorHandler:
    Call Runtime_error("Receive", Err.number, Err.description)
    onlineRECEIVE_ = RECEIVE_RUNTIME_ERROR
    
End Function

Public Function IRISCom_(OwnerForm As Form, Trn As String, rule As String, InputView, OutputView, Optional AuthUser, _
    Optional Appltran, Optional ErrorView, Optional ErrorCount, Optional UpdateTrnNumFlag As Boolean, Optional WriteJournalFlag As Boolean) As Integer
Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
Dim IRISAuthError As String
    If Not Flag610 Then
        LogMsgbox "ƒÂÌ ›˜ÂÈ „ﬂÌÂÈ Û˝Ì‰ÂÛÁ (0610)", vbCritical: IRISCom_ = 999: Exit Function
    End If
        
    Dim aresult As cSNAResult
    If Not (Screen.activeform Is Nothing) Then
       If Not (TypeOf Screen.activeform Is SelectTRNFrm) Then
          Screen.activeform.sbWriteStatusMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ..."
       End If
    End If
    
    Set aresult = iriscomnew_(OwnerForm, Left(Trn & "    ", 4), rule, InputView, OutputView, AuthUser, _
        Appltran, ErrorView, ErrorCount, UpdateTrnNumFlag)
    
    IRISCom_ = aresult.ErrCode
    aresult.UpdateForm OwnerForm
    If Not (Screen.activeform Is Nothing) Then
        If Not (TypeOf Screen.activeform Is SelectTRNFrm) Then
           If aresult.ErrCode <> 0 Then
                LogMsgbox "À‹ËÔÚ: " & CStr(aresult.ErrCode) & " " & aresult.ErrMessage, vbCritical, "–Ò¸‚ÎÁÏ· ≈ÈÍÔÈÌ˘Ìﬂ·Ú...."
           Else
                Screen.activeform.sbWriteStatusMessage ""
           End If
        End If
    End If
    Exit Function

ERROR_LINE:
    IRISCom_ = 999
    NBG_LOG_MsgBox "¡›Ùı˜Â Á ÂÈÍÔÈÌ˘Ìﬂ·: " & Err.source & ": " & Err.number & ", " & Err.description, , "View: " & InputView.name
    If Not (OwnerForm Is Nothing) Then OwnerForm.Enabled = True
End Function

Public Function IRISComLocal_(OwnerForm As Form, Trn As String, rule As String, InputView, OutputView, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount, Optional InputDataFile As String, Optional OutputDataFile As String) As Integer
Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
Dim IRISAuthError As String

    IRISAuthError = ""

    Dim anode As MSXML2.IXMLDOMElement
    If Not (xmlIRISRules.documentElement Is Nothing) Then
        If Not xmlIRISRulesUpdate Is Nothing Then
            If Not xmlIRISRulesUpdate.documentElement Is Nothing Then
                Set anode = xmlIRISRulesUpdate.documentElement.selectSingleNode(rule)
            End If
        End If
        If anode Is Nothing Then Set anode = xmlIRISRules.documentElement.selectSingleNode(rule)
        If Not (anode Is Nothing) Then Trn = anode.Attributes(0).Text
    End If

    'InputDataFile parameter indicates the name of the file to be used in order to set the InputView.data value when working offline
    'Missing InputDataFile parameter indicates that the name of the file to be used results in the actual input view name
    'Zero lenght string InputDataFile parameter indicates that the InputView.data is not to be reset
    If IsMissing(InputDataFile) Then InputDataFile = InputView.name
    If InputDataFile = "" Then
    Else
        Open App.path + "\" + InputDataFile + ".txt" For Binary Access Read As #4
        Dim cbSend
        cbSend = InputB(LOF(4), 4)
        Close #4
        If cbSend = "" Then
            InputView.Clear
        Else
            InputView.Data = StrConv(cbSend, vbUnicode)
        End If
    End If


    If IsMissing(Appltran) Then Appltran = InputView.v2Value("COD_TX")
    If IsMissing(AuthUser) Then AuthUser = UCase(cIRISUserName)
    If Appltran = "" Then Appltran = InputView.v2Value("COD_TX")
    If AuthUser = "" Then AuthUser = UCase(cIRISUserName)

    Appltran = Left(Appltran & String(8, " "), 8)
    AuthUser = UCase(Left(AuthUser & String(8, " "), 8))
    astr = Left(UCase(cIRISComputerName) & "         ", 9) & Left(UCase(cIRISUserName) & "         ", 8) & "00001" & Right("00000" & CStr(IRIS_MAX_RU_SIZE), 5) & "11"
    rule = Left(rule & "        ", 8)
    InputName = Left(InputView.StructID & "        ", 8)
    OutputName = Left(OutputView.StructID & "        ", 8)
    If InputView.length + IRIS_OFFSET > IRIS_MAX_RU_SIZE Then
        Trn = Left(Trn, 4) & "F"
    Else
        Trn = Left(Trn, 4) & "O"
    End If

    astr = AsciiToEbcdic_(Trn & astr) & _
            AsciiToEbcdic_(CStr(AuthUser)) & AsciiToEbcdic_(CStr(Appltran)) & _
            AsciiToEbcdic_(rule) & AsciiToEbcdic_(InputName) & AsciiToEbcdic_(OutputName) & _
            IntToHps_(Len(InputView.Data)) & _
            IntToHps_(Len(OutputView.Data))

    aSize = Len(InputView.Data)
    If Not (OwnerForm Is Nothing) Then OwnerForm.sbWriteStatusMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ. –ÂÒÈÏ›ÌÂÙÂ...":
    DoEvents
    astr = astr & InputView.Data

    UpdateTrnNum_
    If Not (OwnerForm Is Nothing) Then eJournalWrite "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ " & cTRNCode

    If Not (OwnerForm Is Nothing) Then OwnerForm.Enabled = False
    If LogIrisCom Then
        Open NetworkHomeDir & "\" & "view1.txt" For Output As #1
        Print #1, astr
        Close #1
    End If
    Dim StartTime
    StartTime = Time
    Dim StartTickCount
    StartTickCount = GetTickCount
    If Not (OwnerForm Is Nothing) Then OwnerForm.sbWriteStatusMessage "¡Õ¡ ‘«”« ”‘œ…◊≈…ŸÕ. –ÂÒÈÏ›ÌÂÙÂ...":
    DoEvents
    
    Dim EndTime
    Dim EndTickCount
    EndTime = Time
    EndTickCount = GetTickCount
    If Not (OwnerForm Is Nothing) Then OwnerForm.sbWriteStatusMessage ""
    
    'OutputDataFile parameter indicates the name of the file to be used in order to set the OutputView.data value when working offline
    'Missing or zero lenght string OutputDataFile parameter indicates that the name of the file to be used results in the actual output view name
    If IsMissing(OutputDataFile) Or OutputDataFile = "" Then OutputDataFile = OutputView.name
    Open App.path + "\" + OutputDataFile + ".txt" For Binary Access Read As #4
    Dim cbReceive
    cbReceive = InputB(LOF(4), 4)
    Close #4
    If cbReceive = "" Then
        OutputView.Data = OutputView.Clear
        IRISComLocal_ = 999
        CommunicationStarted = False
        Exit Function
    Else
        OutputView.Data = StrConv(cbReceive, vbUnicode)
        IRISComLocal_ = 0
    End If
    
    If Not (OwnerForm Is Nothing) Then OwnerForm.Enabled = True

End Function

Public Function iriscomnewCTG_(OwnerForm As Form, Trn As String, rule As String, InputView, OutputView, _
    Optional AuthUser, Optional Appltran) As cSNAResult
    
    If Not Flag610 Then
        eJournalWrite "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ " & cTRNCode
        Set iriscomnewCTG_ = New cSNAResult
        iriscomnewCTG_.ErrCode = GENERIC_COM_ERROR
        iriscomnewCTG_.ErrMessage = "ƒÂÌ ›˜ÂÈ „ﬂÌÂÈ Û˝Ì‰ÂÛÁ (0610)"
        Exit Function
    End If

    OutputView.ClearData
    Dim caContainer As New Buffers
    Dim adoc As MSXML2.DOMDocument30
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If adoc Is Nothing Then Exit Function
    Dim Node As IXMLDOMElement
    Dim ComArea As cXmlComArea
    Dim datanode As IXMLDOMElement
    Dim datadoc As MSXML2.DOMDocument30
    Dim Result As String
    Dim cod_tx As String
    
    Dim anode As MSXML2.IXMLDOMElement
    If Not (xmlIRISRules.documentElement Is Nothing) Then
        If Not xmlIRISRulesUpdate Is Nothing Then
            If Not xmlIRISRulesUpdate.documentElement Is Nothing Then
                Set anode = xmlIRISRulesUpdate.documentElement.selectSingleNode(rule)
            End If
        End If
        If anode Is Nothing And rule <> "" Then Set anode = xmlIRISRules.documentElement.selectSingleNode(rule)
        If Not (anode Is Nothing) Then
            Trn = anode.Attributes(0).Text

            Dim codtxattr As IXMLDOMAttribute
            Set codtxattr = anode.Attributes.getNamedItem("CODTX")
            If Not codtxattr Is Nothing Then cod_tx = codtxattr.nodeValue
        End If
    End If
    If IsMissing(AuthUser) Then AuthUser = ""
    AuthUser = UCase(AuthUser)
    Dim authSTD_TRN_I_PARM_V
    If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
        Set authSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
        If Trim(AuthUser) = "" Then
            AuthUser = UCase(authSTD_TRN_I_PARM_V.ByName("ID_EMPL_AUT", 1).Value)
        Else
            authSTD_TRN_I_PARM_V.ByName("ID_EMPL_AUT", 1).Value = AuthUser
        End If
    End If
    If cod_tx <> "" Then
        Dim iSTD_TRN_I_PARM_V, iCUF_USR_ACCESS_D, iSTD_APPL_PARM_V

        If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
            Set iSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
            iSTD_TRN_I_PARM_V.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("STD_TRN_I_PARM_V") = iSTD_TRN_I_PARM_V.Data
        End If
        If Not InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing Then
            Set iSTD_APPL_PARM_V = InputView.ByName("STD_APPL_PARM_V")
            iSTD_APPL_PARM_V.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("STD_APPL_PARM_V") = iSTD_APPL_PARM_V.Data
        End If
        If Not InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
            Set iCUF_USR_ACCESS_D = InputView.ByName("CUF_USR_ACCESS_D")
            iCUF_USR_ACCESS_D.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("CUF_USR_ACCESS_D") = iCUF_USR_ACCESS_D.Data
        End If
        If InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing And _
           InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing And _
           InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
           InputView.v2Value("COD_TX") = cod_tx
        End If
    End If
    
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='HPSHEADER_CTG']", "HPSHEADER_CTG", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If Node Is Nothing Then Exit Function
    Set ComArea = New cXmlComArea
    Set ComArea.content = Node
    
    Set ComArea.Container = caContainer
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "HPSHEADER_CTG", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datanode Is Nothing Then Exit Function
    Set datadoc = XmlLoadString(datanode.XML, "DataDoc", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datadoc Is Nothing Then Exit Function
    Result = ComArea.LoadXML(datadoc.XML)
    
    With ComArea.Buffer
        Dim authcd As String
        If cod_tx = "" Then
            Dim aiSTD_TRN_I_PARM_V, aiCUF_USR_ACCESS_D, aiSTD_APPL_PARM_V
    
            If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
                Set aiSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
                authcd = aiSTD_TRN_I_PARM_V.ByName("COD_TX", 1).Value
            End If
            If Not InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing Then
                Set aiSTD_APPL_PARM_V = InputView.ByName("STD_APPL_PARM_V")
                authcd = aiSTD_APPL_PARM_V.ByName("COD_TX", 1).Value
            End If
            If Not InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
                Set aiCUF_USR_ACCESS_D = InputView.ByName("CUF_USR_ACCESS_D")
                authcd = aiCUF_USR_ACCESS_D.ByName("COD_TX", 1).Value
            End If
        Else
            authcd = cod_tx
        End If
        
        If IsMissing(Appltran) Then Appltran = authcd
        If Trim(Appltran) = "" Then Appltran = authcd
        
        If cod_tx <> "" Then Appltran = cod_tx
        
        If Trim(Appltran) = "" Then Appltran = "ZCBB0CAC"
    
        .v2Value("TRANID") = "XXXX"
        If Not IsMissing(Appltran) Then .v2Value("TRANID") = Trn
        .v2Value("REQ_TYPE") = "HPS"
        .v2Value("APPL_PGM") = UCase(rule)
        
        .v2Value("USER_ID") = UCase(cIRISUserName)
        If Trn = "IRST" Or Trn = "LEYY" Then
            .v2Value("WS_ID") = Left(UCase(cIRISComputerName) & String(10, " "), 10) & Left(UCase(Encode_Greek_(cTERMINALID)) & String(5, " "), 3)
            .v2Value("VERSION") = Right(Left(UCase(Encode_Greek_(cTERMINALID)) & String(5, " "), 5), 2)
        Else
            .v2Value("WS_ID") = Left(UCase(cIRISComputerName) & String(10, " "), 10)
        End If
        .v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
        
        .v2Value("TRAN_SCD") = cHEAD
        If Trim(.v2Value("TRAN_KEY")) = "" Then .v2Value("TRAN_KEY") = "TELLER"
        If Not IsMissing(AuthUser) Then
            .v2Value("AUTH_USER") = UCase(AuthUser)
        End If
        .v2Value("AUTH_TRANS") = UCase(Appltran)
        
        .v2Value("IDFLEN") = InputView.length
        .v2Value("ODFLEN") = OutputView.length
        
'        .v2Value("ATERM_ID") = cTERMINALID
'        .v2Value("BRANCH") = UCase(cBRANCH)
    
    End With

Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
Dim onlineAuthError As String
    ComArea.Buffer.GetXMLView
    astr = ComArea.Buffer.Data & InputView.Data & OutputView.Data
    
    eJournalWrite "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ " & cTRNCode
    
    Dim connector As New cCTGConnection
    
    connector.OpClass = "HPS_CTG"
    Dim adescription As String
        
    adescription = InputView.name
    adescription = InputView.BuffType
    If Len(adescription) > 2 Then
        If Right(adescription, 2) = "_I" Then adescription = Left(adescription, Len(adescription) - 2)
    End If
        
    connector.OpCode = Trn
    connector.OpDescription = adescription
    If Not IsMissing(AuthUser) Then connector.AuthUser = Trim(AuthUser)
    
    Set iriscomnewCTG_ = connector.SimpleExec(astr)
    If (iriscomnewCTG_.ErrCode = 0 Or iriscomnewCTG_.ErrCode = COM_OK) And iriscomnewCTG_.SenseCodeMessage = "" Then
        ComArea.Buffer.Data = Left(connector.ReceiveData, Len(ComArea.Buffer.Data))
        If ComArea.Buffer.v2Value("RC") <> 0 Then
            iriscomnewCTG_.ErrCode = ComArea.Buffer.v2Value("RC")
            iriscomnewCTG_.ErrMessage = ComArea.Buffer.v2Value("TRANID") & _
                "(" & ComArea.Buffer.v2Value("RE_MSG") & _
                " ÛÙÔ " & ComArea.Buffer.v2Value("RC_PGM") & ") " & ComArea.Buffer.v2Value("RC_TXT")
        End If
    End If
    If (iriscomnewCTG_.ErrCode = 0 Or iriscomnewCTG_.ErrCode = COM_OK) And iriscomnewCTG_.SenseCodeMessage = "" Then
        OutputView.Data = Right(connector.ReceiveData, Len(OutputView.Data))
    End If
    
    caContainer.ClearAll
    
End Function

Public Function iriscomnew_(OwnerForm As Form, Trn As String, rule As String, InputView, OutputView, _
    Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount, Optional UpdateTrnCountFlag As Boolean) As cSNAResult
    
    If Not Flag610 Then
        eJournalWrite "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ " & cTRNCode
        Set iriscomnew_ = New cSNAResult
        iriscomnew_.ErrCode = GENERIC_COM_ERROR
        iriscomnew_.ErrMessage = "ƒÂÌ ›˜ÂÈ „ﬂÌÂÈ Û˝Ì‰ÂÛÁ (0610)"
        Exit Function
    End If

    OutputView.ClearData
    Dim caContainer As New Buffers
    Dim adoc As MSXML2.DOMDocument30
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If adoc Is Nothing Then Exit Function
    Dim Node As IXMLDOMElement
    Dim ComArea As cXmlComArea
    Dim datanode As IXMLDOMElement
    Dim datadoc As MSXML2.DOMDocument30
    Dim Result As String
    Dim cod_tx As String
    
    Dim anode As MSXML2.IXMLDOMElement
    If Not (xmlIRISRules.documentElement Is Nothing) Then
        If Not xmlIRISRulesUpdate Is Nothing Then
            If Not xmlIRISRulesUpdate.documentElement Is Nothing Then
                Set anode = xmlIRISRulesUpdate.documentElement.selectSingleNode(rule)
            End If
        End If
        If anode Is Nothing And rule <> "" Then Set anode = xmlIRISRules.documentElement.selectSingleNode(rule)
        If Not (anode Is Nothing) Then
            Trn = anode.Attributes(0).Text
            
            Dim codtxattr As IXMLDOMAttribute
            Set codtxattr = anode.Attributes.getNamedItem("CODTX")
            If Not codtxattr Is Nothing Then cod_tx = codtxattr.nodeValue
        End If
    End If
    If IsMissing(AuthUser) Then AuthUser = ""
    AuthUser = UCase(AuthUser)
    Dim authSTD_TRN_I_PARM_V
    If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
        Set authSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
        If Trim(AuthUser) = "" Then
            AuthUser = UCase(authSTD_TRN_I_PARM_V.ByName("ID_EMPL_AUT", 1).Value)
        Else
            authSTD_TRN_I_PARM_V.ByName("ID_EMPL_AUT", 1).Value = AuthUser
        End If
    End If
    If cod_tx <> "" Then
        Dim iSTD_TRN_I_PARM_V, iCUF_USR_ACCESS_D, iSTD_APPL_PARM_V

        If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
            Set iSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
            iSTD_TRN_I_PARM_V.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("STD_TRN_I_PARM_V") = iSTD_TRN_I_PARM_V.Data
        End If
        If Not InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing Then
            Set iSTD_APPL_PARM_V = InputView.ByName("STD_APPL_PARM_V")
            iSTD_APPL_PARM_V.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("STD_APPL_PARM_V") = iSTD_APPL_PARM_V.Data
        End If
        If Not InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
            Set iCUF_USR_ACCESS_D = InputView.ByName("CUF_USR_ACCESS_D")
            iCUF_USR_ACCESS_D.ByName("COD_TX", 1).Value = cod_tx
            InputView.v2Data("CUF_USR_ACCESS_D") = iCUF_USR_ACCESS_D.Data
        End If
        If InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing And _
           InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing And _
           InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
           InputView.v2Value("COD_TX") = cod_tx
        End If
    End If
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='HPSHEADER']", "HPSHEADER", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If Node Is Nothing Then Exit Function
    Set ComArea = New cXmlComArea
    Set ComArea.content = Node
    
    Set ComArea.Container = caContainer
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "HPSHEADER", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datanode Is Nothing Then Exit Function
    Set datadoc = XmlLoadString(datanode.XML, "DataDoc", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datadoc Is Nothing Then Exit Function
    Result = ComArea.LoadXML(datadoc.XML)
    With ComArea.Buffer
    
        Dim authcd As String
        If cod_tx = "" Then
            Dim aiSTD_TRN_I_PARM_V, aiCUF_USR_ACCESS_D, aiSTD_APPL_PARM_V
    
            If Not InputView.xmlDocV2.selectSingleNode("//STD_TRN_I_PARM_V") Is Nothing Then
                Set aiSTD_TRN_I_PARM_V = InputView.ByName("STD_TRN_I_PARM_V")
                authcd = aiSTD_TRN_I_PARM_V.ByName("COD_TX", 1).Value
            End If
            If Not InputView.xmlDocV2.selectSingleNode("//STD_APPL_PARM_V") Is Nothing Then
                Set aiSTD_APPL_PARM_V = InputView.ByName("STD_APPL_PARM_V")
                authcd = aiSTD_APPL_PARM_V.ByName("COD_TX", 1).Value
            End If
            If Not InputView.xmlDocV2.selectSingleNode("//CUF_USR_ACCESS_D") Is Nothing Then
                Set aiCUF_USR_ACCESS_D = InputView.ByName("CUF_USR_ACCESS_D")
                authcd = aiCUF_USR_ACCESS_D.ByName("COD_TX", 1).Value
            End If
        Else
            authcd = cod_tx
        End If
        
        If IsMissing(Appltran) Then Appltran = authcd
        If Trim(Appltran) = "" Then Appltran = authcd
        
        If cod_tx <> "" Then Appltran = cod_tx
        
        If Trim(Appltran) = "" Then Appltran = "ZCBB0CAC"
    
        .v2Value("TRANID") = "XXXX"
        If Not IsMissing(Appltran) Then .v2Value("TRANID") = Trn
        .v2Value("TRAN_SCD") = cHEAD
        If Trim(.v2Value("TRAN_KEY")) = "" Then .v2Value("TRAN_KEY") = "TELLER"
        .v2Value("USER_ID") = UCase(cIRISUserName)
        If Trn = "IRST" Or Trn = "LEYY" Then
            .v2Value("WS_ID") = Left(UCase(cIRISComputerName) & String(10, " "), 10) & Left(UCase(Encode_Greek_(cTERMINALID)) & String(5, " "), 3)
            .v2Value("VERSION") = Right(Left(UCase(Encode_Greek_(cTERMINALID)) & String(5, " "), 5), 2)
        Else
            .v2Value("WS_ID") = Left(UCase(cIRISComputerName) & String(10, " "), 10)
        End If
        .v2Value("ATERM_ID") = cTERMINALID
        .v2Value("YPHRESIA") = cDepartment
        .v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
        .v2Value("APPL_PGM") = UCase(rule)
        .v2Value("BRANCH") = UCase(cBRANCH)
        If Not IsMissing(AuthUser) Then
            .v2Value("AUTH_USER") = UCase(AuthUser)
        End If
        .v2Value("AUTH_TRANS") = UCase(Appltran)
        
        .v2Value("IVFLEN") = InputView.length
        .v2Value("OVFLEN") = OutputView.length
    End With

Dim astr As String, aSize As Long, InputName As String, OutputName As String, res As Integer
Dim onlineAuthError As String
    ComArea.Buffer.GetXMLView
    'ComArea.Buffer.xmlDocV2.save "C:\HPSHeader.xml"
    astr = ComArea.Buffer.Data & InputView.Data & OutputView.Data
    eJournalWrite "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ " & cTRNCode
    
    Dim connector As New cSNAConnection
    
    connector.OpClass = "HPS"
    Dim adescription As String
        
    adescription = InputView.name
    adescription = InputView.BuffType
    If Len(adescription) > 2 Then
        If Right(adescription, 2) = "_I" Then adescription = Left(adescription, Len(adescription) - 2)
    End If
        
    connector.OpCode = Trn
    connector.OpDescription = adescription
    If Not IsMissing(AuthUser) Then connector.AuthUser = Trim(AuthUser)
    
    Set iriscomnew_ = connector.SimpleExec(astr)
    If (iriscomnew_.ErrCode = 0 Or iriscomnew_.ErrCode = COM_OK) And iriscomnew_.SenseCodeMessage = "" Then
        ComArea.Buffer.Data = Left(connector.ReceiveData, Len(ComArea.Buffer.Data))
        If ComArea.Buffer.v2Value("RC") <> 0 Then
            iriscomnew_.ErrCode = ComArea.Buffer.v2Value("RC")
            iriscomnew_.ErrMessage = ComArea.Buffer.v2Value("TRANID") & _
                "(" & ComArea.Buffer.v2Value("RE_MSG") & _
                " ÛÙÔ " & ComArea.Buffer.v2Value("RC_PGM") & ") " & ComArea.Buffer.v2Value("RC_TXT")
        End If
    
        OutputView.Data = Right(connector.ReceiveData, Len(OutputView.Data))
        OutputView.xmlDocV2.documentElement.setAttribute "_journalID", iriscomnew_.MessageID
    
        Dim rc As String, rc_pgm As String, rc_text As String
        rc = ComArea.Buffer.v2Value("RC")
        rc_pgm = ComArea.Buffer.v2Value("RC_PGM")
        rc_text = ComArea.Buffer.v2Value("RC_TXT")
        
        If Is4Eyes(rc, rc_pgm) Then
            Dim authres As String
            authres = L24EyesKey(rc_text)
            If authres <> "" Then
               Dim resultdocument As New MSXML2.DOMDocument30
               resultdocument.LoadXML authres
               If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
                   Load XMLMessageForm
                   Set XMLMessageForm.MessageDocument = resultdocument
                   XMLMessageForm.Show vbModal
                   Set resultdocument = Nothing
               Else
                   Set iriscomnew_ = iriscomnew_(OwnerForm, Trn, rule, InputView, OutputView, resultdocument.selectSingleNode("//MESSAGE/AUTHUSER").Text, Appltran, ErrorView, ErrorCount, False)
               End If
          End If
        End If
    End If
    caContainer.ClearAll
    
End Function


Public Function SNAPool_Communicate(module28flag As Boolean) As Integer
 
On Error GoTo ErrorHandler

Dim com_status As cSNAResult
Dim parse_status As Integer
Dim Result As Boolean
Dim WrongFieldLabel As String
Dim iCount As Integer
Dim CurrentIndex As Integer
Dim i As Integer

Dim StartTickCount As Long
Dim EndTickCount As Long
    
    cTRNTime = 0
    StartTickCount = GetTickCount

    Dim adoc As MSXML2.DOMDocument30
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If adoc Is Nothing Then Exit Function
    
    Dim ComArea As cXmlComArea
    Dim anode As MSXML2.IXMLDOMElement
    
    Dim Node As IXMLDOMElement
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='LEGACYHEADER']", "LEGACYHEADER", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If Node Is Nothing Then Exit Function
    Set ComArea = New cXmlComArea
    Set ComArea.content = Node

    Dim caContainer As New Buffers
    Set ComArea.Container = caContainer
    Dim datanode As IXMLDOMElement
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "LEGACYHEADER", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datanode Is Nothing Then Exit Function
    
    Dim datadoc As MSXML2.DOMDocument30
    Set datadoc = XmlLoadString(datanode.XML, "DataDoc", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datadoc Is Nothing Then Exit Function
    ComArea.LoadXML (datadoc.XML)
    With ComArea.Buffer
    
        .v2Value("TRANID") = Left(cb.send_str, 4)
        If Mid(cb.send_str, 9, 1) = "Œ" Then .v2Value("TRAN_KEY") = "TELL"
        If Mid(cb.send_str, 9, 1) = "A" Then
            .v2Value("TRAN_KEY") = "CHIE"
            .v2Value("AUTH_USER") = UCase(cCHIEFUserName)
        End If
        If Mid(cb.send_str, 9, 1) = "Õ" Then
            .v2Value("TRAN_KEY") = "MANA"
            .v2Value("AUTH_USER") = UCase(cMANAGERUserName)
        End If
        .v2Value("TRAN_SCD") = cHEAD
        .v2Value("USER_ID") = UCase(cIRISUserName)
        .v2Value("WS_ID") = Trim(Left(UCase(cIRISComputerName) & String(10, " "), 10))
        .v2Value("ATERM_ID") = cTERMINALID
        .v2Value("YPHRESIA") = cDepartment
        .v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
        .v2Value("APPL_PGM") = "TR" & Left(cb.send_str, 4)
        .v2Value("BRANCH") = UCase(cBRANCH)
        '.v2Value("AUTH_TRANS") = ""
        
        Dim codtx As String
        codtx = GetCodTx(.v2Value("APPL_PGM"))
        .v2Value("AUTH_TRANS") = codtx
        
        .v2Value("IDFLEN") = Len(cb.send_str)
        If module28flag Then receivelength = 31000 Else receivelength = 2048
        .v2Value("ODFLEN") = receivelength
    End With

    ComArea.Buffer.GetXMLView
    Set cbcomarea = ComArea
    
'-----------------------

Do   'communicate loop
    SNAPool_Communicate = COM_OK
    
    cb.receive_str = ""
    parse_status = PARSE_READ_AGAIN
    ShowStatusMessage "ƒ…¡¬…¬¡”« ƒ≈ƒœÃ≈ÕŸÕ. –≈—…Ã≈Õ≈‘≈..."
    Set com_status = SNAPool_SendReceive(module28flag)
    If com_status.ErrCode <> COM_OK Then
        SNAPool_Communicate = COM_FAILED
        If (cb.receive_str <> "") Then
            If (Trim(com_status.ErrMessage) <> "") Then
                ShowStatusMessage "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage
                Dim afakeowner As Form
                eJournalWriteAll afakeowner, "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage
            Else
                Dim aPos As Integer
                aPos = InStr(1, cb.receive_str, "`")
                Dim Message As String
                
                If (aPos = 0) Then
                    Message = Mid(cb.receive_str, 6)
                Else
                    Message = Mid(Left(cb.receive_str, aPos - 1), 6)
                End If
                
                ShowStatusMessage Message
            End If
        Else
            ShowStatusMessage "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage & " - " & com_status.SenseCodeMessage
            Call NBG_MsgBox("–—œ”œ◊«! ¡–≈‘’◊≈ « À«ÿ« ¡–¡Õ‘«”«”. ¡Õ‘≈ ≈À≈√◊œ √…¡ ‘«Õ ‘’◊« ‘«” ”’Õ¡ÀÀ¡√«”.", True)
            Dim fakeowner As Form
            eJournalWriteAll fakeowner, "ƒ…¡ œ–«= –—œ¬À«Ã¡ ≈–… œ…ÕŸÕ…¡”  " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage & " - " & com_status.SenseCodeMessage & " -  ¡Õ‘≈ ≈À≈√◊œ √…¡ ‘«Õ ‘’◊« ‘«” ”’Õ¡ÀÀ¡√«”."
        End If
        Exit Function
    Else
        ShowStatusMessage "« ”’Õ¡ÀÀ¡√« œÀœ À«—Ÿ»« ≈."
    End If
    
Loop While (False)
cb.BoolTransOk = True
TerminateRead:

Exit Function

ErrorHandler:
    Call Runtime_error("Communicate", Err.number, Err.description)
    SNAPool_Communicate = COM_RUNTIME_ERROR

End Function
Public Function SNAPool_SendReceive(module28flag As Boolean) As cSNAResult
    Dim connector As New cSNAConnection
    Dim Overwrites As New Collection
    Dim ohandler As New cOverwrite
    Dim omessage
    
    Dim oldsend As String
    oldsend = cb.send_str
    
    Dim i As Integer
    
    Do
        Set ohandler = New cOverwrite
        
        Dim IsOverwrite As Boolean
        IsOverwrite = False
        
        For i = 1 To Overwrites.Count
            Set omessage = Overwrites(i)
            If i = 1 Then
                cb.send_str = oldsend & omessage.TimeStamp
            End If
            cb.send_str = cb.send_str & omessage.UpdatedHeader
            cbcomarea.Buffer.v2Value("IDFLEN") = Len(cb.send_str)
        Next i
        
        Set SNAPool_SendReceive = New cSNAResult
        SNAPool_SendReceive.ErrCode = COM_OK
        
        connector.OpClass = "LEGACY"
        Dim adescription As String
            
        adescription = CStr(cTRNCode)
            
        connector.OpCode = Left(cb.send_str, 4)
        connector.OpDescription = adescription
        connector.AuthUser = cbcomarea.Buffer.v2Value("AUTH_USER")
        
        If EventLogWrite Then Call EventLog(8, "Untranslated SEND :" & cb.send_str)
        If SendJournalWrite Then eJournalWrite "S:" & cb.send_str
        
        Dim totalsend As String
        totalsend = cbcomarea.Buffer.Data & AsciiToEbcdic_(cb.send_str) & AsciiToEbcdic_(String(receivelength, " "))
        
        Dim res As cSNAResult
        Set res = connector.SimpleExec(totalsend)
        Set SNAPool_SendReceive = res
        If (res.ErrCode = 0 Or res.ErrCode = COM_OK) And res.SenseCodeMessage = "" Then
            SNAPool_SendReceive.ErrCode = COM_OK
    
            cb.received_data = EbcdicToAscii_(Right(connector.ReceiveData, receivelength))
            cb.receive_str = cb.received_data
            cb.receive_str_length = Len(cb.receive_str)
            
            Dim tempdata As String
            tempdata = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
            If (Len(tempdata) > 39 And Left(tempdata, 2) = "TT") Then
                cb.received_data = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
                ohandler.Data = cb.received_data
                Overwrites.add ohandler
                G0Data.add ohandler.content
                IsOverwrite = True
            Else
                If ReceivedData.Count > 0 Then For i = ReceivedData.Count To 1 Step -1: ReceivedData.Remove i: Next i
                If G0Data.Count > 0 Then For i = G0Data.Count To 1 Step -1: G0Data.Remove (i): Next i
                If module28flag Then
                    Dim aPos As Integer
                    Dim alldata As String
                    aPos = -1
                    alldata = cb.receive_str
                                
                    While aPos <> 0 And alldata <> ""
                        aPos = InStr(1, alldata, "`")
                        Dim Line As String

                        If aPos <> 0 Then
                            If aPos > 1 Then Line = Left(alldata, aPos - 1)
                            ReceivedData.add Line
                            If Len(Line) > 5 And Left(Line, 2) = "02" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            If Len(Line) > 5 And Mid(Line, 2, 3) = "323" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            
                            If aPos < Len(alldata) Then alldata = Right(alldata, Len(alldata) - aPos)
                        ElseIf Trim(alldata) <> "" Then
                            Line = alldata
                            ReceivedData.add Line
                            If Len(Line) > 5 And Left(Line, 2) = "02" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            If Len(Line) > 5 And Mid(Line, 2, 3) = "323" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            
                            alldata = ""
                        
                        End If
                    Wend
                    
                    Dim firstitem As String
                    firstitem = ReceivedData.Item(1)
                    If Len(firstitem) > 5 And Left(firstitem, 5) = "49999" Then
                        firstitem = "731601" & Mid(firstitem, 6)
                        ReceivedData.Remove 1
                        ReceivedData.add firstitem
                        firstitem = "59999" & Mid(firstitem, 6)
                        ReceivedData.add firstitem
                    End If
                    If Left(ReceivedData.Item(1), 1) = "5" Then
                        Dim g5value As String
                        g5value = ReceivedData.Item(1)
                        ReceivedData.Remove (1)
                        ReceivedData.add g5value
                    End If
                    cb.received_data = Mid(ReceivedData(ReceivedData.Count), 6)
                Else
                    cb.received_data = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
                    ReceivedData.add cb.receive_str
                End If
                
                'WriteJournal
                Dim rstr, jstr As String
                For Each rstr In ReceivedData
                    jstr = eJournalClearString(CStr(rstr))
                    If EventLogWrite Then Call EventLog(8, " Translated RECEIVE :" & jstr)
                    If ReceiveJournalWrite Then eJournalWrite "R:" & jstr
                Next
                
            End If
            
            If Len(cb.receive_str) > 5 And Left(cb.receive_str, 5) = "49999" Then
                cb.receive_str = "59999" & Mid(cb.receive_str, 6)
            End If
        
            If (Left(cb.receive_str, 1) = "4") And (Not IsOverwrite) Then
                SNAPool_SendReceive.ErrCode = COM_FAILED
            ElseIf (Left(cb.receive_str, 5) <> "59999") And (Not IsOverwrite) Then
                cbcomarea.Buffer.data_ = Left(connector.ReceiveData, Len(cbcomarea.Buffer.data_))
                Dim rcnode As IXMLDOMElement
                Dim msgnode As IXMLDOMElement
                Set rcnode = GetXmlNode(cbcomarea.Buffer.GetXMLView().documentElement, "//RC", "RC", , "–Ò¸‚ÎÁÏ· ÛÙÔ HandleResp...")
                If Not rcnode Is Nothing Then
                    If Trim(rcnode.Text) <> "0" Then
                        SNAPool_SendReceive.ErrCode = Trim(rcnode.Text)
                        Set msgnode = GetXmlNode(cbcomarea.Buffer.GetXMLView().documentElement, "//RC_TXT", "//RC_TXT", , "–Ò¸‚ÎÁÏ· ÛÙÔ HandleResp...")
                        If Not msgnode Is Nothing Then
                            SNAPool_SendReceive.ErrMessage = Trim(msgnode.Text)
                        End If
                    End If
                End If
            End If
        End If
        If Not IsOverwrite Then Exit Do
        If IsOverwrite Then
            If ohandler.HandleMessage <> vbOK Then
                SNAPool_SendReceive.ErrCode = COM_FAILED
                Exit Do
            End If
        End If
    Loop
    
End Function


Public Function CTGLegacy_Communicate(module28flag As Boolean) As Integer
 
On Error GoTo ErrorHandler

Dim com_status As cSNAResult
Dim parse_status As Integer
Dim Result As Boolean
Dim WrongFieldLabel As String
Dim iCount As Integer
Dim CurrentIndex As Integer
Dim i As Integer

Dim StartTickCount As Long
Dim EndTickCount As Long
    
    cTRNTime = 0
    StartTickCount = GetTickCount

    Dim adoc As MSXML2.DOMDocument30
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If adoc Is Nothing Then Exit Function
    
    Dim ComArea As cXmlComArea
    Dim anode As MSXML2.IXMLDOMElement
    
    Dim Node As IXMLDOMElement
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='LEGACYHEADER_CTG']", "LEGACYHEADER_CTG", "XmlBlocks", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If Node Is Nothing Then Exit Function
    Set ComArea = New cXmlComArea
    Set ComArea.content = Node

    Dim caContainer As New Buffers
    Set ComArea.Container = caContainer
    Dim datanode As IXMLDOMElement
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "LEGACYHEADER_HPS", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datanode Is Nothing Then Exit Function
    
    Dim datadoc As MSXML2.DOMDocument30
    Set datadoc = XmlLoadString(datanode.XML, "DataDoc", "–Ò¸‚ÎÁÏ· ÛÙÁ ‰È·‰ÈÍ·Ûﬂ· ”˝Ì‰ÂÛÁÚ...")
    If datadoc Is Nothing Then Exit Function
    ComArea.LoadXML (datadoc.XML)
    With ComArea.Buffer
    
        .v2Value("TRANID") = Left(cb.send_str, 4)
        .v2Value("REQ_TYPE") = "CMAR"
        .v2Value("APPL_PGM") = "TR" & Left(cb.send_str, 4)
        If Mid(cb.send_str, 9, 1) = "Œ" Then .v2Value("TRAN_KEY") = "TELL"
        If Mid(cb.send_str, 9, 1) = "A" Then
            .v2Value("TRAN_KEY") = "CHIE"
            .v2Value("AUTH_USER") = UCase(cCHIEFUserName)
        End If
        If Mid(cb.send_str, 9, 1) = "Õ" Then
            .v2Value("TRAN_KEY") = "MANA"
            .v2Value("AUTH_USER") = UCase(cMANAGERUserName)
        End If
        .v2Value("TRAN_SCD") = cHEAD
        .v2Value("USER_ID") = UCase(cIRISUserName)
        .v2Value("WS_ID") = Trim(Left(UCase(cIRISComputerName) & String(10, " "), 10))
        .v2Value("TRAN_NBER") = Right("000000" & CStr(cTRNNum), 6)
        .v2Value("SESSID") = StrPad_(CStr(SessID), 2, "0", "L")
        
        Dim codtx As String
        codtx = GetCodTx(.v2Value("APPL_PGM"))
        .v2Value("AUTH_TRANS") = codtx
        
        .v2Value("IDFLEN") = Len(cb.send_str)
        If module28flag Then receivelength = 31000 Else receivelength = 2048
        .v2Value("ODFLEN") = receivelength
    End With

    ComArea.Buffer.GetXMLView
    Set cbcomarea_ctg = ComArea
    
'-----------------------

Do   'communicate loop
    CTGLegacy_Communicate = COM_OK
    
    cb.receive_str = ""
    parse_status = PARSE_READ_AGAIN
    ShowStatusMessage "ƒ…¡¬…¬¡”« ƒ≈ƒœÃ≈ÕŸÕ. –≈—…Ã≈Õ≈‘≈..."
    Set com_status = CTGLegacy_SendReceive(module28flag)
    If com_status.ErrCode <> COM_OK Then
        CTGLegacy_Communicate = COM_FAILED
        If (cb.receive_str <> "") Then
            If (Trim(com_status.ErrMessage) <> "") Then
                ShowStatusMessage "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage
                Dim afakeowner As Form
                eJournalWriteAll afakeowner, "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage
            Else
                Dim aPos As Integer
                aPos = InStr(1, cb.receive_str, "`")
                Dim Message As String
                
                If (aPos = 0) Then
                    Message = Mid(cb.receive_str, 6)
                Else
                    Message = Mid(Left(cb.receive_str, aPos - 1), 6)
                End If
                
                ShowStatusMessage Message
            End If
        Else
            ShowStatusMessage "À¡»œ”: " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage & " - " & com_status.SenseCodeMessage
            Call NBG_MsgBox("–—œ”œ◊«! ¡–≈‘’◊≈ « À«ÿ« ¡–¡Õ‘«”«”. ¡Õ‘≈ ≈À≈√◊œ √…¡ ‘«Õ ‘’◊« ‘«” ”’Õ¡ÀÀ¡√«”.", True)
            Dim fakeowner As Form
            eJournalWriteAll fakeowner, "ƒ…¡ œ–«= –—œ¬À«Ã¡ ≈–… œ…ÕŸÕ…¡”  " & Str(com_status.ErrCode) & " - " & com_status.ErrMessage & " - " & com_status.SenseCodeMessage & " -  ¡Õ‘≈ ≈À≈√◊œ √…¡ ‘«Õ ‘’◊« ‘«” ”’Õ¡ÀÀ¡√«”."
        End If
        Exit Function
    Else
        ShowStatusMessage "« ”’Õ¡ÀÀ¡√« œÀœ À«—Ÿ»« ≈."
    End If
    
Loop While (False)
cb.BoolTransOk = True

Exit Function

ErrorHandler:
    Call Runtime_error("Communicate", Err.number, Err.description)
    CTGLegacy_Communicate = COM_RUNTIME_ERROR

End Function


Private Function CTGLegacy_SendReceive(module28flag As Boolean) As cSNAResult
    Dim connector As New cCTGConnection
    Dim Overwrites As New Collection
    Dim ohandler As New cOverwrite
    Dim omessage
    
    Dim oldsend As String
    oldsend = cb.send_str
    
    Dim i As Integer
    
    Do
        Set ohandler = New cOverwrite
        
        Dim IsOverwrite As Boolean
        IsOverwrite = False
        
        For i = 1 To Overwrites.Count
            Set omessage = Overwrites(i)
            If i = 1 Then
                cb.send_str = oldsend & omessage.TimeStamp
            End If
            cb.send_str = cb.send_str & omessage.UpdatedHeader
            cbcomarea_ctg.Buffer.v2Value("IDFLEN") = Len(cb.send_str)
        Next i
        
        Set CTGLegacy_SendReceive = New cSNAResult
        CTGLegacy_SendReceive.ErrCode = COM_OK
        
        connector.OpClass = "LEGACY_CTG"
        Dim adescription As String
            
        adescription = CStr(cTRNCode)
            
        connector.OpCode = Left(cb.send_str, 4)
        connector.OpDescription = adescription
        connector.AuthUser = cbcomarea_ctg.Buffer.v2Value("AUTH_USER")
        
        If EventLogWrite Then Call EventLog(8, "Untranslated SEND :" & cb.send_str)
        If SendJournalWrite Then eJournalWrite "S:" & cb.send_str
        
        Dim totalsend As String
        totalsend = cbcomarea_ctg.Buffer.Data & AsciiToEbcdic_(cb.send_str) & AsciiToEbcdic_(String(receivelength, " "))
        
        Dim res As cSNAResult
        Set res = connector.SimpleExec(totalsend)
        Set CTGLegacy_SendReceive = res
        If (res.ErrCode = 0 Or res.ErrCode = COM_OK) And res.SenseCodeMessage = "" Then
            CTGLegacy_SendReceive.ErrCode = COM_OK
    
            cb.received_data = EbcdicToAscii_(Right(connector.ReceiveData, receivelength))
            cb.receive_str = cb.received_data
            cb.receive_str_length = Len(cb.receive_str)
            
            Dim tempdata As String
            tempdata = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
            If (Len(tempdata) > 39 And Left(tempdata, 2) = "TT") Then
                cb.received_data = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
                ohandler.Data = cb.received_data
                Overwrites.add ohandler
                G0Data.add ohandler.content
                IsOverwrite = True
            Else
                If ReceivedData.Count > 0 Then For i = ReceivedData.Count To 1 Step -1: ReceivedData.Remove i: Next i
                If G0Data.Count > 0 Then For i = G0Data.Count To 1 Step -1: G0Data.Remove (i): Next i
                If module28flag Then
                    Dim aPos As Integer
                    Dim alldata As String
                    aPos = -1
                    alldata = cb.receive_str
                                
                    While aPos <> 0 And alldata <> ""
                        aPos = InStr(1, alldata, "`")
                        Dim Line As String

                        If aPos <> 0 Then
                            If aPos > 1 Then Line = Left(alldata, aPos - 1)
                            ReceivedData.add Line
                            If Len(Line) > 5 And Left(Line, 2) = "02" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            If Len(Line) > 5 And Mid(Line, 2, 3) = "323" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            
                            If aPos < Len(alldata) Then alldata = Right(alldata, Len(alldata) - aPos)
                        ElseIf Trim(alldata) <> "" Then
                            Line = alldata
                            ReceivedData.add Line
                            If Len(Line) > 5 And Left(Line, 2) = "02" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            If Len(Line) > 5 And Mid(Line, 2, 3) = "323" Then G0Data.add Mid(Line, 6, Len(Line) - 1)
                            
                            alldata = ""
                        
                        End If
                    Wend
                    
                    Dim firstitem As String
                    firstitem = ReceivedData.Item(1)
                    If Len(firstitem) > 5 And Left(firstitem, 5) = "49999" Then
                        firstitem = "731601" & Mid(firstitem, 6)
                        ReceivedData.Remove 1
                        ReceivedData.add firstitem
                        firstitem = "59999" & Mid(firstitem, 6)
                        ReceivedData.add firstitem
                    End If
                    If Left(ReceivedData.Item(1), 1) = "5" Then
                        Dim g5value As String
                        g5value = ReceivedData.Item(1)
                        ReceivedData.Remove (1)
                        ReceivedData.add g5value
                    End If
                    cb.received_data = Mid(ReceivedData(ReceivedData.Count), 6)
                Else
                    cb.received_data = Mid(cb.receive_str, 6, Len(cb.receive_str) - 5)
                    ReceivedData.add cb.receive_str
                End If
                
                'WriteJournal
                Dim rstr, jstr As String
                For Each rstr In ReceivedData
                    jstr = eJournalClearString(CStr(rstr))
                    If EventLogWrite Then Call EventLog(8, " Translated RECEIVE :" & jstr)
                    If ReceiveJournalWrite Then eJournalWrite "R:" & jstr
                Next
                
            End If
            
            If Len(cb.receive_str) > 5 And Left(cb.receive_str, 5) = "49999" Then
                cb.receive_str = "59999" & Mid(cb.receive_str, 6)
            End If
        
            If (Left(cb.receive_str, 1) = "4") And (Not IsOverwrite) Then
                CTGLegacy_SendReceive.ErrCode = COM_FAILED
            ElseIf (Left(cb.receive_str, 5) <> "59999") And (Not IsOverwrite) Then
                cbcomarea_ctg.Buffer.data_ = Left(connector.ReceiveData, Len(cbcomarea_ctg.Buffer.data_))
                Dim rcnode As IXMLDOMElement
                Dim msgnode As IXMLDOMElement
                Set rcnode = GetXmlNode(cbcomarea_ctg.Buffer.GetXMLView().documentElement, "//RC", "RC", , "–Ò¸‚ÎÁÏ· ÛÙÔ HandleResp...")
                If Not rcnode Is Nothing Then
                    If Trim(rcnode.Text) <> "0" Then
                        CTGLegacy_SendReceive.ErrCode = Trim(rcnode.Text)
                        Set msgnode = GetXmlNode(cbcomarea_ctg.Buffer.GetXMLView().documentElement, "//RC_TXT", "//RC_TXT", , "–Ò¸‚ÎÁÏ· ÛÙÔ HandleResp...")
                        If Not msgnode Is Nothing Then
                            CTGLegacy_SendReceive.ErrMessage = Trim(msgnode.Text)
                        End If
                    End If
                End If
            End If
        End If
        If Not IsOverwrite Then Exit Do
        If IsOverwrite Then
            If ohandler.HandleMessage <> vbOK Then
                CTGLegacy_SendReceive.ErrCode = COM_FAILED
                Exit Do
            End If
        End If
    Loop
    
End Function

