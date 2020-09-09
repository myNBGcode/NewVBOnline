Attribute VB_Name = "DatabaseMdl"
Option Explicit

Public xmlProfile As New MSXML2.DOMDocument
'Public xmlMenu As New MSXML2.DOMDocument
Public xmlNewMenu As New MSXML2.DOMDocument

Public xmlModel As New MSXML2.DOMDocument
Public xmlBranchProfile As New MSXML2.DOMDocument
Public xmlUserProfile As New MSXML2.DOMDocument
Public xmlNewProfiles As New MSXML2.DOMDocument
Public xmlNewUserProfiles As New MSXML2.DOMDocument
Public xmlNewBranchProfile As New MSXML2.DOMDocument
Public xmlCRAStructures As New MSXML2.DOMDocument
Public xmlCRAErrors As New MSXML2.DOMDocument
Public xmlIRISStructures As New MSXML2.DOMDocument30
Public xmlIRISRules As New MSXML2.DOMDocument30

Public xmlComAreaCodTX As New MSXML2.DOMDocument60
Public XML4Eyes As New MSXML2.DOMDocument60

Public xmlIRISStructuresUpdate As New MSXML2.DOMDocument30
Public xmlIRISRulesUpdate As New MSXML2.DOMDocument30

Public xmlWebLinks As New MSXML2.DOMDocument30

Public rsIRISErrors As New ADODB.Recordset
Public rsFTFila As New ADODB.Recordset
Public IRISAuthList As New Collection, IRISAuthNames As New Collection

Public xmlXSLTPack As New MSXML2.DOMDocument30
Public SkipCRAUse As Boolean

Public ado_DB As ADODB.Connection, trade_db As ADODB.Connection

Public VBTradeSLink

Public trnModel As New Collection
Public stepModel As New Collection
Public fldModel As New Collection
Public lblModel As New Collection
Public listModel As New Collection
Public gridModel As New Collection
Public FldTypeList As New Collection
Public btnModel As New Collection
Public chkModel As New Collection
Public cmbModel As New Collection
Public chrModel As New Collection

Public PrinterNameList As New Collection

'---------------------------------------------------------------------
Public Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadID As Long
End Type

Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Const NORMAL_PRIORITY_CLASS = &H20&

Public Function ExecCmd(cmdline$) As Long
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret As Long

   ' Initialize the STARTUPINFO structure:
   start.cb = Len(start)

   ' Start the shelled application:
   ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


   ' Wait for the shelled application to finish:
      ret& = WaitForSingleObject(proc.hProcess, INFINITE)
      Call GetExitCodeProcess(proc.hProcess, ret&)
      Call CloseHandle(proc.hThread)
      Call CloseHandle(proc.hProcess)
      ExecCmd = ret&
End Function

Private Function TestFunc(ByVal lVal As Long) As Integer
'this function is necessary since the value returned by Shell is an
'unsigned integer and may exceed the limits of a VB integer
    If (lVal And &H8000&) = 0 Then TestFunc = lVal And &HFFFF& _
    Else TestFunc = &H8000 Or (lVal And &H7FFF&)
End Function

Public Function fnChkFileExistAbs(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Open sFileName For Input As #1
    Close #1
    
    fnChkFileExistAbs = True
    Exit Function
returnfalse:
    fnChkFileExistAbs = False
End Function

Public Function fnChkFileExist(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Open WorkDir & sFileName For Input As #1
    Close #1
    
    fnChkFileExist = True
    Exit Function
returnfalse:
    fnChkFileExist = False
End Function

Public Sub sbWriteLogFile(ByVal sFileName As String, sMessage As String)
On Error GoTo returnfalse
    Open WorkDir & sFileName For Output As #1
    Print #1, sMessage
    Close #1
    Exit Sub
returnfalse:
    LogMsgbox "« ‰ÁÏÈÔıÒ„ﬂ· ÙÔı ·Ò˜ÂﬂÔı " & sFileName & " ·›Ùı˜Â.", vbInformation
End Sub

Public Sub sbWriteLogFileNew(ByVal sFileName As String, sMessage As String)
On Error GoTo returnfalse
    Open NetworkHomeDir & "\" & sFileName For Output As #1
    Print #1, sMessage
    Close #1
    Exit Sub
returnfalse:
    LogMsgbox "« ‰ÁÏÈÔıÒ„ﬂ· ÙÔı ·Ò˜ÂﬂÔı " & NetworkHomeDir & "\" & sFileName & " ·›Ùı˜Â.", vbInformation
End Sub

Public Function StrPad_(PString As String, PIntLen As Integer, Optional PStrChar As Variant, _
    Optional PStrLftRgt As Variant) As String

' « Function StrPad ‰›˜ÂÙ·È ›Ì· string Â‰ﬂÔ Í·È
' ÂÈÛÙÒ›ˆÂÈ ›Ì· string ÙÔı ÏﬁÍÔıÚ Ôı ÔÒﬂÊÂÙ·È
' ÒÔÛË›ÙÔÌÙ·Ú ‰ÂÓÈ‹ ﬁ ·ÒÈÛÙÂÒ‹ ÙÔÌ ˜·Ò·ÍÙﬁÒ· Ôı
' ÔÒﬂÊÂÙ·È ¸ÛÂÚ ˆÔÒ›Ú ˜ÒÂÈ‹ÊÂÙ·È ˛ÛÙÂ ÙÔ Input Â‰ﬂÔ
' Ì· „ﬂÌÂÈ ÛÙÔ ÂÈËıÏÁÙ¸ ÏﬁÍÔÚ
'
' –·Ò‹ÏÂÙÒÔÈ :
' PString    ÙÔ input String
' PIntLen    ÙÔ ÏﬁÍÔÚ ÙÔı string Ôı Ë· ÂÈÙÒ›¯ÂÈ
' PStrChar   ÒÔ·ÈÒÂÙÈÍ‹ Ô ˜·Ò·ÍÙﬁÒ·Ú Ôı Ë· „ÂÏﬂÛÂÈ
'            ÙÔ ı¸ÎÔÈÔ ÏﬁÍÔÚ default <SPACE>
' PStrLftRgt ÒÔ·ÈÒÂÙÈÍ‹ ·Ì Ë· ÒÔÛË›ÛÂÈ ˜·Ò·ÍÙﬁÒÂÚ
'            ‰ÂÓÈ‹ (R) ﬁ ·ÒÈÛÙÂÒ‹ (L) default ‰ÂÓÈ‹ (L)
'
' .˜.
' StrPad("12345",10)         -> "     12345"
' StrPad("12345",10, ,"R")   -> "12345     "
' StrPad("12345",10,"0")     -> "0000012345"
' StrPad("12345",10,"0","R") -> "1234500000"
' StrPad("12345",4)          -> "2345"
' StrPad("12345",4, ,"R")    -> "1234"
    
    If PIntLen <= 0 Then StrPad_ = "": Exit Function
    Dim MString As String, minti As Integer
    
    If IsMissing(PStrChar) Then PStrChar = " "
    If IsMissing(PStrLftRgt) Then PStrLftRgt = "L"
    
    For minti = 1 To PIntLen: MString = MString + PStrChar: Next

    If PStrLftRgt Like "[Ll]" Then StrPad_ = Right(MString + PString, PIntLen) _
    Else StrPad_ = Left(PString + MString, PIntLen)
End Function

Public Function NodeIntegerFld(inParentNode As MSXML2.IXMLDOMElement, _
    inValueName As String, inModel As Collection) As Integer
Dim astr As String
    astr = inModel.Item(inValueName)

    If astr <> "" Then
On Error GoTo noelement
        astr = inParentNode.selectSingleNode(astr).Text
        If astr = "" Then
            NodeIntegerFld = 0
        Else
            NodeIntegerFld = astr
            Exit Function
        End If
noelement:
        NodeIntegerFld = 0
    Else
        NodeIntegerFld = 0
    End If
End Function

Public Function NodeBooleanFld(inParentNode As MSXML2.IXMLDOMElement, _
    inValueName As String, inModel As Collection) As Boolean
Dim astr As String
    astr = inModel.Item(inValueName)

    If astr <> "" Then
        On Error GoTo noelement
        astr = inParentNode.selectSingleNode(astr).Text
        If astr = "" Then
            NodeBooleanFld = False
        Else
            NodeBooleanFld = astr
            Exit Function
        End If
noelement:
        NodeBooleanFld = False
    Else
        NodeBooleanFld = False
    End If
End Function

Private Function SetCRLF(inString As String) As String
Dim i As Integer, alength As Integer, astr As String
On Error GoTo 0
    
'    astr = inString
    astr = Replace(inString, "(=)", vbCrLf)
    
'    i = InStr(1, astr, "(=)", vbBinaryCompare)
'    If i > 0 Then
'        While i > 0
'            alength = Len(astr)
'            astr = Left(astr, i - 1) & vbCrLf & Right(astr, alength - i - 2)
'            i = InStr(1, astr, "(=)", vbBinaryCompare)
'        Wend
'    End If
'    SetCRLF = SetSpace(astr)
    SetCRLF = Replace(astr, Chr(250), " ")
End Function

Private Function SetSpace(inString As String) As String
Dim i As Integer, alength As Integer, astr As String
On Error GoTo 0
    astr = inString
    i = InStr(1, astr, Chr(250), vbBinaryCompare)
    If i > 0 Then
        While i > 0
            alength = Len(astr)
            astr = Left(astr, i - 1) & " " & Right(astr, alength - i)
            i = InStr(1, astr, Chr(250), vbBinaryCompare)
        Wend
    End If
    SetSpace = astr
End Function

Public Function NodeStringFld(inParentNode As MSXML2.IXMLDOMElement, _
    inValueName As String, inModel As Collection) As String
Dim astr As String
    astr = inModel.Item(inValueName)
    If astr <> "" Then
On Error GoTo noelement
        NodeStringFld = SetCRLF(inParentNode.selectSingleNode(astr).Text)
        Exit Function
noelement:
        NodeStringFld = ""
    Else
        NodeStringFld = ""
    End If
End Function

Private Function LoadProfilePart(inName As String)
Dim GroupNode, adoc As MSXML2.DOMDocument
    inName = Replace(inName, " ", "_")
    If Len(inName) > 0 And IsNumeric(Left(inName, 1)) Then inName = "_" & inName
    On Error GoTo SkipLoadProfilePart
    Set GroupNode = xmlProfile.documentElement.selectSingleNode(UCase(inName))
    If Not (GroupNode Is Nothing) Then
        Set adoc = New MSXML2.DOMDocument
        adoc.Load ReadDir & GroupNode.Text & ".xml"
    End If
                
    GoTo AfterSkipLoadProfilePart
SkipLoadProfilePart:
    Set adoc = Nothing
AfterSkipLoadProfilePart:
    Set LoadProfilePart = adoc
End Function

Private Function LoadProfilePartNew(inName As String)
Dim GroupNode, adoc As MSXML2.DOMDocument
    inName = Replace(inName, " ", "_")
    If Len(inName) > 0 And IsNumeric(Left(inName, 1)) Then inName = "_" & inName
    On Error GoTo SkipLoadProfilePart
    Set GroupNode = xmlProfile.documentElement.selectSingleNode(UCase(inName))
    If Not (GroupNode Is Nothing) Then
        Set adoc = New MSXML2.DOMDocument
        adoc.LoadXML xmlNewProfiles.selectSingleNode("//Profile[@name='" + UCase(GroupNode.Text) + "']").XML
    End If
                
    GoTo AfterSkipLoadProfilePart
SkipLoadProfilePart:
    Set adoc = Nothing
AfterSkipLoadProfilePart:
    Set LoadProfilePartNew = adoc
End Function


'Public Function PrepareProfiles()
'Dim GroupNode, anode, aGroupName, adoc As MSXML2.DOMDocument, Line As Integer
'On Error GoTo ProfileError
'    If Not (xmlProfile Is Nothing) Then
'        Line = 10: PrepareProfiles = False
'        Line = 11: xmlBranchProfile.Load (ReadDir & cBranchProfileName & ".xml")
'
'        If LocalFlag Then
'            xmlUserProfile.Load (ReadDir & cBranchProfileName & ".xml")
'        Else
'            For Each aGroupName In UserGroups
'                Line = 12: Set adoc = LoadProfilePart(CStr(aGroupName))
'                Line = 13:
'                If Not (adoc Is Nothing) Then
'                    If Not (adoc.documentElement Is Nothing) Then
'                        Line = 14:
'                        If xmlUserProfile.documentElement Is Nothing Then
'                            Line = 15: Set xmlUserProfile.documentElement = adoc.documentElement.cloneNode(True)
'                        Else
'                            Line = 16:
'                            For Each anode In adoc.documentElement.childNodes
'                                Line = 17: xmlUserProfile.documentElement.appendChild anode.cloneNode(True)
'                            Next anode
'                        End If
'                    End If
'                End If
'            Next aGroupName
'        End If
'        PrepareProfiles = True
'        Exit Function
'    Else
'
'        Line = 50: PrepareProfiles = False
'        Line = 51: xmlBranchProfile.Load (ReadDir & cBranchProfileName & ".xml")
'        Line = 52: xmlUserProfile.Load (ReadDir & cUserProfileName & ".xml")
'        Line = 53: PrepareProfiles = True
'        Exit Function
'    End If
'ProfileError:
'        NBG_LOG_MsgBox "À‹ËÔÚ ÛÙÁÌ ≈ÌÂÒ„ÔÔﬂÁÛÁ Profile... (¡1) " & Line & "-" & error(), True, "À¡»œ”"
'End Function

Public Function PrepareNewProfiles()
    Dim GroupNode, anode, aGroupName, adoc As MSXML2.DOMDocument, Line As Integer
   
    On Error GoTo ProfileError
    If Not (xmlProfile Is Nothing) Then
    
        Line = 19: xmlNewProfiles.Load (ReadDir & "newprofiles.xml")
        Line = 20: PrepareNewProfiles = False
        Line = 21: xmlNewBranchProfile.LoadXML xmlNewProfiles.selectSingleNode("//Profile[@name='" & UCase(cBranchProfileName) & "']").XML
        
        If (LocalFlag) Then
            xmlNewUserProfiles.LoadXML xmlNewProfiles.selectSingleNode("//Profile[@name='" & UCase(cBranchProfileName) & "']").XML
            Exit Function
        Else
        
            Dim addedProfiles As New Collection
            For Each aGroupName In UserGroups
                 Line = 12: Set adoc = LoadProfilePartNew(CStr(aGroupName))
                    Line = 13:
                    If Not (adoc Is Nothing) Then
                        If Not (adoc.documentElement Is Nothing) Then
                            Line = 14:
                            If xmlNewUserProfiles.documentElement Is Nothing Then
                                Line = 15: Set xmlNewUserProfiles.documentElement = adoc.documentElement.cloneNode(True)
                            Else
                                Line = 16:
                                For Each anode In adoc.documentElement.childNodes
                                    Line = 17: xmlNewUserProfiles.documentElement.appendChild anode.cloneNode(True)
                                Next anode
                            End If
                            addedProfiles.add (adoc.documentElement.Attributes.getNamedItem("name").Text)
                        End If
                    End If
            Next aGroupName
            
            Dim trnProfile As MSXML2.IXMLDOMNode, typeAttr As MSXML2.IXMLDOMAttribute
            Dim aprofile, found As Boolean
            found = False
            
            For Each trnProfile In xmlNewProfiles.documentElement.childNodes
                If cUseCicsUserInfo And cHasTellerTrnGroup And trnProfile.Attributes.getNamedItem("name").Text = "TELLER" Then
                    found = False
                    Set aprofile = Nothing
                    For Each aprofile In addedProfiles
                        If aprofile = "TELLER" Then
                            found = True
                            Exit For
                        End If
                    Next aprofile
                    If Not found Then
                        If xmlNewUserProfiles.documentElement Is Nothing Then
                            Line = 15: Set xmlNewUserProfiles.documentElement = trnProfile.cloneNode(True)
                        Else
                            Line = 16:
                            Set anode = Nothing
                            For Each anode In trnProfile.childNodes
                                Line = 17: xmlNewUserProfiles.documentElement.appendChild anode.cloneNode(True)
                            Next anode
                        End If
                    End If
                End If
                
                Set typeAttr = trnProfile.Attributes.getNamedItem("type")
                If Not (typeAttr Is Nothing) Then
                    If typeAttr.Text = "public" Then
                        Line = 14:
                        
                        found = False
                        Set aprofile = Nothing
                        For Each aprofile In addedProfiles
                            If trnProfile.Attributes.getNamedItem("name").Text = aprofile Then
                                found = True
                                Exit For
                            End If
                        Next aprofile
                        
                        If Not found Then
                            If xmlNewUserProfiles.documentElement Is Nothing Then
                                Line = 15: Set xmlNewUserProfiles.documentElement = trnProfile.cloneNode(True)
                            Else
                                Line = 16:
                                Set anode = Nothing
                                For Each anode In trnProfile.childNodes
                                    Line = 17: xmlNewUserProfiles.documentElement.appendChild anode.cloneNode(True)
                                Next anode
                            End If
                        End If
                        
                        If xmlNewBranchProfile.documentElement Is Nothing Then
                            Line = 17: Set xmlNewBranchProfile.documentElement = trnProfile.cloneNode(True)
                        Else
                            Line = 18:
                            Set anode = Nothing
                            For Each anode In trnProfile.childNodes
                                Line = 17: xmlNewBranchProfile.documentElement.appendChild anode.cloneNode(True)
                            Next anode
                        End If
                    
                    End If
                End If
            Next trnProfile
        
        End If
        
        PrepareNewProfiles = True
        Exit Function
    Else
        Line = 50: PrepareNewProfiles = False
        Line = 51: xmlNewBranchProfile.LoadXML xmlNewProfiles.selectSingleNode("//Profile[@name='" & UCase(cBranchProfileName) & "']").XML
        Line = 52: xmlNewUserProfiles.LoadXML xmlNewProfiles.selectSingleNode("//Profile[@name='" & UCase(cUserProfileName) & "']").XML
    End If
    
ProfileError:
    NBG_LOG_MsgBox "À‹ËÔÚ ÛÙÁÌ ≈ÌÂÒ„ÔÔﬂÁÛÁ NewProfile... (¡1) " & Line & "-" & error(), True, "À¡»œ”"
End Function
Public Sub prepareIRISUpdate()
    Set xmlIRISStructuresUpdate = Nothing
    Set xmlIRISRulesUpdate = Nothing
    On Error Resume Next
    xmlIRISStructuresUpdate.Load ReadDir & "IRISStructuresUpdate.xml"
    xmlIRISRulesUpdate.Load ReadDir & "IRISRulesUpdate.xml"
    If xmlIRISStructuresUpdate.XML = "" Then
        Set xmlIRISStructuresUpdate = Nothing
    Else
        If xmlIRISStructuresUpdate.documentElement Is Nothing Or xmlIRISStructuresUpdate.Text = "" Then Set xmlIRISStructuresUpdate = Nothing
    End If
    If xmlIRISRulesUpdate.XML = "" Then
        Set xmlIRISRulesUpdate = Nothing
    Else
        If xmlIRISRulesUpdate.documentElement Is Nothing Or xmlIRISRulesUpdate.Text = "" Then Set xmlIRISRulesUpdate = Nothing
    End If
End Sub

Public Sub prepareIRIS()
    xmlIRISStructures.Load ReadDir & "IRISStructures.xml"
    On Error Resume Next
    xmlIRISRules.Load ReadDir & "IRISRules.xml"
    On Error GoTo 0
    xmlXSLTPack.Load ReadDir & "XSLTPack.xml"
    rsIRISErrors.open ReadDir & "IRISErrors.xml", "Provider=msPersist", adOpenStatic, adLockReadOnly
    rsFTFila.open ReadDir & "FT_FILA_TBL.xml", "Provider=msPersist", adOpenStatic, adLockReadOnly
    
End Sub

Public Sub prepareWebLinks()
    On Error GoTo defaultWebLinks
    xmlWebLinks.Load ReadDir & "WebLink.xml"
    On Error GoTo 0
    If xmlWebLinks.Text = "" Then GoTo defaultWebLinks
    
    Dim aattr As IXMLDOMAttribute

    Set aattr = xmlWebLinks.documentElement.Attributes.getNamedItem("environment")
    If Not (aattr Is Nothing) Then
        If Trim(aattr.Text) <> "" Then
            WorkEnvironment_ = Left(Trim(aattr.Text), 4)
            WorkEnvironment_ = String(7, "0") & "." & WorkEnvironment_ & String(4, "0")
            UpdatexmlEnvironment "WorkEnvironment", Left(Right(WorkEnvironment_, 8), 4)
        End If
    End If
    
    Dim links As IXMLDOMElement
    Set links = xmlWebLinks.documentElement.selectSingleNode("//WEBLINKS/V1")
    Dim link As IXMLDOMElement
    For Each link In links.childNodes
        If link.Text <> "" Then
            WebLinks.add link.Text, link.nodename
        End If
    Next link
    GoTo ExitPoint

defaultWebLinks:
    WebLinks.add "", "EDUCTRADEWEBLINK"
    WebLinks.add "http://N00000032/VirtualTradeEduc/soap", "PRODTRADEWEBLINK"
    WebLinks.add "", "EDUCADMINWEBLINK"
    WebLinks.add "http://N00000032/TRNStatistics/soap", "PRODADMINWEBLINK"
    WebLinks.add "", "EDUCKPSWEBLINK"
    WebLinks.add "http://N00000032/KPSRequest/soap", "PRODKPSWEBLINK"
    
ExitPoint:

End Sub

Public Function GetFTFilaName_(inTbl As String, inAppl As String, invalue As String) As String
    Dim astr As String
    astr = ""
    rsFTFila.Filter = "COD_TBL_REF='" & inTbl & "' and COD_APLCCN_SUBAPL = '" & inAppl & "' and CLAVE_FILA = '" & invalue & "'"
    If rsFTFila.RecordCount = 1 Then astr = rsFTFila!DESCR_CORTA
    rsFTFila.Filter = ""
    GetFTFilaName_ = astr
End Function

Public Function GetFTFilaDescription_(inTbl As String, inAppl As String, invalue As String) As String
    Dim astr As String
    astr = ""
    rsFTFila.Filter = "COD_TBL_REF='" & inTbl & "' and COD_APLCCN_SUBAPL = '" & inAppl & "' and CLAVE_FILA = '" & invalue & "'"
    If rsFTFila.RecordCount = 1 Then astr = rsFTFila!DESCR_LARGA
    rsFTFila.Filter = ""
    GetFTFilaDescription_ = astr
End Function
Public Function GetIRISErrorData_(inKeys() As String) As Variant
   Dim ReturnArray() As String
   Dim i As Integer
   ReDim ReturnArray(UBound(inKeys))
   For i = LBound(inKeys) To UBound(inKeys)
         rsIRISErrors.Filter = "VALUE_IMP_NAME=" & "'" & inKeys(i) & "'"
         If rsIRISErrors.RecordCount = 1 Then
            ReturnArray(i) = rsIRISErrors.fields("DATA").value
         Else
            ReturnArray(i) = ""
         End If
   Next i
   rsIRISErrors.Filter = ""
   GetIRISErrorData_ = ReturnArray()
End Function

Public Function GetCodTx(comareaName As String) As String

    GetCodTx = ""
    comareaName = Trim(UCase(comareaName))
    
    Dim anode As IXMLDOMNode
    Dim aattr As IXMLDOMAttribute
    
    Set anode = xmlComAreaCodTX.selectSingleNode("//LIST/ComArea[@name='" & comareaName & "']")
    If Not anode Is Nothing Then
        Set aattr = anode.Attributes.getNamedItem("codtx")
        If aattr Is Nothing Then
            GetCodTx = "H" & Right(comareaName, 4)
        Else
            GetCodTx = aattr.Text
        End If
    End If

End Function
Public Function Is4Eyes(rc As String, rcmodule As String) As Boolean

    If (rc = "0") Then
        Is4Eyes = False
        Exit Function
    End If
    rcmodule = Trim(UCase(rcmodule))
    Dim anode As IXMLDOMNode
    Set anode = XML4Eyes.selectSingleNode("//eyes4/resp[@rc='" & rc & "' and @module='" & rcmodule & "']")
    If Not anode Is Nothing Then
        Is4Eyes = True
        Exit Function
    End If
    Is4Eyes = False
    
End Function
Public Function IsFake4EyesHpsRC(fakeRc As String) As Boolean
    If (fakeRc = "0" Or fakeRc = "") Then
        IsFake4EyesHpsRC = False
        Exit Function
    End If
    Dim anode As IXMLDOMNode
    Set anode = XML4Eyes.selectSingleNode("//eyes4/resp[@hps_rc='" & fakeRc & "']")
    If Not anode Is Nothing Then
       IsFake4EyesHpsRC = True
       Exit Function
    End If
    IsFake4EyesHpsRC = False
End Function

Public Sub PrepareXML()
Dim anode, stepnode
Dim i As Integer
    xmlNewMenu.Load ReadDir & "newmenu.xml"
    xmlModel.Load ReadDir & "model.xml"
    xmlComAreaCodTX.Load ReadDir & "CACodTx.xml"
    XML4Eyes.Load ReadDir & "foureyes.xml"
    
    On Error GoTo SkipProfileList
    xmlProfile.Load ReadDir & "profiles.xml"
    GoTo AfterSkipProfileList
SkipProfileList:
    Set xmlProfile = Nothing
AfterSkipProfileList:
    

On Error GoTo SkipCRAUse_
    xmlCRAStructures.Load (ReadDir & "CRAStructures.xml")
    xmlCRAErrors.Load (ReadDir & "CRAErrors.xml")
    GoTo AfterSkipCRAUse_
SkipCRAUse_:
    SkipCRAUse = True
AfterSkipCRAUse_:
    On Error GoTo 0
    prepareIRIS
    'prepareWebLinks
    
    
    Set anode = xmlModel.documentElement.selectSingleNode("TRN")
    For i = 0 To anode.childNodes.length - 1
        trnModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = xmlModel.documentElement.selectSingleNode("STEP")
    Set anode = anode.selectSingleNode("S")
    Set stepnode = anode
    For i = 0 To anode.childNodes.length - 1
        stepModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = stepnode.selectSingleNode("FIELDS")
    Set anode = anode.selectSingleNode("FLD")
    For i = 0 To anode.childNodes.length - 1
        fldModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = stepnode.selectSingleNode("BUTTONS")
    Set anode = anode.selectSingleNode("BTN")
    For i = 0 To anode.childNodes.length - 1
        btnModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = stepnode.selectSingleNode("CHECKS")
    Set anode = anode.selectSingleNode("CHK")
    For i = 0 To anode.childNodes.length - 1
        chkModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = stepnode.selectSingleNode("COMBOS")
    Set anode = anode.selectSingleNode("CMB")
    For i = 0 To anode.childNodes.length - 1
        cmbModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = stepnode.selectSingleNode("LABELS")
    Set anode = anode.selectSingleNode("LABEL")
    For i = 0 To anode.childNodes.length - 1
        lblModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = xmlModel.documentElement.selectSingleNode("LIST")
    Set anode = anode.selectSingleNode("L")
    For i = 0 To anode.childNodes.length - 1
        listModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = xmlModel.documentElement.selectSingleNode("GRID")
    Set anode = anode.selectSingleNode("G")
    For i = 0 To anode.childNodes.length - 1
        gridModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = xmlModel.documentElement.selectSingleNode("TYPE")
    For i = 0 To anode.childNodes.length - 1
        Dim cdnode As IXMLDOMElement
        Set cdnode = anode.childNodes.Item(i).selectSingleNode("CD")
        If Not (cdnode Is Nothing) Then
            FldTypeList.add anode.childNodes.Item(i), "T" & cdnode.Text
        Else
            FldTypeList.add anode.childNodes.Item(i), anode.childNodes.Item(i).tagName
        End If
    Next i
    
    Set anode = stepnode.selectSingleNode("CHARTS")
    Set anode = anode.selectSingleNode("CHR")
    For i = 0 To anode.childNodes.length - 1
        chrModel.add anode.childNodes.Item(i).Text, anode.childNodes.Item(i).tagName
    Next i
    
    Set anode = xmlModel.documentElement.selectSingleNode("KEYS")
    cTELLERKEY = anode.selectSingleNode("TELLERKEY").Text
    cCHIEFKEY = anode.selectSingleNode("CHIEFKEY").Text
    cMANAGERKEY = anode.selectSingleNode("MANAGERKEY").Text
    cTELLERCHIEFKEY = anode.selectSingleNode("TELLERCHIEFKEY").Text
    cTELLERMANAGERKEY = anode.selectSingleNode("TELLERMANAGERKEY").Text
        
End Sub

Public Sub UpdateParams()
    
    Dim astr As String, recno As Integer
    UpdatexmlEnvironment "POSTDATE", format(cPOSTDATE, "ddmmyyyy")
    UpdatexmlEnvironment "TRNNUM", CStr(cTRNNum)

End Sub
Public Sub PrepareEnv()

    If Trim(WorkEnvironment_) = "" Then
        Dim aReg As New cRegistry, Line As Integer, i As Integer
        aReg.ClassKey = HKEY_LOCAL_MACHINE
        aReg.SectionKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        aReg.ValueKey = "IRIS_DB_NAME"
        If aReg.KeyExists Then WorkEnvironment_ = aReg.value Else WorkEnvironment_ = ""
    
        If WorkEnvironment_ = "" Then WorkEnvironment_ = Trim(Mid(Left(LogonServer & String(9, "0"), 9), 3, 9)) & ".PROD" & Right("0000" & cBRANCH, 4)
    End If

    If Left(Right(WorkEnvironment_, 8), 4) <> "EDUC" And Left(Right(WorkEnvironment_, 8), 4) <> "PROD" Then
        MsgBox "À¡»œ” –≈—…¬¡ÀÀœÕ À≈…‘œ’—√…¡”: " & WorkEnvironment_, vbCritical
        WorkEnvironment_ = Trim(Mid(LogonServer, 3, 9)) & ".PROD" & Right("0000" & cBRANCH, 4)
    End If
    UpdatexmlEnvironment "WorkEnvironment", Left(Right(WorkEnvironment_, 8), 4)

    cPOSTDATE = Date

End Sub

Public Function ChkBranchProfileAccess(aTrnCD As Long)
Dim trnNode
Dim aFlag As Integer

ChkBranchProfileAccess = False

Set trnNode = xmlBranchProfile.documentElement.selectSingleNode("T" & CStr(aTrnCD))
If Not (trnNode Is Nothing) Then aFlag = CInt(trnNode.Text) Else aFlag = 0
ChkBranchProfileAccess = (aFlag = 1)

End Function

Public Function ChkProfileAccess(aTrnCD As Long)
Dim trnNode
Dim aFlag As Integer

ChkProfileAccess = False

Set trnNode = xmlBranchProfile.documentElement.selectSingleNode("T" & CStr(aTrnCD))
If Not (trnNode Is Nothing) Then aFlag = CInt(trnNode.Text) Else aFlag = 0
ChkProfileAccess = (aFlag = 1)
If aFlag = 1 Then
    Set trnNode = xmlUserProfile.documentElement.selectSingleNode("T" & CStr(aTrnCD))
    If Not (trnNode Is Nothing) Then aFlag = CInt(trnNode.Text) Else aFlag = 0
    ChkProfileAccess = (aFlag = 1)
End If

End Function


Public Function ChkProfileAccessNew(aTrnCD As Long)
Dim trnNode
Dim aFlag As Integer

ChkProfileAccessNew = False

Set trnNode = xmlNewBranchProfile.documentElement.selectSingleNode("trn[@id='" & CStr(aTrnCD) & "']")
If Not (trnNode Is Nothing) Then aFlag = 1 Else aFlag = 0
ChkProfileAccessNew = (aFlag = 1)
If aFlag = 1 Then
    Set trnNode = xmlNewUserProfiles.documentElement.selectSingleNode("trn[@id='" & CStr(aTrnCD) & "']")
    If Not (trnNode Is Nothing) Then aFlag = 1 Else aFlag = 0
    ChkProfileAccessNew = (aFlag = 1)
End If

End Function

Public Function ChkAccess(aTrnCD As Long)
    ChkAccess = True
End Function

Public Function NVLDouble_(invalue, retValue As Double) As Double
    NVLDouble_ = retValue
    On Error GoTo NVLDoubleError
    NVLDouble_ = IIf((VarType(invalue) = vbNull) Or (VarType(invalue) = vbEmpty), retValue, invalue)
NVLDoubleError:
End Function

Public Function NVLInteger_(invalue, retValue As Integer) As Long
    NVLInteger_ = retValue
    On Error GoTo NVLIntegerError
    NVLInteger_ = IIf((VarType(invalue) = vbNull) Or (VarType(invalue) = vbEmpty), retValue, invalue)
NVLIntegerError:
End Function

Public Function NVLString_(invalue, retValue As String) As String
    NVLString_ = retValue
    On Error GoTo NVLStringError:
    NVLString_ = IIf((VarType(invalue) = vbNull) Or (VarType(invalue) = vbEmpty), retValue, invalue)
NVLStringError:
End Function

Public Function NVLBoolean_(invalue, retValue As Boolean) As Boolean
    NVLBoolean_ = retValue
    On Error GoTo NVLBooleanError
    NVLBoolean_ = IIf((VarType(invalue) = vbNull) Or (VarType(invalue) = vbEmpty), retValue, invalue)
NVLBooleanError:
End Function

Public Function NVLDate_(invalue, retValue As Date) As Date
    NVLDate_ = retValue
    On Error GoTo NVLDateError
    NVLDate_ = IIf((VarType(invalue) = vbNull) Or (VarType(invalue) = vbEmpty), retValue, invalue)
NVLDateError:
End Function

Public Function fnDisplayTotals(TotalsGrid As Object)
Dim nstr As String, astr As String, bstr As String, aValue As Double, bvalue As Double, i As Integer
TotalsGrid.FormatString = "<¡»—.    |>_______________”˝ÌÔÎÔ|_____|<–ÂÒÈ„Ò·ˆﬁ_________________________________"
        
    fnDisplayTotals = True
    GoTo ExitPoint
GenError:
    NBG_LOG_MsgBox "À‹ËÔÚ ÛÙÁÌ ¡Ì‹ÍÙÁÛÁ ¡ËÒÔÈÛÙ˛Ì... (¬2)", True, "À¡»œ”"
    fnDisplayTotals = False
ExitPoint:
On Error Resume Next
End Function

Public Function UpdateIRISTotals_(owner As Form) As Integer
UpdateIRISTotals_ = 1
End Function

Public Function ClearIRISTotals_(owner As Form) As Integer

ClearIRISTotals_ = 1
End Function

Public Function GetIRISTotal_(TotalName As String) As Double
GetIRISTotal_ = 0
End Function

Public Function GetBranchIRISTotal_(TotalName As String) As Double
GetBranchIRISTotal_ = 0
End Function

Public Function GetTotal_(TotalName As String) As Double
    GetTotal_ = 0
End Function

Public Function GetBranchTotal_(TotalName As String) As Double
    GetBranchTotal_ = 0
End Function

Public Function GetCurTotal_(TotalName As String, inCurrency As Integer) As Double
    GetCurTotal_ = 0
End Function

Public Function GetDBTotal_(TotalName As String) As Double
    GetDBTotal_ = 0
End Function

Public Function GetDBTotalTerm_(TotalName As String, term As String, adate As Date) As Double
    GetDBTotalTerm_ = 0
End Function

Public Function GetBranchDBTotal_(TotalName As String) As Double
    GetBranchDBTotal_ = 0
End Function

Public Function GetBranchCRTotal_(TotalName As String) As Double
    GetBranchCRTotal_ = 0
End Function

Public Function GetCurDBTotal_(TotalName As String, inCurrency As Integer, inTerminal As String) As Double
    GetCurDBTotal_ = 0
End Function

Public Function GetCRTotal_(TotalName As String) As Double
    GetCRTotal_ = 0
End Function

Public Function GetCRTotalTerm_(TotalName As String, term As String, adate As Date) As Double
    GetCRTotalTerm_ = 0
End Function

Public Function GetCurCRTotal_(TotalName As String, inCurrency As Integer, inTerminal As String) As Double
    GetCurCRTotal_ = 0
End Function

Public Function GetNextCur_(TotalName As String, inCurrency As Integer, inTerminal As String) As Integer
    GetNextCur_ = -1
End Function

Public Sub AddDBTotal_(TotalName As String, aValue As Double)

End Sub

Public Sub AddCurDBTotal_(TotalName As String, inCurrency As Integer, aValue As Double)
End Sub

Public Sub AddCRTotal_(TotalName As String, aValue As Double)
End Sub

Public Sub AddCurCRTotal_(TotalName As String, inCurrency As Integer, aValue As Double)
End Sub

Public Sub eJournalWrite(aDataLine As String)
    Dim astr As String, bstr As String, cmdStr As String, recno As Integer
    Dim RFlag As Boolean, SFlag As Boolean
    
    On Error GoTo error_handler
    astr = eJournalClearString(aDataLine)
    If Len(astr) > 0 Then
        If Right(astr, 1) = "`" Then astr = Trim(Left(astr, Len(astr) - 1))
    End If
    
   
    With GenWorkForm.vJournal
        If (Trim(LastTRNCode) <> CStr(cTRNCode)) Or (LastTRNNum <> cTRNNum) Then 'Or (LastKeyChanged = True) Then
            LastTRNCode = cTRNCode
            LastTRNNum = cTRNNum
            'SetLastKey
            
            .SelStart = Len(GenWorkForm.vJournal.Text): .SelLength = 0: .SelBold = True
            .SelText = vbCrLf & vbCrLf & "”ıÌ·ÎÎ·„ﬁ: " & _
                LastTRNCode & "   A/A: " & CStr(LastTRNNum) & "   ◊ÒﬁÛÙÁÚ: " & cUserName '& "   ≈„ÍÒÈÛÁ: " & LastChief
        
        End If
        .SelStart = Len(GenWorkForm.vJournal.Text): .SelLength = 0: .SelBold = False
        .SelText = vbCrLf & astr
        
        If Left(astr, 3) = "R:5" Then
            .SelStart = Len(GenWorkForm.vJournal.Text): .SelLength = 0: .SelBold = False
            .SelText = vbCrLf & "***« ≈–… œ…ÕŸÕ…¡ œÀœ À«—Ÿ»« ≈***"
        End If
        
    End With
    
    GoTo ExitPoint

error_handler:
    Call NBG_LOG_MsgBox("Error :" & error$, True)
ExitPoint:
End Sub

Public Function eJournalWriteAll(owner As Form, aDataLine As String) As Boolean
    Dim astr As String, cmdStr As String, bstr As String, recno As Integer
    Dim aUser As String
    Dim RFlag As Boolean, SFlag As Boolean
        
    eJournalWriteAll = False
    On Error GoTo error_handler
    astr = eJournalClearString(aDataLine)
    If Len(astr) > 0 Then
        If Right(astr, 1) = "`" Then astr = Trim(Left(astr, Len(astr) - 1))
    End If
    
    If (Trim(LastTRNCode) <> CStr(cTRNCode)) Or (LastTRNNum <> cTRNNum) Then 'Or (LastKeyChanged = True) Then
        LastTRNCode = cTRNCode
        LastTRNNum = cTRNNum
        'SetLastKey
        
        GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
        GenWorkForm.vJournal.SelLength = 0
        GenWorkForm.vJournal.SelBold = True
        GenWorkForm.vJournal.SelText = vbCrLf & vbCrLf & "”ıÌ·ÎÎ·„ﬁ: " & _
            LastTRNCode & "   A/A: " & CStr(LastTRNNum) & "   ◊ÒﬁÛÙÁÚ: " & cUserName
        
    End If
    GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
    GenWorkForm.vJournal.SelLength = 0
    GenWorkForm.vJournal.SelBold = False
    GenWorkForm.vJournal.SelText = vbCrLf & astr
    
        If Left(astr, 3) = "R:5" Then
            GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
            GenWorkForm.vJournal.SelLength = 0: GenWorkForm.vJournal.SelBold = False
            GenWorkForm.vJournal.SelText = vbCrLf & "***« ≈–… œ…ÕŸÕ…¡ œÀœ À«—Ÿ»« ≈***"
        End If
        
    GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
    GenWorkForm.vJournal.SelLength = 0
    eJournalWriteAll = True
    GoTo ExitPoint

error_handler:
    Call NBG_LOG_MsgBox("Error :" & error$, True)
    Resume ExitPoint
ExitPoint:
End Function

Public Function eJournalWriteFld(owner As Form, aFldNo As Integer, aFldTitle As String, aDataLine As String) As Boolean
    Dim astr As String, bstr As String, cmdStr As String, reccount As Integer, recno As Integer
    Dim aUser As String
    eJournalWriteFld = False
    On Error GoTo error_handler
    
    astr = eJournalClearString(aDataLine)
    If Len(astr) > 0 Then
        If Right(astr, 1) = "`" Then astr = Trim(Left(astr, Len(astr) - 1))
    End If
    If (Trim(astr) = "") Or (Trim(astr) = "R:") Then eJournalWriteFld = True: Exit Function
    aFldTitle = eJournalClearString(aFldTitle)
        
    If (Trim(LastTRNCode) <> CStr(cTRNCode)) Or (LastTRNNum <> cTRNNum) Then
        LastTRNCode = cTRNCode
        LastTRNNum = cTRNNum
        
        GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
        GenWorkForm.vJournal.SelLength = 0
        GenWorkForm.vJournal.SelBold = True
        GenWorkForm.vJournal.SelText = vbCrLf & vbCrLf & "”ıÌ·ÎÎ·„ﬁ: " & _
            LastTRNCode & "   A/A: " & CStr(LastTRNNum) & "   ◊ÒﬁÛÙÁÚ: " & cUserName
        
    End If
    GenWorkForm.vJournal.SelLength = 0
    GenWorkForm.vJournal.SelBold = True
    GenWorkForm.vJournal.SelBold = False
    GenWorkForm.vJournal.SelText = vbCrLf & aFldTitle & astr
    
    GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
    GenWorkForm.vJournal.SelLength = 0
    eJournalWriteFld = True
    GoTo ExitPoint
error_handler:
    Call NBG_LOG_MsgBox("Error :" & error$, True)
    Resume ExitPoint
ExitPoint:
End Function

Public Function eJournalWriteFinal(owner As Form) As Boolean
    Dim cmdStr As String, reccount As Integer, recno As Integer, aBonus As Integer
    Dim aUser As String
    eJournalWriteFinal = False
    On Error GoTo error_handler
    
    If (Trim(LastTRNCode) <> CStr(cTRNCode)) Or (LastTRNNum <> cTRNNum) Then
        LastTRNCode = cTRNCode
        LastTRNNum = cTRNNum
        
        GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
        GenWorkForm.vJournal.SelLength = 0
        GenWorkForm.vJournal.SelBold = True
        GenWorkForm.vJournal.SelText = vbCrLf & vbCrLf & "”ıÌ·ÎÎ·„ﬁ: " & _
            LastTRNCode & "   A/A: " & CStr(LastTRNNum) & "   ◊ÒﬁÛÙÁÚ: " & cUserName
        
    End If
    GenWorkForm.vJournal.SelLength = 0
    GenWorkForm.vJournal.SelBold = True
    GenWorkForm.vJournal.SelBold = False
    GenWorkForm.vJournal.SelText = vbCrLf & "« ”’Õ¡ÀÀ¡√« " & CStr(cTRNCode) & " œÀœ À«—Ÿ»« ≈"
    
    GenWorkForm.vJournal.SelStart = Len(GenWorkForm.vJournal.Text)
    GenWorkForm.vJournal.SelLength = 0
    eJournalWriteFinal = True
    GoTo ExitPoint

error_handler:
    Call NBG_LOG_MsgBox("Error :" & error$, True)
    Resume ExitPoint
ExitPoint:
End Function

Public Function BuildCRAStruct(owner As Buffers, StructureName, Alias, Optional LastLevel As Boolean) As Boolean
    If SkipCRAUse Then
        MsgBox "ƒÂÌ ıÔÛÙÁÒﬂÊÂÙ·È Á ÎÂÈÙÔıÒ„ﬂ· ·¸ ÙÔ Û˝ÛÙÁÏ·. (BuildAppStructFromDB)", vbCritical: Exit Function
    End If
    If IsMissing(LastLevel) Then LastLevel = True
    BuildCRAStruct = False
    On Error GoTo GenError
    If owner.Exists(Alias) Then Exit Function
    Dim aDesc As String, bDesc As String
    aDesc = xmlCRAStructures.documentElement.selectSingleNode(StructureName).Text
    Dim aPos As Integer, neststruct As String
    aDesc = UCase(aDesc): bDesc = aDesc
    aPos = InStr(1, aDesc, "STRUCT ")
    While aPos > 0
        aDesc = Right(aDesc, Len(aDesc) - aPos + 1)
        aDesc = Right(aDesc, Len(aDesc) - 7)
        aPos = InStr(1, aDesc, " ")
        neststruct = Trim(Left(aDesc, aPos))
        If Not owner.Exists(neststruct) Then BuildCRAStruct owner, neststruct, neststruct, False
        aPos = InStr(1, aDesc, "STRUCT ")
    Wend
    owner.DefineBuffer CStr(Alias), CStr(StructureName), bDesc, CStr(StructureName), LastLevel
    BuildCRAStruct = True
    Exit Function
GenError:
    MsgBox "–Ò¸‚ÎÁÏ· ÛÙÁ ‰ﬁÎ˘ÛÁ ÙÁÚ ‰ÔÏﬁÚ: " & Alias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "À¡»œ”"
End Function

Public Function BuildCRAAppStruct(StructureName, Alias, Optional LastLevel As Boolean) As Boolean
    If IsMissing(LastLevel) Then LastLevel = True
    BuildCRAAppStruct = BuildCRAStruct(GenWorkForm.AppBuffers, StructureName, Alias, LastLevel)
End Function

Public Function BuildIRISStruct(owner As Buffers, StructureName, Alias, Optional LastLevel As Boolean) As Boolean
Dim aDesc As String, bDesc As String, astructid As String
Dim aPos As Integer, neststruct As String
    If IsMissing(LastLevel) Then LastLevel = True
    BuildIRISStruct = False
    On Error GoTo GenError
    If owner.Exists(Alias) Then Exit Function
    
    aDesc = "": astructid = ""
    If Not xmlIRISStructuresUpdate Is Nothing Then
        If Not xmlIRISStructuresUpdate.documentElement Is Nothing Then
            If Not xmlIRISStructuresUpdate.documentElement.selectSingleNode(StructureName) Is Nothing Then
                aDesc = xmlIRISStructuresUpdate.documentElement.selectSingleNode(StructureName).Text
                astructid = xmlIRISStructuresUpdate.documentElement.selectSingleNode(StructureName).Attributes(0).Text
            End If
        End If
    End If
    If aDesc = "" And astructid = "" Then
        aDesc = xmlIRISStructures.documentElement.selectSingleNode(StructureName).Text
        astructid = xmlIRISStructures.documentElement.selectSingleNode(StructureName).Attributes(0).Text
    End If
    
    aDesc = UCase(aDesc): bDesc = aDesc
    aPos = InStr(1, aDesc, "STRUCT ")
    While aPos > 0
        aDesc = Right(aDesc, Len(aDesc) - aPos + 1)
        aDesc = Right(aDesc, Len(aDesc) - 7)
        aPos = InStr(1, aDesc, " ")
        neststruct = Trim(Left(aDesc, aPos))
        If Not owner.Exists(neststruct) Then BuildIRISStruct owner, neststruct, neststruct, False
        aPos = InStr(1, aDesc, "STRUCT ")
    Wend
    owner.DefineBuffer CStr(Alias), astructid, bDesc, CStr(StructureName), LastLevel
    BuildIRISStruct = True
    Exit Function
GenError:
    MsgBox "–Ò¸‚ÎÁÏ· ÛÙÁ ‰ﬁÎ˘ÛÁ ÙÁÚ ‰ÔÏﬁÚ: " & Alias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "À¡»œ”"
End Function

Public Function BuildIRISAppStruct(StructureName, Alias, Optional LastLevel As Boolean) As Boolean
    If IsMissing(LastLevel) Then LastLevel = True
    BuildIRISAppStruct = BuildIRISStruct(GenWorkForm.AppBuffers, StructureName, Alias, LastLevel)
End Function

Public Function BuildComArea(owner As Buffers, StructureName As String, filename As String) As Boolean
    BuildComArea = False
On Error GoTo GenError
    If Not owner.Exists(StructureName) Then
        Dim structurecode As String
        structurecode = ""
        
        On Error Resume Next
        Close #1
        On Error GoTo FileNotFoundError
        Dim s As String
        Open ReadDir & ComAreaDir & filename & ".txt" For Input As #1
        Do While Not Eof(1)
            Line Input #1, s
            structurecode = structurecode & s
        Loop
        Close #1
        
        Dim res As Buffer
        Set res = owner.DefineComArea(structurecode, StructureName, False)
        
        If res Is Nothing Then
            BuildComArea = False
            Exit Function
        Else
        
        End If
    End If
    BuildComArea = True
    Exit Function
FileNotFoundError:
    MsgBox "ƒÂÌ ‚Ò›ËÁÍÂ ÙÔ ·Ò˜ÂﬂÔ: " & ComAreaDir & filename & ".txt" & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "À¡»œ”"
    Exit Function
GenError:
    MsgBox "–Ò¸‚ÎÁÏ· ÛÙÁ ‰ﬁÎ˘ÛÁ ÙÁÚ ‰ÔÏﬁÚ: " & StructureName & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "À¡»œ”"
End Function

Public Function GetBankName_(BankCode As Integer) As String
Dim abank As String
On Error GoTo GenError
      Select Case Right("000" & BankCode, 2)
      Case 11: abank = "≈–…‘¡√≈” ≈»Õ… «” ‘—¡–≈∆¡”"
      Case 12: abank = "≈–…‘¡√≈” ≈Ã–œ—… «” ‘—¡–≈∆¡”"
      Case 13: abank = "≈–…‘¡√≈” …œÕ… «” ‘—¡–≈∆¡”"
      Case 14: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡” –…”‘≈Ÿ”"
      Case 15: abank = "≈–…‘¡√≈” √≈Õ… «” ‘—¡–≈∆¡”"
      Case 16: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡” ¡‘‘… «”"
      Case 17: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡” –≈…—¡…Ÿ”"
      Case 18: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡” ¡»«ÕŸÕ"
      Case 19: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡”  —«‘«”"
      Case 20: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡” ≈—√¡”…¡”"
      Case 22: abank = "≈–…‘¡√≈” ‘—¡–.  ≈Õ‘—. ≈ÀÀ¡ƒ¡”"
      Case 25: abank = "≈–…‘¡√≈” TELESIS"
      Case 26: abank = "≈–…‘¡√≈” EFG EUROBANK-ERGASIAS"
      Case 27: abank = "≈–…‘¡√≈” EUROBANK"
      Case 28: abank = "≈–…‘¡√≈” MARFIN EGNATIA BANK"
      Case 31: abank = "≈–…‘¡√≈” ≈’—Ÿ–¡⁄ «”-À¡⁄ «”"
      Case 32: abank = "≈ÀÀ«Õ… « ‘—¡–≈∆¡ À‘ƒ"
      Case 34: abank = "≈–…‘¡√≈” ≈–≈Õƒ’‘… «” ‘—¡–≈∆¡” ≈ÀÀ¡ƒœ” ¡.≈."
      Case 36: abank = "≈ÀÀ«Õ… « ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” ƒ’‘.Ã¡ ≈ƒœÕ…¡” ”’Õ.–.≈. À‘ƒ"
      Case 37: abank = "≈–…‘¡√≈” Õ≈¡” PROTON"
      Case 38: abank = "≈–…‘¡√≈” NOVABANK"
      Case 41: abank = "≈–…‘¡√≈” ≈‘¬¡"
      Case 43: abank = "≈–…‘¡√≈” ¡√—œ‘… «” ‘—¡–≈∆¡”"
      Case 47: abank = "≈–…‘¡√≈” ASPIS BANK"
      Case 49: abank = "≈–…‘¡√≈” –¡Õ≈ÀÀ«Õ…¡” ‘—¡–≈∆¡”"
      Case 54: abank = "≈–…‘¡√≈” PROBANK"
      Case 55: abank = "≈–…‘¡√≈” FBB-FIRST BUSINESS BANK"
      Case 60: abank = "≈–…‘¡√≈” THE ROYAL BANK OF SCOTLAND N.V."
      Case 62: abank = "≈–…‘¡√≈” ANZ-GRINDLAYS"
      Case 63: abank = "≈–…‘¡√≈” NATIONAL WESTMINSTER"
      Case 64: abank = "≈–…‘¡√≈” THE ROYAL BANK OF SCOTLAND PLC"
      Case 65: abank = "≈–…‘¡√≈” BARCLAYS BANK"
      Case 67: abank = "≈–…‘¡√≈” SOCIETE CENERALE"
      Case 68: abank = "≈–…‘¡√≈” CREDIT COMMERCIAL FR"
      Case 69: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” ◊¡Õ…ŸÕ"
      Case 70: abank = "≈–…‘¡√≈” BANK NATION.DE PARIS"
      Case 71: abank = "≈–…‘¡√≈” HSBC"
      Case 72: abank = "≈–…‘¡√≈” UNICREDIT BANK AG"
      Case 73: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡”  ’–—œ’"
      Case 74: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” À¡Ã…¡”"
      Case 75: abank = "≈–…‘¡√≈” ‘—¡–≈∆¡ «–≈…—œ’ ”’Õ.–.≈."
      Case 77: abank = "≈–…‘¡√≈” ¡◊¡⁄ «” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡”"
      Case 78: abank = "≈–…‘¡√≈” ING-BANK"
      Case 79: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ƒŸƒ≈ ¡Õ«”œ’"
      Case 80: abank = "≈–…‘¡√≈” AMERICAN EXPRESS"
      Case 81: abank = "≈–…‘¡√≈” BANK OF AMERICA"
      Case 84: abank = "≈–…‘¡√≈” CITIBANK"
      Case 87: abank = "≈–…‘¡√≈” –¡√ —«‘…¡” ”’Õ≈‘¡…—…”‘… «”"
      Case 88: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” ≈¬—œ’"
      Case 89: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «”  ¡—ƒ…‘”¡” ”’Õ.–.≈."
      Case 91: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” »≈””¡À…¡” ”’Õ.–≈."
      Case 92: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” –≈Àœ–œÕÕ«”œ’"
      Case 95: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” ƒ—¡Ã¡”"
      Case 96: abank = "≈–…‘¡√≈” ‘¡◊’ƒ—œÃ… œ’ ‘¡Ã…≈’‘«—…œ’"
      Case 97: abank = "≈–…‘¡√≈” ‘¡Ã≈…œ’ –¡—¡ ¡‘.&ƒ¡Õ≈…ŸÕ"
      Case 98: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” À≈”¬œ’ À«ÃÕœ’"
      Case 99: abank = "≈–…‘¡√≈” ”’Õ≈‘¡…—…”‘… «” ‘—¡–≈∆¡” Õ. ”≈——ŸÕ"
      Case 107: abank = "≈–…‘¡√≈” ≈–…‘¡√≈” GREEK BRANCH OF CLOSED JOINT STOCK COMPANY COMMERCIAL BANK KEDR"
      Case 109: abank = "≈–…‘¡√≈” T.C. ZIRAAT BANKASI A.S."
      Case Else: abank = ""
      End Select
      GetBankName_ = abank
      Exit Function
GenError:
    GetBankName_ = ""
End Function

Public Function ISOTOCURR_(ByVal inUnit As String) As String
     Dim astr, bstr As String, aPos As Integer
     ISOTOCURR_ = ""
     On Error GoTo ISOTOCURRError
     astr = "AED,All,ARP,ATS,AUD,AWG,BEF,BGL,BHD,BRC,CAD,CFA,CHF,CHP,CSK,CYP,DEM,DKK,SUR,GRD,ZRZ,ZAK,USD,JPY,INR,JOD,IQD,IRR,IEP,ESC,ISK,ESP,ILS,ITL,XAF,QAR,RMY,KRW,WKR,KWD,LBR,LKR,LUF,LYD,MAD,MTP,NGN,NLG,NOK,NZD,OMR,PLZ,PTE,ROL,SAR,SGD,SEK,SYR,TWD,TRL,TND,FIM,HKD,FRF,GBP,XEU,PKR,EUR"
     bstr = "062,027,036,021,049,090,059,019,065,029,010,091,008,092,015,032,005,014,045,001,051,093,002,043,037,094,089,040,057,085,052,018,031,004,088,066,098,038,095,025,030,096,068,039,046,056,042,017,013,071,053,022,023,024,060,067,012,061,097,007,044,035,064,003,050,070,072,070"
     aPos = InStr(1, astr, inUnit)
     If aPos > 0 Then ISOTOCURR_ = Mid(bstr, aPos, 3):
     Exit Function
ISOTOCURRError:
End Function

Public Function CURRTOISO_(ByVal inUnit As String) As String
     Dim astr, bstr As String, aPos As Integer
     CURRTOISO_ = ""
     On Error GoTo CURRTOISOError
     astr = "AED,All,ARP,ATS,AUD,AWG,BEF,BGL,BHD,BRC,CAD,CFA,CHF,CHP,CSK,CYP,DEM,DKK,SUR,GRD,ZRZ,ZAK,USD,JPY,INR,JOD,IQD,IRR,IEP,ESC,ISK,ESP,ILS,ITL,XAF,QAR,RMY,KRW,WKR,KWD,LBR,LKR,LUF,LYD,MAD,MTP,NGN,NLG,NOK,NZD,OMR,PLZ,PTE,ROL,SAR,SGD,SEK,SYR,TWD,TRL,TND,FIM,HKD,FRF,GBP,XEU,PKR,EUR"
     bstr = "062,027,036,021,049,090,059,019,065,029,010,091,008,092,015,032,005,014,045,001,051,093,002,043,037,094,089,040,057,085,052,018,031,004,088,066,098,038,095,025,030,096,068,039,046,056,042,017,013,071,053,022,023,024,060,067,012,061,097,007,044,035,064,003,050,070,072,070"
     aPos = InStr(1, bstr, inUnit)
     If aPos > 0 Then CURRTOISO_ = Mid(astr, aPos, 3):
     Exit Function
CURRTOISOError:
End Function

Public Property Get AppVariable_(inName As String)
    Dim i As Integer
    AppVariable_ = ""
    For i = 1 To GenWorkForm.AppVariables.count
        If UCase(GenWorkForm.AppVariables.Item(i).name) = UCase(inName) Then
            AppVariable_ = GenWorkForm.AppVariables.Item(i).value
            Exit Property
        End If
    Next i
End Property

Public Property Let AppVariable_(inName As String, invalue)
    Dim i As Integer
    For i = 1 To GenWorkForm.AppVariables.count
        If UCase(GenWorkForm.AppVariables.Item(i).name) = UCase(inName) Then
            GenWorkForm.AppVariables.Item(i).value = invalue
            Exit Property
        End If
    Next i
    Dim aVariable As New VariableEntry
    aVariable.name = inName
    aVariable.value = invalue
    GenWorkForm.AppVariables.add aVariable
End Property

Public Function GetInCur_() As String
'À…”‘¡ Ã≈ IN ÕœÃ…”Ã¡‘¡
    GetInCur_ = "001,003,004,005,017,018,021,023,035,057,059,070,"
End Function

Public Function FormatIBAN_(IBAN As String) As String
Dim iban4, ibanpart As String
Dim i As Integer

iban4 = ""
For i = 1 To ((Len(IBAN) / 4) + 1)
  ibanpart = Mid(IBAN, 1, 4)
  iban4 = iban4 & ibanpart & " "
  IBAN = Mid(IBAN, 5)
Next i

FormatIBAN_ = iban4
End Function

Public Function EUROText_() As String
    EUROText_ = "…”œ‘…Ãœ ”≈ ≈’—Ÿ:           "
End Function

Public Function EUROText2002_() As String
    If cVersion >= 20030101 Then
        EUROText2002_ = ""
    Else
        EUROText2002_ = IIf(cVersion >= 20020101, "…”œ‘…Ãœ ”≈ ƒ—◊.:           ", "…”œ‘…Ãœ ”≈ ≈’—Ÿ:           ")
    End If
End Function

Public Function GRDText_() As String
    GRDText_ = "…”œ‘…Ãœ ”≈ ƒ—◊:           "
End Function

Public Function EURORate_() As Double
    EURORate_ = 340.75
End Function

Public Function EUROAmount_(inAmount) As String
    EUROAmount_ = StrPad_(FormatNumber(CStr(CDbl(inAmount) / 34075), 2), 20, " ", "L")
End Function

Public Function EUROAmount2002_(inAmount) As String
Dim aval As Double
    If cVersion >= 20030101 Then
        EUROAmount2002_ = ""
    ElseIf cVersion >= 20020101 Then
        aval = Round(CDbl(inAmount) * 3.4075, 0)
        EUROAmount2002_ = StrPad_(FormatNumber(aval, 2), 20, " ", "L")
    Else
        EUROAmount2002_ = StrPad_(FormatNumber(CStr(CDbl(inAmount) / 34075), 2), 20, " ", "L")
    End If
End Function

Public Function GRDAmount_(inAmount) As String
    GRDAmount_ = StrPad_(FormatNumber(CStr(Round(CDbl(inAmount) * 3.4075)), 0), 20, " ", "L")
End Function

Public Function EUROAmount5_(inAmount) As String
'Format ÔÛÔı „È· ÙÔ Ò¸„Ò·ÏÏ· 5
    EUROAmount5_ = StrPad_(FormatNumber(CStr(CDbl(inAmount) / 100), 2), 20, " ", "L")
End Function

Public Function AddFTFilaRecordset_(inName, inFilter, Optional inSort) As ADODB.Recordset
    Dim ars As New ADODB.Recordset, aRSRow As RecordsetEntry
    On Error GoTo invalidRS
    Set ars = rsFTFila.Clone
    ars.Filter = inFilter
    If Not IsMissing(inSort) Then
        ars.Sort = inSort
    End If
    Set AddFTFilaRecordset_ = ars
    Set aRSRow = New RecordsetEntry
    Set aRSRow.rs = ars
    aRSRow.name = inName
    GenWorkForm.AppRS.add aRSRow
    Exit Function
invalidRS:
    Dim astr As String, aMsg As Long
    astr = Err.description: aMsg = Err.number
    Set AddFTFilaRecordset_ = Nothing
    Set aRSRow = New RecordsetEntry
    Set aRSRow.rs = Nothing
    aRSRow.name = inName
    GenWorkForm.AppRS.add aRSRow
    MsgBox "À¡»œ” (" & aMsg & "). " & astr, vbCritical, "On Line ≈ˆ·ÒÏÔ„ﬁ"
End Function

Public Function AppRecordsetByName_(inName) As ADODB.Recordset
    Dim i As Integer
    Set AppRecordsetByName_ = Nothing
    For i = GenWorkForm.AppRS.count To 1 Step -1
        If UCase(GenWorkForm.AppRS.Item(i).name) = UCase(inName) Then
            Set AppRecordsetByName_ = GenWorkForm.AppRS.Item(i).rs
            Exit For
        End If
    Next i
End Function

Public Function AppRSEntryByName_(inName) As RecordsetEntry
    Dim i As Integer
    Set AppRSEntryByName_ = Nothing
    For i = GenWorkForm.AppRS.count To 1 Step -1
        If UCase(GenWorkForm.AppRS.Item(i).name) = UCase(inName) Then
            Set AppRSEntryByName_ = GenWorkForm.AppRS.Item(i)
            Exit For
        End If
    Next i
End Function

Public Function AppRecordsetByIndex_(inIdx) As ADODB.Recordset
    Dim i As Integer
    Set AppRecordsetByIndex_ = GenWorkForm.AppRS.Item(inIdx).rs
End Function

Public Sub FreeAppRecordset_(inName)
    Dim i As Integer
    For i = 1 To GenWorkForm.AppRS.count
        If UCase(GenWorkForm.AppRS.Item(i).name) = UCase(inName) Then
            Set GenWorkForm.AppRS.Item(i).rs = Nothing
            GenWorkForm.AppRS.Remove (i)
            Exit For
        End If
    Next i
End Sub

Public Sub FreeAppCRecordset_(inName)
    Dim i As Integer
    For i = 1 To GenWorkForm.AppCRS.count
        If UCase(GenWorkForm.AppCRS.Item(i).name) = UCase(inName) Then
            Set GenWorkForm.AppCRS.Item(i).Recordset = Nothing
            GenWorkForm.AppCRS.Remove (i)
            Exit For
        End If
    Next i
End Sub

Public Function BuildIRISErrorMessage_(OutputView) As String
    Dim i As Integer, aCode As Long, amessage As String, aMsgString As String
    BuildIRISErrorMessage_ = ""
    If Not (OutputView.xmlNode("STD_TRN_MSJ_PARM_V") Is Nothing) Then
        For i = 1 To 5
            aCode = OutputView.ByName("STD_TRN_MSJ_PARM_V").ByName("TEXT_CODE", i).value
            amessage = OutputView.ByName("STD_TRN_MSJ_PARM_V").ByName("TEXT_ARG1", i).value
            If aCode <> 0 Then
                rsIRISErrors.Filter = "VALUE_IMP_NAME=" & CStr(aCode)
                If rsIRISErrors.RecordCount > 0 Then
                    aMsgString = rsIRISErrors!Data
                    aMsgString = Replace(aMsgString, "XX", amessage)
                    BuildIRISErrorMessage_ = BuildIRISErrorMessage_ & aMsgString & vbCrLf
                End If
            End If
        Next i
    ElseIf Not (OutputView.xmlNode("STD_MSJ_PARM_V") Is Nothing) Then
        For i = 1 To 5
            aCode = OutputView.ByName("STD_MSJ_PARM_V").ByName("TEXT_CODE", i).value
            amessage = OutputView.ByName("STD_MSJ_PARM_V").ByName("TEXT_ARG1", i).value
            If aCode <> 0 Then
                rsIRISErrors.Filter = "VALUE_IMP_NAME=" & CStr(aCode)
                If rsIRISErrors.RecordCount > 0 Then
                    aMsgString = rsIRISErrors!Data
                    aMsgString = Replace(aMsgString, "XX", amessage)
                    BuildIRISErrorMessage_ = BuildIRISErrorMessage_ & aMsgString & vbCrLf
                End If
            End If
        Next i
    End If
End Function

Public Function ChkIRISOutput_(aBuffer, Optional looseChk) As Boolean
Dim i As Integer, acounter As Integer, aCode, amessage As String
    If IsMissing(looseChk) Then looseChk = False
    ChkIRISOutput_ = False
    
    If Not looseChk Then
        If aBuffer.ByName("RTRN_CD").value <> 1 Then
            amessage = BuildIRISErrorMessage_(aBuffer)
            If amessage <> "" Then
                LogMsgbox amessage, vbCritical, "À¡»œ”"
            Else
                LogMsgbox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
            End If
'            Set aForm = New IRISErrFrm
'            Set aForm.OutputView = aBuffer
'            aForm.Show vbModal
        Else
            ChkIRISOutput_ = True
        End If
    Else
        For i = 1 To 5
            aCode = aBuffer.ByName("STD_TRN_MSJ_PARM_V").ByName("TEXT_CODE", i).value
            If aCode <> 0 Then acounter = acounter + 1
        Next i
        If acounter = 0 Then
'            If aBuffer.ByName("RTRN_CD").Value <> 1 Then ValidationError = "ƒ≈Õ ¬—≈»« ¡Õ ”‘œ…◊≈…¡....": sbWriteStatusMessage ValidationError
            
            ChkIRISOutput_ = True
        Else
            amessage = BuildIRISErrorMessage_(aBuffer)
            If amessage <> "" Then
                LogMsgbox amessage, vbCritical, "À¡»œ”"
            Else
                LogMsgbox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
            End If
            
'            Set aForm = New IRISErrFrm
'            Set aForm.OutputView = aBuffer
'            aForm.Show vbModal
        End If
    End If
End Function




Public Function ChkIRISOutputSkip7_(aBuffer, Optional looseChk) As Boolean
Dim i As Integer, acounter As Integer, aCode, amessage As String
    If IsMissing(looseChk) Then looseChk = False
    ChkIRISOutputSkip7_ = False
    
    If Not looseChk Then
        If aBuffer.ByName("RTRN_CD").value <> 1 And aBuffer.ByName("RTRN_CD").value <> 7 Then
            amessage = BuildIRISErrorMessage_(aBuffer)
            If amessage <> "" Then
                LogMsgbox amessage, vbCritical, "À¡»œ”"
            Else
                LogMsgbox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
            End If
'            Set aForm = New IRISErrFrm
'            Set aForm.OutputView = aBuffer
'            aForm.Show vbModal
        Else
            ChkIRISOutputSkip7_ = True
        End If
    Else
        For i = 1 To 5
            aCode = aBuffer.ByName("STD_TRN_MSJ_PARM_V").ByName("TEXT_CODE", i).value
            If aCode <> 0 Then acounter = acounter + 1
        Next i
        If acounter = 0 Then
'            If aBuffer.ByName("RTRN_CD").Value <> 1 Then ValidationError = "ƒ≈Õ ¬—≈»« ¡Õ ”‘œ…◊≈…¡....": sbWriteStatusMessage ValidationError
            
            ChkIRISOutputSkip7_ = True
        Else
            amessage = BuildIRISErrorMessage_(aBuffer)
            If amessage <> "" Then
                LogMsgbox amessage, vbCritical, "À¡»œ”"
            Else
                LogMsgbox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
            End If
            
'            Set aForm = New IRISErrFrm
'            Set aForm.OutputView = aBuffer
'            aForm.Show vbModal
        End If
    End If
End Function

Public Function ChkCRA2Output_(aBuffer, Optional looseChk) As Boolean
Dim i As Integer, acounter As Integer, aCode, amessage As String
    If IsMissing(looseChk) Then looseChk = False
    ChkCRA2Output_ = False
    
    If Not looseChk Then
        If aBuffer.ByName("NBG_STD_ERR_VIEW", 1).ByName("C_RSLT", 1).value <> 1 Then
        
            amessage = "À‹ËÔÚ: " & aBuffer.ByName("NBG_STD_ERR_VIEW", 1).ByName("C_RSLT_ERRNO", 1).value & _
                       " " & aBuffer.ByName("NBG_STD_ERR_VIEW", 1).ByName("C_RSLT_PGM", 1).value & vbCrLf & _
                       aBuffer.ByName("NBG_STD_ERR_VIEW", 1).ByName("C_RSLT_TEXT", 1).value
            If amessage <> "" Then
                LogMsgbox amessage, vbCritical, "À¡»œ”"
            Else
                LogMsgbox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
            End If
        Else
            ChkCRA2Output_ = True
        End If
'    Else
'        For i = 1 To 5
'            aCode = aBuffer.ByName("STD_TRN_MSJ_PARM_V").ByName("TEXT_CODE", i).value
'            If aCode <> 0 Then acounter = acounter + 1
'        Next i
'        If acounter = 0 Then
'            ChkIRISOutput_ = True
'        Else
'            aMessage = BuildIRISErrorMessage_(aBuffer)
'            If aMessage <> "" Then
'                MsgBox aMessage, vbCritical, "À¡»œ”"
'            Else
'                MsgBox "À¡»œ”  ¡‘¡ ‘«Õ œÀœ À«—Ÿ”« ‘«” À≈…‘œ’—√…¡”", vbCritical, "À¡»œ”"
'            End If
'        End If
    End If
End Function

Public Function UpdateTrnNum_()
    cTRNNum = cTRNNum + 1: UpdateParams
End Function
Public Function gFormatType_(FormatString, inParams)
    Dim apctpos As Integer, bpctpos As Integer
    Dim CopyString As String
    Dim typeStr As String
    Dim astr As String
    Dim fieldTypeNode As IXMLDOMElement
    Dim displayMaskAttr As IXMLDOMAttribute
    CopyString = FormatString
    On Error GoTo ErrorPos
    apctpos = InStr(1, CopyString, "%", vbTextCompare)
    If (apctpos > 0 And apctpos < Len(CopyString)) Then
        bpctpos = InStr(apctpos + 1, CopyString, "%", vbTextCompare)
    Else
        bpctpos = 0
    End If
    If (bpctpos > 2) Then
       typeStr = Mid(CopyString, bpctpos - 2, 2)
       If (typeStr = "ST" Or typeStr = "SR" Or typeStr = "FS" Or typeStr = "FD" Or typeStr = "RP") Then
            gFormatType_ = gFormat_(FormatString, inParams)
       Else
           typeStr = Mid(CopyString, apctpos + 1, bpctpos - apctpos - 1)
           If (UCase(typeStr) = UCase("Account2CD")) Then
              astr = inParams(0)
              If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
              If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
           ElseIf (UCase(typeStr) = UCase("Account1CD")) Then
              astr = inParams(0)
              If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
              If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
              If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
              astr = astr & CalcCd2_(Left(astr, 10))
           ElseIf (UCase(typeStr) = UCase("Account0CD")) Then
              astr = inParams(0)
              If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
              If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
              astr = StrPad_(astr, 9, "0", "L")
              astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
              astr = astr & CalcCd2_(Left(astr, 10))
           ElseIf (UCase(typeStr) = UCase("I_IP")) Then
              astr = inParams(0)
              astr = astr & CalcCd1_(astr, 9)
           ElseIf (UCase(typeStr) = UCase("Date")) Then
              astr = inParams(0)
              If (IsDate(astr)) Then
                astr = format(astr, "DD/MM/YYYY")
              Else
                If Len(astr) = 8 Then
                    astr = Mid(astr, 1, 2) & "/" & Mid(astr, 3, 2) & "/" & Mid(astr, 5, 4)
                End If
              End If
           End If
           Set fieldTypeNode = L2ModelFile.documentElement.selectSingleNode("//datatype[@name='" & UCase(typeStr) & "']")
           If Not (fieldTypeNode Is Nothing) Then
             Set displayMaskAttr = fieldTypeNode.Attributes.getNamedItem("displaymask")
             If Not (displayMaskAttr Is Nothing And astr = "") Then
                gFormatType_ = format(astr, displayMaskAttr.Text)
             Else
                gFormatType_ = astr
             End If
           Else
             gFormatType_ = astr
           End If
       End If
    Else
        gFormatType_ = ""
    End If
    Exit Function
ErrorPos:
    gFormatType_ = ""
End Function
Public Function gFormat_(FormatString, inParams)
' 123456789012
' %-nnn.nnnFD%
' %nnnST%
' %AAAAAAAFS%

' %-nnn.nnnUD%
' %-nnnFI%
' %-nnnUI%
' %10FD%
' %8UD%
' %8FD%
' %6UD%
    Dim CopyString As String, aSizeString As String, allSize As Integer, decimalsize As Integer, paramNo As Integer, NewPart As String, ResultString As String, astr As String
    Dim apctpos As Integer, bpctpos As Integer, apointpos As Integer, asignflag As Boolean, aLeadZero As Boolean, aType As String
    paramNo = 0: CopyString = FormatString: ResultString = ""
    
    Dim temp As String
    temp = ""
    Do
    
        apctpos = InStr(1, CopyString, "%", vbTextCompare)
        If (apctpos > 0 And apctpos < Len(CopyString)) Then bpctpos = InStr(apctpos + 1, CopyString, "%", vbTextCompare) Else bpctpos = 0
'        If Len(CopyString) >= 4 Then
'            if left
'        End If
        If apctpos * bpctpos > 0 And bpctpos > apctpos + 3 Then
            asignflag = (Mid(CopyString, apctpos + 1, 1) = "-")
            If asignflag And bpctpos <= apctpos + 4 Then
                apctpos = 0: bpctpos = 0
            Else
                aType = UCase(Mid(CopyString, bpctpos - 2, 2))
                aSizeString = Mid(CopyString, apctpos + IIf(asignflag, 2, 1), bpctpos - apctpos - 1 - 2 - IIf(asignflag, 1, 0))
                apointpos = InStr(1, aSizeString, ".")
                If apointpos > 0 And (apointpos = 1 Or apointpos = Len(aSizeString)) Then
                    apctpos = 0: bpctpos = 0
                Else
                    If aType <> "FS" Then
                        aLeadZero = (Left(aSizeString, 1) = "0")
                        If apointpos = 0 Then allSize = CInt(aSizeString) Else allSize = CInt(Left(aSizeString, apointpos - 1))
                        If apointpos = 0 Then decimalsize = 0 Else decimalsize = CInt(Right(aSizeString, Len(aSizeString) - apointpos))
                    End If
                    Select Case aType
                        Case "FD": NewPart = Right(String(allSize, IIf(aLeadZero, "0", " ")) & FormatNumber(inParams(paramNo), decimalsize), allSize)
                        Case "ST": '”‘œ…◊…”« ¡—…”‘≈—¡
                                    If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    End If
                        Case "SR": '”‘œ…◊…”« ƒ≈Œ…¡
                                    If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Right(String(allSize, " ") & inParams(paramNo), allSize)
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then NewPart = String(allSize, " ") Else NewPart = Right(String(allSize, " ") & inParams(paramNo), allSize)
                                    End If
                        Case "PG":
                                    If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    End If
                        Case "GG":
                                    If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    End If
                                    
                        Case "FS": NewPart = Left(format(inParams(paramNo), aSizeString) & String(Len(aSizeString), " "), Len(aSizeString))
                        Case "RP": '≈–¡Õ¡À«ÿ« ◊¡—¡ ‘«—ŸÕ
                                    If IsObject(inParams(paramNo)) Then
                                        'NewPart = String(allSize, inParams(paramNo))
                                        NewPart = String(allSize, " ")
                                    Else
                                       
                                        If IsEmpty(inParams(paramNo)) Then
                                            NewPart = String(allSize, " ")
                                        Else
                                            While (Len(temp) < allSize)
                                                temp = temp + IIf(Trim(inParams(paramNo)) = "", " ", inParams(paramNo))
                                            Wend
                                            
                                            If Len(temp) > allSize Then temp = Left(temp, allSize)
                                            
                                        End If
                                        
                                        NewPart = temp
                                    End If
                        Case "IP": 'CRA I_IP
                                   Dim CRAstr As String
                                   If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then
                                            NewPart = String(allSize, " ")
                                        Else
                                            CRAstr = Right(String(9, "0") & inParams(paramNo), 9)
                                            CRAstr = CRAstr & CalcCd1_(CRAstr, 9)
                                            NewPart = CRAstr
                                        End If
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then
                                            NewPart = String(allSize, " ")
                                        Else
                                            CRAstr = Right(String(9, "0") & inParams(paramNo), 9)
                                            CRAstr = CRAstr & CalcCd1_(CRAstr, 9)
                                            NewPart = Left(CRAstr & String(allSize, " "), allSize)
                                        End If
                                    End If
                    End Select
                    
                    If apctpos > 1 Then astr = Left(CopyString, apctpos - 1) Else astr = ""
                    ResultString = ResultString & astr & NewPart
                    
                    If bpctpos < Len(CopyString) Then CopyString = Right(CopyString, Len(CopyString) - bpctpos) Else CopyString = ""
                    paramNo = paramNo + 1
                End If
            End If
        ElseIf apctpos * bpctpos > 0 And bpctpos = apctpos + 1 Then
            If apctpos > 1 Then astr = Left(CopyString, apctpos - 1) Else astr = ""
            ResultString = ResultString & astr & "%"
            If bpctpos < Len(CopyString) Then CopyString = Right(CopyString, Len(CopyString) - bpctpos) Else CopyString = ""
        ElseIf apctpos * bpctpos > 0 And bpctpos = apctpos + 3 Then
           
            If UCase(Mid(CopyString, apctpos, 4)) = "%ST%" Or UCase(Mid(CopyString, apctpos, 4)) = "%PG%" Or UCase(Mid(CopyString, apctpos, 4)) = "%GG%" Then
                If IsObject(inParams(paramNo)) Then
                    If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                Else
                    If IsEmpty(inParams(paramNo)) Then
                        NewPart = String(allSize, " ")
                    Else
                        allSize = Len(inParams(paramNo))
                        NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                    End If
                End If
                 
                If apctpos > 1 Then astr = Left(CopyString, apctpos - 1) Else astr = ""
                ResultString = ResultString & astr & NewPart
                
                If bpctpos < Len(CopyString) Then CopyString = Right(CopyString, Len(CopyString) - bpctpos) Else CopyString = ""
                paramNo = paramNo + 1
            Else
                Err.Raise 9999, "gFormat_", "ƒÂÌ ‚Ò›ËÁÍÂ ›Ì‰ÂÈÓÁ ST, FD, ..." & " ÛÙÔ " & FormatString

            End If
        Else
            apctpos = 0: bpctpos = 0
'            If Len(CopyString) > 0 Then
'                ResultString = ResultString & CopyString
'                CopyString = ""
'            End If


        End If
    
    Loop While (Len(CopyString) > 0 And apctpos * bpctpos > 0)
    
    gFormat_ = ResultString
End Function

Public Sub ShowIRISMessages_(inMessageView)
    Dim Counter As Integer
    Counter = 0
    If inMessageView.v2Value("STD_DEC_3") = 0 Then Exit Sub
    
    Dim aFrm As New IRISMsgFrm
    Set aFrm.MsgView = inMessageView
    aFrm.Show vbModal
    Set aFrm = Nothing
End Sub

Public Function ChkHPSComResult_(inRslt, inErrors) As Integer
    If inRslt > 0 Or inRslt < -1 Then
        LogMsgbox "À‹ËÔÚ: " & CStr(inRslt), vbOKOnly, "À¡»œ”"
        
        Load HPSErrForm: Set HPSErrForm.ErrBuffer = inErrors: HPSErrForm.Show vbModal
    ElseIf inRslt = -1 Then
        
    End If
    ChkHPSComResult_ = inRslt
    
End Function

Public Function ChkCRAOutput_(aBuffer, Optional ErrorView) As Boolean
Dim i As Integer, acounter As Integer, aCode, amessage As String
    ChkCRAOutput_ = False
    If Trim(ErrorView) = "" Then Set ErrorView = aBuffer.ByName("CUF_ERR_MSG_D")
    If aBuffer.ByName("C_RSLT").value <> 0 Then
        ChkCRAOutput_ = ChkHPSComResult_(aBuffer.ByName("C_RSLT").value, ErrorView)
    Else
        ChkCRAOutput_ = True
    End If
End Function

Public Function AddAppCRecordset_(inName, inCmd, inDBName, inVirtualDirectory, Optional inCursorType, Optional inLockType) As cADORecordset
    Dim cars As New cADORecordset
    cars.DBName = inDBName
    cars.VirtualDirectoryName = inVirtualDirectory
    cars.Open_ inCmd, inCursorType, inLockType
    cars.name = inName
    cars.RecordsetEntry.name = inName
    GenWorkForm.AppCRS.add cars
    Set AddAppCRecordset_ = cars
End Function

Public Function AppCRecordsetByName_(inName) As cADORecordset
    Dim i As Integer
    Set AppCRecordsetByName_ = Nothing
    For i = GenWorkForm.AppCRS.count To 1 Step -1
        If UCase(GenWorkForm.AppCRS.Item(i).name) = UCase(inName) Then
            Set AppCRecordsetByName_ = GenWorkForm.AppCRS.Item(i)
            Exit For
        End If
    Next i
End Function

Public Function AppCRSEntryByName_(inName) As RecordsetEntry
    Dim i As Integer
    Set AppCRSEntryByName_ = Nothing
    For i = GenWorkForm.AppCRS.count To 1 Step -1
        If UCase(GenWorkForm.AppCRS.Item(i).name) = UCase(inName) Then
            Set AppCRSEntryByName_ = GenWorkForm.AppCRS.Item(i).RecordsetEntry
            Exit For
        End If
    Next i
End Function

Public Sub SaveJournal()

End Sub
'    Dim filename As String
'
'    filename = NetworkHomeDir() + "\" + MachineName + "_" + CStr(Date) + ".rtf"
'    filename = Replace(filename, "/", "_")
'    GenWorkForm.vJournal.SaveFile filename
'End Sub

Public Sub LoadJournal()

End Sub
'    Dim filename As String
'    filename = NetworkHomeDir() + "\" + MachineName + "_" + CStr(Date) + ".rtf"
'    filename = Replace(filename, "/", "_")
'    If fnChkFileExistAbs(filename) Then
'        GenWorkForm.vJournal.LoadFile filename
'    End If
'End Sub

