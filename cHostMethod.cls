VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHostMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public owner As cXMLDocumentManager
Public content As IXMLDOMElement
Private System As String

Private Sub Class_Terminate()
    Set owner = Nothing
    Set content = Nothing
End Sub

'<call name="findcustomer">
'   <system name="CRA">
'   <rule name="ZADBX43" trncall="ACBTO">
'   <inputview system="cra" model="WCACU20B" name="WCACU20BL2USE">
'   <outputview system="cra" model="WCACU20B" name="WCACU20BL2USE">
'</trnlink>

Public Sub UpdateCRAView(inViewRoot, inDocRoot As IXMLDOMElement, inscount As Long)
Dim elm As IXMLDOMElement, selm As IXMLDOMElement, scount As Long
    
    If inDocRoot.childNodes.length > 0 Then
        If inDocRoot.firstChild.nodeTypeString <> "text" Then
        For Each elm In inDocRoot.childNodes
            'Ελεγχος αν υπάρχουν περισσότερα απο ενα buffer parts με το ιδιο ονομα
            scount = 1:
            Set selm = elm.previousSibling
            Do While Not (selm Is Nothing)
                If selm.baseName = elm.baseName Then
                    scount = scount + 1
                Else
                    Exit Do
                End If
                Set selm = selm.previousSibling
            Loop
            UpdateCRAView inViewRoot.ByName(elm.baseName, inscount), elm, scount
        Next elm
        Else
        If inDocRoot.Text <> "" Then inViewRoot.Value = inDocRoot.Text
        End If
    Else
        If inDocRoot.Text <> "" Then inViewRoot.Value = inDocRoot.Text
    End If
End Sub

Public Function ParseCRACall(callElm As IXMLDOMElement) As String
Dim ruleName As String, ruleTrn As String, ruleAuth
Dim ruleApplTran As String
Dim InputViewSystem As String, InputViewModel As String, InputViewName As String
Dim OutputViewSystem As String, OutputViewModel As String, OutputViewName As String, ResultCodeName As String, ErrorViewName As String
Dim attr As IXMLDOMAttribute, elm As IXMLDOMElement, inputViewElm As IXMLDOMElement
Dim UpdateTrnNumOption As String, WriteJournalFlag As String
Dim CheckIRISOutputMethod As String
Dim CheckCRAOutputMethod As String

    CheckIRISOutputMethod = "ChkIRISOutput_"
    CheckCRAOutputMethod = "ChkHPSComResult_"
    
    UpdateTrnNumOption = "before"
    WriteJournalFlag = "yes"
    Set elm = content.selectSingleNode(".")
    If Not elm Is Nothing Then
        Set attr = elm.getAttributeNode("trncounter")
        If Not (attr Is Nothing) Then UpdateTrnNumOption = attr.Text
        Set attr = elm.getAttributeNode("writejournal")
        If Not (attr Is Nothing) Then WriteJournalFlag = attr.Text
    End If
    
    'Στοιχεία από τη δήλωση του rule
    Set elm = content.selectSingleNode(".//rule")
    If Not elm Is Nothing Then
        Set attr = elm.getAttributeNode("name")
        If Not (attr Is Nothing) Then ruleName = attr.Text
        Set attr = elm.getAttributeNode("trncall")
        If Not (attr Is Nothing) Then ruleTrn = attr.Text
        Set attr = elm.getAttributeNode("authuser")
        If Not (attr Is Nothing) Then ruleAuth = attr.Text
        Set attr = elm.getAttributeNode("appltran")
        If Not (attr Is Nothing) Then ruleApplTran = attr.Text
    End If
    Set elm = content.selectSingleNode(".//inputview")
    If Not elm Is Nothing Then
        Set attr = elm.getAttributeNode("system")
        If Not (attr Is Nothing) Then InputViewSystem = attr.Text
        Set attr = elm.getAttributeNode("model")
        If Not (attr Is Nothing) Then InputViewModel = attr.Text
        Set attr = elm.getAttributeNode("name")
        If Not (attr Is Nothing) Then InputViewName = attr.Text
        Set inputViewElm = elm
    End If
    Set elm = content.selectSingleNode(".//outputview")
    If Not elm Is Nothing Then
        ResultCodeName = "": ErrorViewName = ""
        Set attr = elm.getAttributeNode("system")
        If Not (attr Is Nothing) Then OutputViewSystem = attr.Text
        Set attr = elm.getAttributeNode("model")
        If Not (attr Is Nothing) Then OutputViewModel = attr.Text
        Set attr = elm.getAttributeNode("name")
        If Not (attr Is Nothing) Then OutputViewName = attr.Text
        Set attr = elm.getAttributeNode("resultcode")
        If Not (attr Is Nothing) Then ResultCodeName = attr.Text
        Set attr = elm.getAttributeNode("errorview")
        If Not (attr Is Nothing) Then ErrorViewName = attr.Text
    End If
    CheckIRISOutputMethod = "ChkIRISOutput_"
    CheckCRAOutputMethod = "ChkHPSComResult_"
    Set elm = content.selectSingleNode(".//chkoutput")
    If Not elm Is Nothing Then
       Set attr = elm.getAttributeNode("method")
       If (Not attr Is Nothing) Then
        CheckIRISOutputMethod = attr.Text: CheckCRAOutputMethod = attr.Text
       End If
    End If
    For Each elm In callElm.childNodes
        If UCase(elm.baseName) = "RULE" Then
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then ruleName = attr.Text
            Set attr = elm.getAttributeNode("trncall")
            If Not (attr Is Nothing) Then ruleTrn = attr.Text
            Set attr = elm.getAttributeNode("authuser")
            If Not (attr Is Nothing) Then ruleAuth = attr.Text
        ElseIf UCase(elm.baseName) = "INPUTVIEW" Then
            Set attr = elm.getAttributeNode("system")
            If Not (attr Is Nothing) Then InputViewSystem = attr.Text
            Set attr = elm.getAttributeNode("model")
            If Not (attr Is Nothing) Then InputViewModel = attr.Text
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then InputViewName = attr.Text
            Set inputViewElm = elm
        ElseIf UCase(elm.baseName) = "OUTPUTVIEW" Then
            ResultCodeName = "": ErrorViewName = ""
            Set attr = elm.getAttributeNode("system")
            If Not (attr Is Nothing) Then OutputViewSystem = attr.Text
            Set attr = elm.getAttributeNode("model")
            If Not (attr Is Nothing) Then OutputViewModel = attr.Text
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then OutputViewName = attr.Text
            Set attr = elm.getAttributeNode("resultcode")
            If Not (attr Is Nothing) Then ResultCodeName = attr.Text
            Set attr = elm.getAttributeNode("errorview")
            If Not (attr Is Nothing) Then ErrorViewName = attr.Text
        End If
    Next elm
    
    If UpdateTrnNumOption = "before" Then UpdateTrnNum_
    
    Dim BuildInputViewFlag As Boolean, BuildOutputViewFlag As Boolean, FindInputViewFlag As Boolean, FindOutputViewFlag As Boolean
    If InputViewModel <> "" And InputViewName <> "" Then
        If owner.TrnBuffers.Exists(InputViewName) Then
            FindInputViewFlag = True
            owner.TrnBuffers.ByName(InputViewName).ClearData
        Else
            If UCase(InputViewSystem) = "CRA" Then
                BuildInputViewFlag = BuildCRAStruct(owner.TrnBuffers, InputViewModel, InputViewName, True)
            ElseIf UCase(OutputViewSystem) = "IRIS" Then
                BuildInputViewFlag = BuildIRISStruct(owner.TrnBuffers, InputViewModel, InputViewName, True)
            End If
        End If
    End If
    
    If OutputViewModel <> "" And OutputViewName <> "" Then
        If owner.TrnBuffers.Exists(OutputViewName) Then
            FindOutputViewFlag = True
            owner.TrnBuffers.ByName(OutputViewName).ClearData
        Else
            If UCase(OutputViewSystem) = "CRA" Then
                BuildOutputViewFlag = BuildCRAStruct(owner.TrnBuffers, OutputViewModel, OutputViewName, True)
            ElseIf UCase(OutputViewSystem) = "IRIS" Then
                BuildOutputViewFlag = BuildIRISStruct(owner.TrnBuffers, OutputViewModel, OutputViewName, True)
            End If
        End If
    End If
    
    Dim inputDoc As MSXML2.DOMDocument30
    Set inputDoc = owner.TrnBuffers.ByName(InputViewName).GetXMLView
    If Not (inputDoc.selectSingleNode(".//VCUUP01") Is Nothing) Then
        owner.TrnBuffers.ByName(InputViewName).v2Data("VCUUP01") = GenWorkForm.AppBuffers.ByName("VCUUP01").Data
    End If
    If Not (inputDoc.selectSingleNode(".//CUF_USR_PRFL_D") Is Nothing) Then
        owner.TrnBuffers.ByName(InputViewName).v2Data("CUF_USR_PRFL_D") = GenWorkForm.AppBuffers.ByName("VCUUP01").Data
    End If
    If Not (inputDoc.selectSingleNode(".//CUF_USR_OL_PRFL_D") Is Nothing) Then
        With owner.TrnBuffers.ByName(InputViewName).ByName("CUF_USR_OL_PRFL_D")
            If .ByName("I_ENTP", 1).Value = 0 Then .ByName("I_ENTP", 1).Value = GenWorkForm.AppBuffers.ByName("VCUUP01").ByName("I_ENTP").Value
            If Trim(.ByName("C_ACOD_OU", 1).Value) = "" Then .ByName("C_ACOD_OU", 1).Value = GenWorkForm.AppBuffers.ByName("VCUUP01").ByName("C_ACOD_OU").Value
            If Trim(.ByName("C_ACOD_FI", 1).Value) = "" Then .ByName("C_ACOD_FI", 1).Value = GenWorkForm.AppBuffers.ByName("VCUUP01").ByName("C_ACOD_FI").Value
            If Trim(.ByName("C_USR_ID", 1).Value) = "" Then .ByName("C_USR_ID", 1).Value = GenWorkForm.AppBuffers.ByName("VCUUP01").ByName("C_USR_ID").Value
            If Trim(.ByName("C_WKST_ID", 1).Value) = "" Then .ByName("C_WKST_ID", 1).Value = GenWorkForm.AppBuffers.ByName("VCUUP01").ByName("C_WKST_ID").Value
        End With
    End If
    If Not (inputDoc.selectSingleNode(".//STD_APPL_PARM_V") Is Nothing) Then
        Dim connectsend, connectreceive
        Set connectsend = GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_I")
        Set connectreceive = GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_O")
        With owner.TrnBuffers.ByName(InputViewName).ByName("STD_APPL_PARM_V")
            .ByName("COD_NRBE_EN", 1).Value = connectreceive.v2Value("COD_NRBE_EN")
            .ByName("COD_INTERNO_UO", 1).Value = connectreceive.v2Value("COD_INTERNO_UO")
            .ByName("COD_CSB_OF", 1).Value = connectreceive.v2Value("COD_CSB_OF")
            .ByName("ID_INTERNO_TERM_TN", 1).Value = connectsend.v2Value("ID_INTERNO_TERM_TN")
            .ByName("ID_INTERNO_EMPL_EP", 1).Value = connectreceive.v2Value("ID_INTERNO_EMPL_EP")
            .ByName("NUM_SEC", 1).Value = cTRNNum
            .ByName("FECHA_OPRCN", 1).Value = Date
            .ByName("HORA_OPRCN", 1).Value = Time
            .ByName("FECHA_CTBLE", 1).Value = connectreceive.v2Value("FECHA_CTBLE")
        End With
    End If

    UpdateCRAView owner.TrnBuffers.ByName(InputViewName), inputViewElm, 1
    
    Dim res As Integer
    SenseCode = "": SenseCodeMessage = ""
    res = IRISCom_(GenWorkForm, ruleTrn, ruleName, owner.TrnBuffers.ByName(InputViewName), owner.TrnBuffers.ByName(OutputViewName), ruleAuth, ruleApplTran, , , False, (WriteJournalFlag = "yes"))
    Dim ErrorView
    
    If res = 0 Then
        Dim HPSComResult As Integer
        
        If ErrorViewName <> "" And ResultCodeName <> "" Then
            Dim ResultCode
            ResultCode = owner.TrnBuffers.ByName(OutputViewName).v2Value(ResultCodeName)
            Set ErrorView = owner.TrnBuffers.ByName(OutputViewName).ByName(ErrorViewName)
            
            If ErrorView.name = "NBG_STD_ERR_VIEW" Then
               Dim res1 As Boolean
               res1 = ChkCRA2Output_(owner.TrnBuffers.ByName(OutputViewName))
            Else
                If UCase(OutputViewSystem) = "CRA" Then
                    If CheckCRAOutputMethod = "ChkHPSComResult_" Then
                        HPSComResult = ChkHPSComResult_(ResultCode, ErrorView)
                    Else
                        HPSComResult = ChkCRA2Output_(owner.TrnBuffers.ByName(OutputViewName))
                    End If
                Else
                    If CheckIRISOutputMethod = "ChkIRISOutput_" Then
                        HPSComResult = ChkIRISOutput_(owner.TrnBuffers.ByName(OutputViewName), False)
                    Else
                        HPSComResult = ChkIRISOutputSkip7_(owner.TrnBuffers.ByName(OutputViewName), False)
                    End If
                End If
            End If
            If UpdateTrnNumOption = "after" Then UpdateTrnNum_
        End If
        ParseCRACall = owner.TrnBuffers.ByName(OutputViewName).GetXMLView.XML
    Else
        If ErrorViewName <> "" Then
            Set ErrorView = owner.TrnBuffers.ByName(OutputViewName).ByName(ErrorViewName)
            If ErrorView.name = "NBG_STD_ERR_VIEW" Then
                If SenseCode <> "" Or SenseCodeMessage <> "" Then
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT") = 0
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT_ERRNO") = SenseCode
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT_TEXT") = SenseCodeMessage
                Else
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT") = 0
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT_ERRNO") = 999
                    owner.TrnBuffers.ByName(OutputViewName).v2Value("C_RSLT_TEXT") = "ΠΡΟΒΛΗΜΑ ΕΠΙΚΟΙΝΩΝΙΑΣ-ΑΠΕΤΥΧΕ Η ΣΥΝΑΛΛΑΓΗ"
                End If
                ParseCRACall = owner.TrnBuffers.ByName(OutputViewName).GetXMLView.XML
            ElseIf ErrorView.name = "STD_TRN_MSJ_PARM_V" Then
                ParseCRACall = owner.TrnBuffers.ByName(OutputViewName).GetXMLView.XML
            Else
                ParseCRACall = ""
            End If
        End If
    End If
End Function

Public Function LoadXML(inDocument As String) As String

    Dim workDocument As New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    
    Set attr = workDocument.documentElement.getAttributeNode("system")
    If attr Is Nothing Then
        If Not content Is Nothing Then
            Set attr = content.getAttributeNode("system")
        End If
    End If
    If Not (attr Is Nothing) Then System = attr.Text
    If UCase(System) = "CRA" Or UCase(System) = "IRIS" Then
    
        Dim res As String
        res = ParseCRACall(workDocument.documentElement)
        LoadXML = RemoveInvalidXMLChars(res)
        
'        LoadXML = ParseCRACall(workDocument.documentElement)
'        LoadXML = Replace(LoadXML, Chr(4), " ")
'        LoadXML = Replace(LoadXML, Chr(28), " ")
    End If

End Function

Function RemoveInvalidXMLChars(inputStr As String) As String
    Dim doc As New MSXML2.DOMDocument60
    Dim OutputStr As String
    Dim res As Boolean
    res = doc.LoadXML(inputStr)
    If res = False And doc.parseError <> 0 And doc.parseError.filepos <> 0 Then
        OutputStr = Replace(inputStr, Mid(inputStr, doc.parseError.filepos + 1, 1), " ")
        inputStr = RemoveInvalidXMLChars(OutputStr)
    End If
    
    RemoveInvalidXMLChars = inputStr
End Function


Public Function XML() As String
    If content Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(content.XML)
    End If
End Function

