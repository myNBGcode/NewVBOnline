VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLLocalMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement
Public scriptHost

Private Function BuildHost(filename As String) As Boolean
    BuildHost = False
    On Error GoTo ReportCreateError
    If IsEmpty(scriptHost) Then Set scriptHost = CreateObject("MSScriptControl.ScriptControl.1")
    On Error Resume Next
    Close #1
    On Error GoTo ReportFileOpenError
    Open ReadDir & "\" & filename For Input As #1
    On Error GoTo ReportFileReadError
    Dim code As String, s As String
    Do While Not Eof(1)
        Line Input #1, s
        code = code & vbCrLf & s
    Loop
    Close #1
    On Error GoTo ReportScriptCodeError
    scriptHost.language = "vbscript"
    scriptHost.AddCode code
    BuildHost = True: Exit Function
ReportCreateError:
    LogMsgbox "Λάθος XMLLocalMethod.BuildHost.Create: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην ανάκτηση πρόσθετων προγραμμάτων."
    Exit Function
ReportFileOpenError:
    LogMsgbox "Λάθος XMLLocalMethod.BuildHost.FileOpeFn: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην ανάκτηση πρόσθετων προγραμμάτων."
    Exit Function
ReportFileReadError:
    LogMsgbox "Λάθος XMLLocalMethod.BuildHost.FileRead: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην ανάκτηση πρόσθετων προγραμμάτων."
    Exit Function
ReportScriptCodeError:
    LogMsgbox "Λάθος XMLLocalMethod.BuildHost.FileScriptCode: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην ανάκτηση πρόσθετων προγραμμάτων."
    Exit Function
End Function

Private Function MethodCall(methodnode As MSXML2.IXMLDOMElement, inputnode As MSXML2.IXMLDOMElement) As String
Dim paramList As MSXML2.IXMLDOMNodeList
Dim argList(), i As Integer
Dim ParamNode As MSXML2.IXMLDOMElement, attr As MSXML2.IXMLDOMAttribute
        MethodCall = ""
        On Error GoTo paramListBuildError
        Set paramList = methodnode.SelectNodes("./input")
        If paramList.length > 0 Then
            ReDim argList(paramList.length - 1): i = 0
            For Each ParamNode In paramList
                Set attr = ParamNode.Attributes.getNamedItem("name")
                If Not (attr Is Nothing) Then
                    If Not (inputnode.selectSingleNode(".//" & attr.value) Is Nothing) Then
                        argList(i) = inputnode.selectSingleNode(".//" & attr.value).Text
                    End If
                Else
                    GoTo ParamNameNotFound
                End If
                i = i + 1
            Next ParamNode
            On Error GoTo ScriptRunError
            Set attr = methodnode.Attributes.getNamedItem("name")
            If Not (attr Is Nothing) Then
                MethodCall = scriptHost.Run(attr.value, argList)
            Else
                GoTo MethodNameNotFound
            End If
        Else
            On Error GoTo ScriptRunError
            Set attr = methodnode.Attributes.getNamedItem("name")
            If Not (attr Is Nothing) Then
                MethodCall = scriptHost.Run(attr.value)
            Else
                GoTo MethodNameNotFound
            End If
        End If
        Exit Function
paramListBuildError:
    LogMsgbox "Λάθος XMLLocalMethod.MethodCall.ParamListBuild: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
ScriptRunError:
    LogMsgbox "Λάθος XMLLocalMethod.MethodCall.ScriptRun: " & Err.number & "-" & Err.description, vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
ParamNameNotFound:
    LogMsgbox "Λάθος XMLLocalMethod.MethodCall.ParamName: Δεν βρέθηκε το όνομα της παραμέτρου", vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
MethodNameNotFound:
    LogMsgbox "Λάθος XMLLocalMethod.MethodCall.MethodName: Δεν βρέθηκε το όνομα της διαδικασίας", vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
End Function


'<method>
'   <declare name="xxxx"/>
'   <input>
'        <XXXXX>yyyyyy</XXXXX>
'   </input>
'</method>

Private Function parseMethod(Node As IXMLDOMNode) As String
    Dim elm As IXMLDOMNode, attr As IXMLDOMAttribute
    Dim TrnCode As String, inDoc As MSXML2.DOMDocument30, outDoc As MSXML2.DOMDocument30
    Dim trnHandler As L2TrnHandler
    
    Dim methodname As String, inputnode As IXMLDOMElement
    
    For Each elm In Node.childNodes
        If elm.nodeTypeString = "element" Then
            If UCase(elm.baseName) = UCase("declare") Then
                Set attr = elm.Attributes.getNamedItem("name")
                If Not (attr Is Nothing) Then methodname = attr.Text
            ElseIf UCase(elm.baseName) = UCase("input") Then
                Set inputnode = elm
            End If
        End If
    Next elm
                  
    If UCase(methodname) = UCase("L2ChkBankAccount") Then
        parseMethod = L2ChkBankAccount(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkBankCheque") Then
        parseMethod = L2ChkBankCheque(inputnode)
    ElseIf UCase(methodname) = UCase("L2CDETECheque") Then
        parseMethod = L2CDETECheque(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkBank") Then
        parseMethod = L2ChkBank(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkEteAccount") Then
        parseMethod = L2ChkETEAccount(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkEteCheque") Then
        parseMethod = L2ChkETECheque(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChiefKey") Then
        parseMethod = L2ChiefKey(inputnode)
    ElseIf UCase(methodname) = UCase("L2ManagerKey") Then
        parseMethod = L2ManagerKey(inputnode)
    ElseIf UCase(methodname) = UCase("L2AnyKey") Then
        parseMethod = L2AnyKey(inputnode)
    ElseIf UCase(methodname) = UCase("L2AnyKeyUnconditional") Then
        parseMethod = L2AnyKeyUnconditional(inputnode)
    ElseIf UCase(methodname) = UCase("L2IRISAuth") Then
        parseMethod = L2IRISAuth(inputnode)
    ElseIf UCase(methodname) = UCase("L2IRISAuthLevel") Then
        parseMethod = L2IRISAuthLevel(inputnode)
    ElseIf UCase(methodname) = UCase("L2FTFilaName") Then
        parseMethod = L2FTFilaName(inputnode)
    ElseIf UCase(methodname) = UCase("L2FTFilaRecordset") Then
        parseMethod = L2FTFilaRecordset(inputnode)
    ElseIf UCase(methodname) = UCase("L2GetSingleFTFilaName") Then
        parseMethod = L2GetSingleFTFilaName(inputnode)
    ElseIf UCase(methodname) = UCase("L2Recordset") Then
        parseMethod = L2Recordset(inputnode)
    ElseIf UCase(methodname) = UCase("L2GetAmountText2002") Then
        parseMethod = L2GetAmountText2002(inputnode)
    ElseIf UCase(methodname) = UCase("L2GetAmountText") Then
        parseMethod = L2GetAmountText(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkTaxID") Then
        parseMethod = L2ChkTaxID(inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookLine") Then
        parseMethod = L2PrintPassbookLine(inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbook") Then
        parseMethod = L2PrintPassbook(inputnode)
'    ElseIf UCase(methodname) = UCase("L2PrintPassbookNew") Then
'        parseMethod = L2PrintPassbookNew(Me, inputnode)
'    ElseIf UCase(methodname) = UCase("L2PrintPassbookNew_WithRowsMore") Then
'        parseMethod = L2PrintPassbookNew_WithRowsMore(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookNewVersion3") Then
        parseMethod = L2PrintPassbookNewVersion3(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookNewVersion4") Then
        parseMethod = L2PrintPassbookNewVersion4(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookNewVersion5") Then
        parseMethod = L2PrintPassbookNewVersion5(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookNewVersion6") Then
        parseMethod = L2PrintPassbookNewVersion6(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintPassbookNewVersion7") Then
        parseMethod = L2PrintPassbookNewVersion7(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintExchangePassbook") Then
        parseMethod = L2PrintExchangePassbook(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintExchangePassbookNew") Then
        parseMethod = L2PrintExchangePassbookNew(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2PrintExchangePassbookNew1") Then
        parseMethod = L2PrintExchangePassbookNew1(Me, inputnode)
    ElseIf UCase(methodname) = UCase("L2Show1041Messages") Then
        parseMethod = L2Show1041Messages(inputnode)
    ElseIf UCase(methodname) = UCase("L2SCreateIBAN") Then
        parseMethod = L2SCreateIBAN(inputnode)
    ElseIf UCase(methodname) = UCase("L2FormatIBAN") Then
        parseMethod = L2FormatIBAN(inputnode)
    ElseIf UCase(methodname) = UCase("L2GetChequeAmountText") Then
        parseMethod = L2GetChequeAmountText(inputnode)
    ElseIf UCase(methodname) = UCase("L2ReadOCR") Then
        parseMethod = L2ReadOCR()
    ElseIf UCase(methodname) = UCase("L2ChkETEBankCheque") Then
        parseMethod = L2ChkETEBankCheque(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkValidIBAN") Then
        parseMethod = L2ChkValidIBAN(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkCard") Then
        parseMethod = L2ChkCard(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkSATNo") Then
        parseMethod = L2ChkSATNo(inputnode)
    ElseIf UCase(methodname) = UCase("L2AreaBuilder") Then
        parseMethod = L2AreaBuilder(inputnode)
    ElseIf UCase(methodname) = UCase("L2SendXmlBuffer") Then
        parseMethod = L2SendXmlBuffer(inputnode)
    ElseIf UCase(methodname) = UCase("L2MessageBox") Then
        parseMethod = L2MessageBox(inputnode)
    ElseIf UCase(methodname) = UCase("L2CalcCheckDigits") Then
        parseMethod = L2CalcCheckDigits(inputnode)
    ElseIf UCase(methodname) = UCase("L2CalculateTerminalID") Then
        parseMethod = L2CalculateTerminalID(inputnode)
    ElseIf UCase(methodname) = UCase("L2SuspectedTrnHandler") Then
        parseMethod = L2SuspectedTrnHandler(inputnode)
    ElseIf UCase(methodname) = UCase("L2BuildSwiftMessage") Then
        parseMethod = L2BuildSwiftMessage(inputnode)
    ElseIf UCase(methodname) = UCase("L2CalculateIRISTime") Then
        parseMethod = L2CalculateIRISTime(inputnode)
    ElseIf UCase(methodname) = UCase("L2CalcCd1") Then
        parseMethod = L2CalcCd1(inputnode)
    ElseIf UCase(methodname) = UCase("L2ShowIRISMessages") Then
        parseMethod = L2ShowIRISMessages(inputnode)
    ElseIf UCase(methodname) = UCase("L2DateAdd") Then
        parseMethod = L2DateAdd(inputnode)
    ElseIf UCase(methodname) = UCase("L2DateDiff") Then
        parseMethod = L2DateDiff(inputnode)
    ElseIf UCase(methodname) = UCase("L2ValidateCheckInput") Then
        parseMethod = L2ValidateCheckInput(inputnode)
    ElseIf UCase(methodname) = UCase("L2CalcCd2") Then
        parseMethod = L2CalcCd2(inputnode)
    ElseIf UCase(methodname) = UCase("L2ChkDocumentNo") Then
        parseMethod = L2ChkDocumentNo(inputnode)
    ElseIf UCase(methodname) = UCase("L2GetIRISErrorData") Then
        parseMethod = L2GetIRISErrorData(inputnode)
    Else
        If L2AddInFile.documentElement Is Nothing Then
            parseMethod = ""
        Else
            Dim methodnode As IXMLDOMElement, paramList As MSXML2.IXMLDOMNodeList
            Dim i As Integer, Result As String, argList()
                        
            Set methodnode = L2AddInFile.selectSingleNode("//package/function[@name = '" & methodname & "']")
            If Not (methodnode Is Nothing) Then
                If Not (methodnode.parentNode Is Nothing) Then
                    Set attr = methodnode.parentNode.Attributes.getNamedItem("file")
                    If Not (attr Is Nothing) Then
                        If Not BuildHost(attr.value) Then Exit Function
                        Result = MethodCall(methodnode, inputnode)
                        parseMethod = "<MESSAGE><DATA>" & Result & "</DATA></MESSAGE>"
                        Exit Function
                    Else
                        GoTo MethodFileNotFound
                    End If
                Else
                    GoTo MethodPackageNotFound
                End If
            Else
                GoTo MethodNameNotFound
            End If
        End If
        parseMethod = ""
    End If
    
    If Not (trnHandler Is Nothing) Then
        trnHandler.CleanUp
        Set trnHandler = Nothing
    End If
    Exit Function
    
MethodFileNotFound:
    LogMsgbox "Λάθος XMLLocalMethod.Parse.PackageFile: Δεν βρέθηκε το αρχείο της ενότητας", vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
MethodPackageNotFound:
    LogMsgbox "Λάθος XMLLocalMethod.Parse.MethodPackage: Δεν βρέθηκε το όνομα της ενότητας", vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
MethodNameNotFound:
    LogMsgbox "Λάθος XMLLocalMethod.Parse.MethodName: Δεν βρέθηκε το όνομα της διαδικασίας", vbCritical, "Λάθος στην κλήση πρόσθετων προγραμμάτων."
    Exit Function
End Function


Public Function LoadXML(inDocument As String) As String
    Dim workDocument As New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    
    Set elm = workDocument.documentElement
    If UCase(elm.baseName) = "METHOD" Then
        Dim newdocument As MSXML2.DOMDocument30
        Set newdocument = New MSXML2.DOMDocument30
        If content Is Nothing Then
            newdocument.appendChild newdocument.createElement("method")
        Else
            ImportElement content, newdocument
        End If
        
        Dim localnode As IXMLDOMNode
        For Each localnode In elm.childNodes
            If localnode.nodeType = NODE_ELEMENT Then
                Dim localelm As IXMLDOMElement
                Set localelm = localnode
                If Not newdocument.documentElement.selectSingleNode("./" & localelm.baseName) Is Nothing Then
                    newdocument.documentElement.removeChild newdocument.documentElement.selectSingleNode("./" & localelm.baseName)
                End If
                If newdocument.documentElement.selectSingleNode("./" & localelm.baseName) Is Nothing Then
                    ImportElement localelm, newdocument.documentElement
                End If
            End If
        Next localnode
        
        'Set content = elm
        
        LoadXML = parseMethod(newdocument.documentElement)
        
        Dim resultdocument As New MSXML2.DOMDocument30
        resultdocument.LoadXML LoadXML
        If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            Load XMLMessageForm
            Set XMLMessageForm.MessageDocument = resultdocument
            XMLMessageForm.Show vbModal
            Exit Function
        End If
    End If
End Function

Public Function XML() As String
    If content Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(content.XML)
    End If
End Function


