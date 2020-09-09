VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLDocumentManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public owner As L2TrnHandler
'Public OwnerForm As Form

Public XmlObjectList As New cCollection
Public nameList As New Collection
Public typeList As New Collection

Public TrnBuffers As Buffers

Private TemplateDocument As MSXML2.DOMDocument30
Private TempJobDocument As MSXML2.DOMDocument30

Private currentType As String

Private Sub Class_Initialize()
    Set TrnBuffers = New Buffers
End Sub

Public Property Get activeform()
    If owner Is Nothing Then
        Set activeform = Nothing
    Else
        Set activeform = owner.activeform
    End If
End Property

Public Sub ShowMessage(amessage As String)
    If activeform Is Nothing Then
    Else
        activeform.sbWriteStatusMessage amessage
    End If
End Sub

Public Function getTempJob(code As String) As cXMLDocumentJob
    Set getTempJob = New cXMLDocumentJob
    Set getTempJob.Manager = Me
    Dim codedoc As New MSXML2.DOMDocument30
    codedoc.LoadXML "<job name=""tmpjob"">" & code & "</job>"
    Set getTempJob.content = codedoc.documentElement
    getTempJob.Title = "tmpjob"
    
End Function

Property Get objectType(ByVal Index As Long) As String
    objectType = typeList(Index)
End Property

Property Get namedobjectType(ByVal name As String) As String
    On Error GoTo ErrorPos
    namedobjectType = typeList(name)
    Exit Property
ErrorPos:
    namedobjectType = ""
End Property

Property Get objectName(ByVal Index As Long) As String
    objectName = nameList(Index)
End Property

Property Get xmlObject(ByVal Index As Long)
    Set xmlObject = XmlObjectList.Item(CStr(Index))
End Property

Property Get namedxmlObject(ByVal name As String)
    On Error GoTo ErrorPos
    Set namedxmlObject = XmlObjectList.Item(name)
    Exit Property
ErrorPos:
    Set namedxmlObject = Nothing
End Property

Public Function Exec(ByVal name As String) As String
    Exec = ""
    Dim aobj
    Set aobj = FindXmlObject(name)
    If (aobj Is Nothing) Then
        LogMsgbox "Δεν βρέθηκε ο κώδικας: " & name, vbCritical, "Λάθος"
        Exit Function
    Else
        Exec = aobj.XML
    End If
End Function

Public Function ExecCommand(ByVal command As String) As String
    If TempJobDocument Is Nothing Then Set TempJobDocument = New MSXML2.DOMDocument30
    
    ExecCommand = ""
    Dim aobj As cXMLDocumentJob
    
    Set aobj = getTempJob(command)
    If (aobj Is Nothing) Then
        LogMsgbox "Δεν δημιουργήθηκε το προσωρινό job ", vbCritical, "Λάθος"
        Exit Function
    Else
        ExecCommand = aobj.LoadXML(TempJobDocument.XML)
        TempJobDocument.LoadXML ExecCommand
    End If
    
End Function

Property Get xmlObjectContent(ByVal Index As Long) As String
    On Error GoTo ErrorPos
    Select Case typeList(Index)
        Case "job"
            xmlObjectContent = XmlObjectList.Item(CStr(Index)).content.XML
        Case "transformation"
            xmlObjectContent = XmlObjectList.Item(CStr(Index)).content.XML
        Case "document", "form", "formupdate", "codeformupdate", "report", "genworkform", "l2form"
            xmlObjectContent = XmlObjectList.Item(CStr(Index)).XML
        Case Else
            xmlObjectContent = ""
    End Select
    Exit Property
ErrorPos:
    LogMsgbox "Αδύνατη η ανάκτηση περιεχομένου από το αντικείμενο " & Index & " " & Err.number & " " & Err.description, vbCritical, "Λάθος..."
End Property

Property Get namedxmlObjectContent(ByVal name As String) As String
    On Error GoTo ErrorPos
    Select Case typeList(name)
        Case "report", "job", "transformation", "weblink", "webdocumentmethod"
            If XmlObjectList.Item(name).content Is Nothing Then
                namedxmlObjectContent = ""
            Else
                namedxmlObjectContent = XmlObjectList.Item(name).content.XML
            End If
        Case "document", "form", "formupdate", "codeformupdate", "genworkform", "l2form"
            namedxmlObjectContent = XmlObjectList.Item(name).XML
        Case Else
            namedxmlObjectContent = ""
    End Select
    Exit Property
ErrorPos:
    LogMsgbox "Αδύνατη η ανάκτηση περιεχομένου από το αντικείμενο " & name & " " & Err.number & " " & Err.description, vbCritical, "Λάθος..."
    namedxmlObjectContent = ""
End Property

Private Function FindXmlObject(name As String)
    On Error GoTo ErrorPos
    Set FindXmlObject = XmlObjectList.Item(name)
    Exit Function
ErrorPos:
    Set FindXmlObject = Nothing
End Function

Public Function Upload(ByVal value As String) As String
    Upload = ""
    Dim adoc As New MSXML2.DOMDocument30, aattr As IXMLDOMAttribute
    adoc.LoadXML value
    If adoc.documentElement Is Nothing Then
        LogMsgbox "Απέτυχε η διαδικασία ανάκτησης για τον κώδικα: " & value, vbCritical, "Λάθος"
        Exit Function
    End If
    Set aattr = adoc.documentElement.Attributes.getNamedItem("name")
    If aattr Is Nothing Then
        LogMsgbox "Δεν ορίστικε στοιχείο ονόματος για τον κώδικα: " & value, vbCritical, "Λάθος"
        Exit Function
    End If
    Dim aobj, aname As String
    
    aname = aattr.Text
    Set aobj = FindXmlObject(aname)
    If Not (aobj Is Nothing) Then
        XmlObjectList.Remove aname
        nameList.Remove aname
        typeList.Remove aname
    End If
    Set aobj = ParseElement(adoc.documentElement)
    If aobj Is Nothing Then
        LogMsgbox "Απέτυχε η επεξεργασία για τον κώδικα: " & value, vbCritical, "Λάθος"
        Exit Function
    Else
        XmlObjectList.add aobj, aname
        nameList.add aattr.Text, aname
        typeList.add currentType, aname
    End If
    Upload = aname
End Function

Public Function ParseElement(elm As IXMLDOMElement)
Dim innerDocument As cXmlDocument
Dim TrnLink As cTRNLink
Dim L2TrnLink As cL2TrnLink
Dim HostMethod As cHostMethod
Dim HostMethodAuth As cHostMethodAuth
Dim XmlHostMethodBuilder As cXmlHostMethodBuilder
Dim transformation As CXMLDocumentTransformation
Dim job As cXMLDocumentJob
Dim journal As cXMLJournal
Dim Validation As cXMLValidations
Dim aWeblink As cXMLWebLink, VirtualDirectory As String, aWebMethod As cXMLWebMethod
Dim LocalMethod As cXMLLocalMethod
Dim Report As cXMLReport, ReportL2ListBox As cXMLReportL2Listbox
Dim ComArea As cXmlComArea
Dim KAAMessage As cXmlKAAMessage
Dim totals As cL2Totals
Dim aname As String, anamespace As String
Dim nameattr As IXMLDOMAttribute, name As String
Dim messagerversion3 As cXmlDepositMessageHandlerVersion3
Dim messagerversion4 As cXmlDepositMessageHandlerVersion4
Dim exchangemessager As cXmlExchangeDepositMessageHandler
Dim CurrencyPassbook As cXMLCurrencyPassbook
Dim massmessager As cXMLDepositMassiveMessageHandler
'Dim ESignature As cXmlESignatureReport
Dim curMessager As cXmlCurrencyMessageHandler
Dim aWebApi As cXmlWebAPI, WebAPIVirtualDirectory As String, WrapperWebAPIVirtualDirectory As String
    
    Set ParseElement = Nothing
    Set nameattr = elm.Attributes.getNamedItem("name")
    If nameattr Is Nothing Then name = "" Else name = nameattr.Text
    If UCase(elm.baseName) = "FORM" Then
        On Error GoTo FormLoadError
        Set ParseElement = elm
        currentType = "form"
    ElseIf UCase(elm.baseName) = "FORMUPDATE" Then
        On Error GoTo FormUpdateLoadError
        Set ParseElement = elm
        currentType = "formupdate"
    ElseIf UCase(elm.baseName) = "DOCUMENT" Then
        On Error GoTo DocumentLoadError
        Set innerDocument = New cXmlDocument
        Set innerDocument.content = elm
        Set ParseElement = innerDocument
        currentType = "document"
    ElseIf UCase(elm.baseName) = "REPORT" Then
        On Error GoTo ReportLoadError
        Set Report = New cXMLReport
        Set Report.content = elm
        Set ParseElement = Report
        currentType = "report"
    ElseIf UCase(elm.baseName) = "L2LISTBOXREPORT" Then
        On Error GoTo ReportLoadError
        Set ReportL2ListBox = New cXMLReportL2Listbox
        Set ReportL2ListBox.content = elm
        Set ReportL2ListBox.ESignature = elm.Attributes.getNamedItem("esignature")
        Set ParseElement = ReportL2ListBox
        currentType = "l2listboxreport"
   'ElseIf UCase(elm.baseName) = "ESIGNATUREREPORT" Then
   '     On Error GoTo EsignLoadError
   '     Set ESignature = New cXmlESignatureReport
   '     Set ESignature.content = elm
   '     Set ESignature.FallbackReport = elm.Attributes.getNamedItem("fallback")
   '     Set ParseElement = ESignature
   '     currentType = "esignature"
    ElseIf UCase(elm.baseName) = "CODEFORMUPDATE" Then
        On Error GoTo CodeFormUpdateLoadError
        Set ParseElement = elm
        currentType = "codeformupdate"
    ElseIf UCase(elm.baseName) = "TRNLINK" Then
        On Error GoTo TrnLinkLoadError
        Set TrnLink = New cTRNLink
        Set TrnLink.content = elm
        Set ParseElement = TrnLink
        currentType = "trnlink"
    ElseIf UCase(elm.baseName) = "LINKL2TRN" Then
        On Error GoTo LinkL2TrnLoadError
        Set L2TrnLink = New cL2TrnLink
        Set L2TrnLink.content = elm
        Set ParseElement = L2TrnLink
        currentType = "l2trnlink"
    ElseIf UCase(elm.baseName) = "CALL" Then
        On Error GoTo CallLoadError
        Set HostMethod = New cHostMethod
        Set HostMethod.content = elm
        Set ParseElement = HostMethod
        Set HostMethod.owner = Me
        currentType = "hostmethod"
    ElseIf UCase(elm.baseName) = "CALL_AUTH" Then
        On Error GoTo CallLoadErrorAuth
        Set HostMethodAuth = New cHostMethodAuth
        Set HostMethodAuth.content = elm
        Set ParseElement = HostMethodAuth
        Set HostMethodAuth.owner = Me
        currentType = "hostmethodauth"
    ElseIf UCase(elm.baseName) = "HOSTMETHOD" Then
        On Error GoTo XmlHostMethodLoadError
        Set XmlHostMethodBuilder = New cXmlHostMethodBuilder
        Set XmlHostMethodBuilder.content = elm
        Set ParseElement = XmlHostMethodBuilder
        currentType = "xmlhostmethod"
    ElseIf UCase(elm.baseName) = "COMAREA" Then
        On Error GoTo ComAreaLoadError
        Set ComArea = New cXmlComArea
        Set ComArea.content = elm
        Set ParseElement = ComArea
        Set ComArea.Container = Me.TrnBuffers
        currentType = "xmlcomarea"
    ElseIf UCase(elm.baseName) = "TRANSFORMATION" _
    Or UCase(elm.baseName) = "STYLESHEET" Then
        On Error GoTo TransformationLoadError
        Set transformation = New CXMLDocumentTransformation
        Set transformation.content = elm
        Set transformation.Manager = Me
        Set ParseElement = transformation
        currentType = "transformation"
    ElseIf UCase(elm.baseName) = "METHOD" Then
        On Error GoTo MethodLoadError
        Set LocalMethod = New cXMLLocalMethod
        Set LocalMethod.Manager = Me
        Set LocalMethod.content = elm
        Set ParseElement = LocalMethod
        currentType = "localmethod"
    ElseIf UCase(elm.baseName) = "JOB" Then
        On Error GoTo JobLoadError
        Set job = New cXMLDocumentJob
        Set job.content = elm
        Set job.Manager = Me
        Set ParseElement = job
        currentType = "job"
    ElseIf UCase(elm.baseName) = "JOURNAL" Then
        On Error GoTo JournalLoadError
        Set journal = New cXMLJournal
        Set journal.StaticLink = elm
        Set journal.Manager = Me
        Set ParseElement = journal
        currentType = "journal"
    ElseIf UCase(elm.baseName) = "CURRENCYPASSBOOK" Then
        On Error GoTo CurrencyPassbookError
        Set CurrencyPassbook = New cXMLCurrencyPassbook
        Set CurrencyPassbook.StaticLink = elm
        Set CurrencyPassbook.Manager = Me
        Set ParseElement = CurrencyPassbook
        currentType = "CurrencyPassbook"
    ElseIf UCase(elm.baseName) = "VALIDATION" Then
        On Error GoTo ValidationLoadError
        Set Validation = New cXMLValidations
        Set Validation.content = elm
        Set Validation.Manager = Me
        Set ParseElement = Validation
        currentType = "validation"
    ElseIf UCase(elm.baseName) = "WEBAPI" Then
        On Error GoTo WebApiLoadError
        WebAPIVirtualDirectory = elm.getAttribute("apivirtualdirectory")
        WrapperWebAPIVirtualDirectory = elm.getAttribute("wrappervirtualdirectory")
        Set aWebApi = New cXmlWebAPI
        If Trim(WebAPIVirtualDirectory) <> "" Then aWebApi.ApiUrl = owner.WebLinkWithNoExists(WebAPIVirtualDirectory)
        If Trim(WrapperWebAPIVirtualDirectory) <> "" Then aWebApi.WrapperUrl = owner.WebLinkWithNoExists(WrapperWebAPIVirtualDirectory)
        Set ParseElement = aWebApi
        currentType = "webapi"
    ElseIf UCase(elm.baseName) = "WEBLINK" Then
        On Error GoTo WebLinkLoadError
        VirtualDirectory = elm.getAttribute("virtualdirectory")
        Set aWeblink = New cXMLWebLink
        Set aWeblink.content = elm
        aWeblink.VirtualDirectory = owner.WebLink(VirtualDirectory)
        Set ParseElement = aWeblink
        currentType = "weblink"
    ElseIf UCase(elm.baseName) = "WEBMETHOD" Then
        VirtualDirectory = elm.getAttribute("virtualdirectory")
        Set aWeblink = XmlObjectList.Item(VirtualDirectory)
        If aWeblink Is Nothing Then
            On Error GoTo WebLinkLoadError
            Set aWeblink = New cXMLWebLink
            Set aWeblink.content = elm
            aWeblink.VirtualDirectory = owner.WebLink(VirtualDirectory)
            XmlObjectList.add aWeblink, VirtualDirectory
            nameList.add VirtualDirectory, VirtualDirectory
            typeList.add "weblink", VirtualDirectory
        End If
        On Error GoTo WebMethodLoadError
        Set nameattr = elm.Attributes.getNamedItem("name")
        If Not (nameattr Is Nothing) Then
            Set aWebMethod = XmlObjectList.Item(nameattr.value)
            If aWebMethod Is Nothing Then
                anamespace = elm.getAttribute("namespace")
                Set aWebMethod = aWeblink.DefineMethod(name, anamespace)
                Set aWebMethod.content = elm
                Set ParseElement = aWebMethod
                currentType = "webmethod"
            End If
        End If
    ElseIf UCase(elm.baseName) = "WEBDOCUMENTMETHOD" Then
        VirtualDirectory = elm.getAttribute("virtualdirectory")
        Set aWeblink = XmlObjectList.Item(VirtualDirectory)
        If aWeblink Is Nothing Then
            On Error GoTo WebLinkLoadError
            Set aWeblink = New cXMLWebLink
            Set aWeblink.content = elm
            aWeblink.VirtualDirectory = owner.WebLink(VirtualDirectory)
            XmlObjectList.add aWeblink, VirtualDirectory
            nameList.add VirtualDirectory, VirtualDirectory
            typeList.add "weblink", VirtualDirectory
        End If
        On Error GoTo WebDocumentMethodLoadError
        Set nameattr = elm.Attributes.getNamedItem("name")
        If Not (nameattr Is Nothing) Then
            Set aWebMethod = XmlObjectList.Item(nameattr.value)
            If aWebMethod Is Nothing Then
                anamespace = elm.getAttribute("namespace")
                Set aWebMethod = aWeblink.DefineDocumentMethod(name, anamespace)
                Set aWebMethod.content = elm
                Set ParseElement = aWebMethod
                currentType = "webdocumentmethod"
            End If
        End If
    ElseIf UCase(elm.baseName) = "KAAMESSAGE" Then
        If (Not elm.selectSingleNode("./webdocumentmethod") Is Nothing) Then
            Set aWebMethod = ParseElement(elm.selectSingleNode("./webdocumentmethod"))
            If aWebMethod Is Nothing Then
            Else
                Set nameattr = elm.selectSingleNode("./webdocumentmethod").Attributes.getNamedItem("name")
                If nameattr Is Nothing Then
                Else
                    XmlObjectList.add aWebMethod, nameattr.value
                    nameList.add nameattr.value, nameattr.value
                    typeList.add currentType, nameattr.value
                End If
            End If
        End If
        On Error GoTo KAAMessageLoadError
        
        Set KAAMessage = New cXmlKAAMessage
        Set KAAMessage.content = elm
        Set ParseElement = KAAMessage
        Set KAAMessage.owner = Me
        currentType = "xmlcomarea"
    ElseIf UCase(elm.baseName) = "TOTALS" Then
        On Error GoTo TotalsLoadError
        Set totals = New cL2Totals
        Set totals.StaticLink = elm
        Set totals.Manager = Me
        Set ParseElement = totals
        currentType = "totals"
    ElseIf UCase(elm.baseName) = "CURRENCYMESSAGER" Then
        On Error GoTo CurrencyMessagerError
        Set curMessager = New cXmlCurrencyMessageHandler
        Set curMessager.Manager = Me
        Set curMessager.content = elm
        Set ParseElement = curMessager
        currentType = "currencymessager"
    ElseIf UCase(elm.baseName) = "DEPOSITMESSAGERVERSION3" Then
        On Error GoTo DepositMessagerError
        Set messagerversion3 = New cXmlDepositMessageHandlerVersion3
        Set messagerversion3.Manager = Me
        Set messagerversion3.content = elm
        Set ParseElement = messagerversion3
        currentType = "depositmessagerversion3"
    ElseIf UCase(elm.baseName) = "EXCHANGEDEPOSITMESSAGER" Then
        On Error GoTo ExchangeDepositMessagerError
        Set exchangemessager = New cXmlExchangeDepositMessageHandler
        Set exchangemessager.Manager = Me
        Set exchangemessager.content = elm
        Set ParseElement = exchangemessager
        currentType = "exchangedepositmessager"
    ElseIf UCase(elm.baseName) = "DEPOSITMASSIVEMESSAGER" Then
        On Error GoTo DepositMassiveMessagerError
        Set massmessager = New cXMLDepositMassiveMessageHandler
        Set massmessager.Manager = Me
        Set massmessager.content = elm
        Set ParseElement = massmessager
        currentType = "depositmassivemessager"
    ElseIf UCase(elm.baseName) = "DEPOSITMESSAGERVERSION4" Then
        On Error GoTo DepositMessagerError
        Set messagerversion4 = New cXmlDepositMessageHandlerVersion4
        Set messagerversion4.Manager = Me
        Set messagerversion4.content = elm
        Set ParseElement = messagerversion4
        currentType = "depositmessagerversion4"
    End If
    Exit Function
FormLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου form :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
FormUpdateLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου formupdate :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
ReportLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου report :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
EsignLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου esign :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
DocumentLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου document :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
CodeFormUpdateLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου codeformupdate :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
CurrencyPassbookError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου CurrencyPassbook :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
TrnLinkLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου trnlink :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
LinkL2TrnLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου l2trn :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
CallLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου call :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
CallLoadErrorAuth:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου call_auth :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
XmlHostMethodLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου HostMethod :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
TransformationLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου transformation :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
MethodLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου method :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
JobLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου job :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
JournalLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου journal :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
ValidationLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου validation :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
WebApiLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου webapi :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
WebLinkLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου weblink :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
WebMethodLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου webmethod :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
WebDocumentMethodLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου webdocument :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
ComAreaLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου comarea :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
KAAMessageLoadError:
    LogMsgbox "Λαθος κατά τη δημιουργία του αντικειμένου KAAMessage :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου..."
    Exit Function
TotalsLoadError:
    LogMsgbox "Λάθος κατά τη δημιουργία του αντικειμένου  totals :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου.... "
    Exit Function
DepositMessagerError:
    LogMsgbox "Λάθος κατά τη δημιουργία του αντικειμένου  depositmessanger :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου.... "
    Exit Function
ExchangeDepositMessagerError:
    LogMsgbox "Λάθος κατά τη δημιουργία του αντικειμένου  exchangedepositmessanger :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου.... "
    Exit Function
CurrencyMessagerError:
    LogMsgbox "Λάθος κατά τη δημιουργία του αντικειμένου  currencymessanger :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου.... "
    Exit Function
DepositMassiveMessagerError:
    LogMsgbox "Λάθος κατά τη δημιουργία του αντικειμένου  depositmassivemessanger :" & name & Err.number & " " & Err.description, vbCritical, "ΛΑΘΟΣ L2 Διαχείριση Εγγράφου.... "
    Exit Function

End Function

Public Sub LoadTemplates(rootnode As IXMLDOMElement)
    Dim Node As IXMLDOMNode, attr As IXMLDOMAttribute
    XmlObjectList.add GenWorkForm, "application"
    nameList.add "application", "application"
    typeList.add "genworkform", "application"
    If Not (owner Is Nothing) Then
        XmlObjectList.add owner, "form"
        nameList.add "form", "form"
        typeList.add "l2form", "form"
    'ElseIf Not (OwnerForm Is Nothing) Then
    '    XmlObjectList.Add OwnerForm, "form"
    End If
    
    
    If rootnode Is Nothing Then Exit Sub
    Set TemplateDocument = rootnode.ownerDocument
    Dim aname As String, aXmlObject
    
    For Each Node In rootnode.childNodes
        If Node.nodeTypeString = "element" Then
            Set attr = Node.Attributes.getNamedItem("name")
            If attr Is Nothing Then
            Else
                aname = attr.Text
                Set aXmlObject = ParseElement(Node)
                If aXmlObject Is Nothing Then
                Else
                    XmlObjectList.add aXmlObject, aname
                    nameList.add aname, aname
                    typeList.add currentType, aname
                End If
            End If
        End If
    Next Node
End Sub

Public Sub CleanUp()
    Dim name
    For Each name In nameList
        XmlObjectList.Remove CStr(name)
    Next name
    
    While nameList.count > 0
        nameList.Remove 1
    Wend
    While typeList.count > 0
        typeList.Remove 1
    Wend
    Dim Buffer
    If TrnBuffers Is Nothing Then
    Else
        If TrnBuffers.name <> GenWorkForm.AppBuffers.name Then
            TrnBuffers.ClearAll
        Else
            Dim apoint As String
            apoint = ""
        End If
        Set TrnBuffers = Nothing
    End If
    Set XmlObjectList = Nothing
    Set nameList = Nothing
    Set typeList = Nothing
    Set owner = Nothing
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub

