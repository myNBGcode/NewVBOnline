VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlWorkstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim terminalDoc As New MSXML2.DOMDocument30
Dim computersNode As IXMLDOMElement
Dim computername_  As IXMLDOMElement
Dim branch_  As IXMLDOMElement
Dim branchname_  As IXMLDOMElement
Dim passbookprinter_  As IXMLDOMElement
Dim listtopassbook_  As IXMLDOMElement
Dim printerserver_  As IXMLDOMElement
'Dim smtpserver_  As IXMLDOMElement
'Dim smtpport_  As IXMLDOMElement
Dim newterminalid_  As IXMLDOMElement
Dim pdc_  As IXMLDOMElement
Dim debug_  As IXMLDOMElement
Dim ocrreaderserver_  As IXMLDOMElement
Dim branchprofilename_  As IXMLDOMElement
Dim iriscomputername_  As IXMLDOMElement
Dim irisusername_  As IXMLDOMElement
Dim debugsnapoollink_  As IXMLDOMElement
Dim execruleb64receivefile_  As IXMLDOMElement
Dim printerport_  As IXMLDOMElement
Dim ocrport_  As IXMLDOMElement
Dim opencobol_  As IXMLDOMElement
Dim useactivedirectory_  As IXMLDOMElement
Dim setmanager_  As IXMLDOMElement
Dim setchief_  As IXMLDOMElement
Dim setteller_  As IXMLDOMElement
Dim journaltype_  As IXMLDOMElement
Dim lasttrnnumber_ As IXMLDOMElement
Dim cicsuserinfo_ As IXMLDOMElement
Dim ishostteller_ As IXMLDOMElement
Dim ishostchief_ As IXMLDOMElement
Dim ishostmanager_ As IXMLDOMElement
Dim haswinpanel_ As IXMLDOMElement
Dim tellertrn_ As IXMLDOMElement
Dim haspad_ As IXMLDOMElement

Private Sub Class_Initialize()
    Set computersNode = Nothing
End Sub

Public Function Initialize(value As String) As cXmlWorkstation
    Dim station As cXmlWorkstation
    
        Dim adoc As MSXML2.DOMDocument30
        Set adoc = New MSXML2.DOMDocument30
        adoc.LoadXML value
        If adoc Is Nothing Then Exit Function
        
        Set computersNode = GetXmlNode(adoc.documentElement, "//ODATA/Computers ", "Computers ", "ODATA", "Πρόβλημα στη διαδικασία Έναρξης...")
        If computersNode Is Nothing Then Exit Function
        
        terminalDoc.LoadXML computersNode.XML
        
        Set station = New cXmlWorkstation
        station.Document = terminalDoc
        station.InitializeNodes
        Set Initialize = station
        Set station = Nothing


End Function

 Sub InitializeNodes()

    Set computersNode = GetXmlNode(terminalDoc.documentElement, "//Computers ", "Computers", "ComputersDocument", "Πρόβλημα στη διαδικασία Έναρξης...")
    If computersNode Is Nothing Then Exit Sub
    Set computername_ = GetXmlNode(computersNode, "./_COMPUTERNAME")
    Set branch_ = GetXmlNodeIfPresent(computersNode, "./_BRANCH")
    Set branchname_ = GetXmlNodeIfPresent(computersNode, "./_BRANCHNAME")
    Set newterminalid_ = GetXmlNodeIfPresent(computersNode, "./_NEWTERMINALID")
    Set passbookprinter_ = GetXmlNodeIfPresent(computersNode, "./_PASSBOOKPRINTER")
    Set listtopassbook_ = GetXmlNodeIfPresent(computersNode, "./_LISTTOPASBOOK")
    Set pdc_ = GetXmlNodeIfPresent(computersNode, "./_PDC")
    Set debug_ = GetXmlNodeIfPresent(computersNode, "./_DEBUG")
'    Set smtpserver_ = GetXmlNodeIfPresent(computersNode, "./_SMTPSERVER")
'    Set smtpport_ = GetXmlNodeIfPresent(computersNode, "./_SMTPPORT")
    Set printerserver_ = GetXmlNodeIfPresent(computersNode, "./_PRINTERSERVER")
    Set ocrreaderserver_ = GetXmlNodeIfPresent(computersNode, "./_OCRREADERSERVER")
    Set branchprofilename_ = GetXmlNodeIfPresent(computersNode, "./_BRANCHPROFILENAME")
    Set iriscomputername_ = GetXmlNodeIfPresent(computersNode, "./_IRISCOMPUTERNAME")
    Set irisusername_ = GetXmlNodeIfPresent(computersNode, "./_IRISUSERNAME")
    Set debugsnapoollink_ = GetXmlNodeIfPresent(computersNode, "./_DEBUGSNAPOOLLINK")
    Set execruleb64receivefile_ = GetXmlNodeIfPresent(computersNode, "./_EXECRULEB64RECEIVEFILE")
    Set printerport_ = GetXmlNodeIfPresent(computersNode, "./_PRINTERPORT")
    Set ocrport_ = GetXmlNodeIfPresent(computersNode, "./_OCRPORT")
    Set opencobol_ = GetXmlNodeIfPresent(computersNode, "./_OPENCOBOL")
    Set useactivedirectory_ = GetXmlNodeIfPresent(computersNode, "./_USEACTIVEDIRECTORY")
    Set setmanager_ = GetXmlNodeIfPresent(computersNode, "./_SETMANAGER")
    Set setchief_ = GetXmlNodeIfPresent(computersNode, "./_SETCHIEF")
    Set setteller_ = GetXmlNodeIfPresent(computersNode, "./_SETTELLER")
    Set journaltype_ = GetXmlNodeIfPresent(computersNode, "./_JOURNALTYPE")
    Set lasttrnnumber_ = GetXmlNodeIfPresent(computersNode, "./_LASTTRNNUMBER")
    Set cicsuserinfo_ = GetXmlNodeIfPresent(computersNode, "./_CICSUSERINFO")
    Set ishostteller_ = GetXmlNodeIfPresent(computersNode, "./_ISHOSTTELLER")
    Set ishostchief_ = GetXmlNodeIfPresent(computersNode, "./_ISHOSTCHIEF")
    Set ishostmanager_ = GetXmlNodeIfPresent(computersNode, "./_ISHOSTMANAGER")
    Set haswinpanel_ = GetXmlNodeIfPresent(computersNode, "./_HASWINPANEL")
    Set tellertrn_ = GetXmlNodeIfPresent(computersNode, "./_TELLERTRN")
    Set haspad_ = GetXmlNodeIfPresent(computersNode, "./_HASPAD")

End Sub

Public Property Get Document() As MSXML2.DOMDocument30
   Set Document = terminalDoc
End Property
Public Property Let Document(value As MSXML2.DOMDocument30)
    Set terminalDoc = value
End Property

Public Property Get ComputerName() As IXMLDOMElement
    Set ComputerName = computername_
End Property

Public Property Let ComputerName(value As IXMLDOMElement)
    Set computername_ = value
End Property
Public Property Get branch() As IXMLDOMElement
    Set branch = branch_
End Property

Public Property Let branch(value As IXMLDOMElement)
   Set branch_ = value
End Property

Public Property Get BranchName() As IXMLDOMElement
  Set BranchName = branchname_
End Property

Public Property Let BranchName(value As IXMLDOMElement)
    branchname_ = value
End Property
Public Property Get NewTerminalId() As IXMLDOMElement
    Set NewTerminalId = newterminalid_
End Property

Public Property Let NewTerminalId(value As IXMLDOMElement)
  Set newterminalid_ = value
End Property

Public Property Get PassbookPrinter() As IXMLDOMElement
    Set PassbookPrinter = passbookprinter_
End Property

Public Property Let PassbookPrinter(value As IXMLDOMElement)
    Set passbookprinter_ = value
End Property

Public Property Get ListToPassbook() As IXMLDOMElement
    Set ListToPassbook = listtopassbook_
End Property

Public Property Let ListToPassbook(value As IXMLDOMElement)
    Set listtopassbook_ = value
End Property

Public Property Get PDC() As IXMLDOMElement
    Set PDC = pdc_
End Property

Public Property Let PDC(value As IXMLDOMElement)
    Set pdc_ = value
End Property


Public Property Get Debugg() As IXMLDOMElement
    Set Debugg = debug_
End Property

Public Property Let Debugg(value As IXMLDOMElement)
    Set debug_ = value
End Property


'Public Property Get SmtpServer() As IXMLDOMElement
'    Set SmtpServer = smtpserver_
'End Property
'
'Public Property Let SmtpServer(value As IXMLDOMElement)
'    Set smtpserver_ = value
'End Property
'
'Public Property Get SmtpPort() As IXMLDOMElement
'    Set SmtpPort = smtpport_
'End Property
'
'Public Property Let SmtpPort(value As IXMLDOMElement)
'    Set smtpport_ = value
'End Property
'
'


Public Property Get PrinterServer() As IXMLDOMElement
    Set PrinterServer = printerserver_
End Property

Public Property Let PrinterServer(value As IXMLDOMElement)
    Set printerserver_ = value
End Property

Public Property Get OcrReaderServer() As IXMLDOMElement
    Set OcrReaderServer = ocrreaderserver_
End Property

Public Property Let OcrReaderServer(value As IXMLDOMElement)
    Set ocrreaderserver_ = value
End Property

Public Property Get BranchProfileName() As IXMLDOMElement
   Set BranchProfileName = branchprofilename_
End Property

Public Property Let BranchProfileName(value As IXMLDOMElement)
    Set branchprofilename_ = value
End Property

Public Property Get IRISComputerName() As IXMLDOMElement
    Set IRISComputerName = iriscomputername_
End Property

Public Property Let IRISComputerName(value As IXMLDOMElement)
    Set iriscomputername_ = value
End Property

Public Property Get IRISUSERName() As IXMLDOMElement
    Set IRISUSERName = irisusername_
End Property

Public Property Let IRISUSERName(value As IXMLDOMElement)
    Set irisusername_ = value
End Property

Public Property Get DebugSNAPoolLink() As IXMLDOMElement
   Set DebugSNAPoolLink = debugsnapoollink_
End Property

Public Property Let DebugSNAPoolLink(value As IXMLDOMElement)
    Set debugsnapoollink_ = value
End Property

Public Property Get ExecRuleB64ReceiveFile() As IXMLDOMElement
    Set ExecRuleB64ReceiveFile = execruleb64receivefile_
End Property

Public Property Let ExecRuleB64ReceiveFile(value As IXMLDOMElement)
    Set execruleb64receivefile_ = value
End Property

Public Property Get PrinterPort() As IXMLDOMElement
    Set PrinterPort = printerport_
End Property

Public Property Let PrinterPort(value As IXMLDOMElement)
    Set printerport_ = value
End Property

Public Property Get OcrPort() As IXMLDOMElement
    Set OcrPort = ocrport_
End Property

Public Property Let OcrPort(value As IXMLDOMElement)
    Set ocrport_ = value
End Property

Public Property Get OpenCobol() As IXMLDOMElement
    Set OpenCobol = opencobol_
End Property

Public Property Let OpenCobol(value As IXMLDOMElement)
    Set opencobol_ = value
End Property

Public Property Get UseActiveDirectory() As IXMLDOMElement
    Set UseActiveDirectory = useactivedirectory_
End Property

Public Property Let UseActiveDirectory(value As IXMLDOMElement)
    Set useactivedirectory_ = value
End Property
Public Property Get SetManager() As IXMLDOMElement
    Set SetManager = setmanager_
End Property

Public Property Let SetManager(value As IXMLDOMElement)
    Set setmanager_ = value
End Property
Public Property Get SetChief() As IXMLDOMElement
    Set SetChief = setchief_
End Property

Public Property Let SetChief(value As IXMLDOMElement)
    Set setchief_ = value
End Property
Public Property Get SetTeller() As IXMLDOMElement
    Set SetTeller = setteller_
End Property

Public Property Let SetTeller(value As IXMLDOMElement)
    Set setteller_ = value
End Property

Public Property Get JournalType() As IXMLDOMElement
    Set JournalType = journaltype_
End Property

Public Property Let JournalType(value As IXMLDOMElement)
    Set journaltype_ = value
End Property

Public Property Get LastTrnNumber() As IXMLDOMElement
    Set LastTrnNumber = lasttrnnumber_
End Property

Public Property Let LastTrnNumber(value As IXMLDOMElement)
    Set lasttrnnumber_ = value
End Property

Public Property Get CicsUserInfo() As IXMLDOMElement
    Set CicsUserInfo = cicsuserinfo_
End Property

Public Property Let CicsUserInfo(value As IXMLDOMElement)
    Set cicsuserinfo_ = value
End Property

Public Property Get IsHostTeller() As IXMLDOMElement
    Set IsHostTeller = ishostteller_
End Property

Public Property Let IsHostTeller(value As IXMLDOMElement)
    Set ishostteller_ = value
End Property

Public Property Get IsHostChief() As IXMLDOMElement
    Set IsHostChief = ishostchief_
End Property

Public Property Let IsHostChief(value As IXMLDOMElement)
    Set ishostchief_ = value
End Property

Public Property Get IsHostManager() As IXMLDOMElement
    Set IsHostManager = ishostmanager_
End Property

Public Property Let IsHostManager(value As IXMLDOMElement)
    Set ishostmanager_ = value
End Property

Public Property Get HasWinPanel() As IXMLDOMElement
    Set HasWinPanel = haswinpanel_
End Property

Public Property Let HasWinPanel(value As IXMLDOMElement)
    Set haswinpanel_ = value
End Property

Public Property Get TellerTrn() As IXMLDOMElement
    Set TellerTrn = tellertrn_
End Property

Public Property Let TellerTrn(value As IXMLDOMElement)
    Set tellertrn_ = value
End Property

Public Property Get HasPad() As IXMLDOMElement
    Set HasPad = haspad_
End Property

Public Property Let HasPad(value As IXMLDOMElement)
    Set haspad_ = value
End Property

Private Sub Class_Terminate()

    Set computername_ = Nothing
    Set branch_ = Nothing
    Set branchname_ = Nothing
    Set passbookprinter_ = Nothing
    Set listtopassbook_ = Nothing
    Set printerserver_ = Nothing
    'Set smtpserver_ = Nothing
    'Set smtpport_ = Nothing
    Set newterminalid_ = Nothing
    Set pdc_ = Nothing
    Set debug_ = Nothing
    Set ocrreaderserver_ = Nothing
    Set branchprofilename_ = Nothing
    Set iriscomputername_ = Nothing
    Set irisusername_ = Nothing
    Set debugsnapoollink_ = Nothing
    Set execruleb64receivefile_ = Nothing
    Set printerport_ = Nothing
    Set ocrport_ = Nothing
    Set opencobol_ = Nothing
    Set useactivedirectory_ = Nothing
    Set setmanager_ = Nothing
    Set setchief_ = Nothing
    Set setteller_ = Nothing
    Set journaltype_ = Nothing
    Set lasttrnnumber_ = Nothing
    Set cicsuserinfo_ = Nothing
    Set ishostteller_ = Nothing
    Set ishostchief_ = Nothing
    Set ishostmanager_ = Nothing
    Set haswinpanel_ = Nothing
    Set tellertrn_ = Nothing
    Set haspad_ = Nothing
    
    Set terminalDoc = Nothing
    Set computersNode = Nothing

End Sub
