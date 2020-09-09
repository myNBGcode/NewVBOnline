Attribute VB_Name = "Mainmdl"
'Attribute VB_Name = "Initial"
Option Explicit

Public UseActiveDirectory As Boolean
Public WorkstationParams As cWorkStationParams
Public ExecutionResults As New Collection

Public xmlstation As cXmlWorkstation

Public OpenCobolServer As Boolean

Public ConnectionInfo As New Collection
Public DisconnectedTest As Stream

Public cTotalsVersion As Long

Public tmpSendViewName As String
Public tmpReceiveViewName As String

Public DebugSNAPoolLink As Boolean
Public ExecRuleB64ReceiveFile As String

Public LogIrisCom As Boolean

Public Const LB_SETHORIZONTALEXTENT = &H194

'TrnFrm Action Types
Public Const taNo_Action = 0
Public Const taGet_Input = 200
Public Const taSend_Buffer = 201
Public Const taPrint_Document = 202
Public Const taExit_Form = 203
Public Const taEscape_Form = 204
Public Const taStay_In_Form = 205

Public LocalFlag As Boolean

Public xmlEnvironment As New MSXML2.DOMDocument30
Public L2AddInFile As MSXML2.DOMDocument30

'Document definitions
Public Const DocumentLines = 55
Public DocLines(DocumentLines - 1) As String
Public LastDocLine As Integer
Public cVersion As Long

'---------------------------------------------------
Public Const ReservedControlPrefixes = ",FLD,SPD,LST,LBL,BTN,CMB,CHK,"

'---------------------------------------------------


Public LogonShare As String
Public cClientName As String
Public cClientIP As String
Public HasPad As String
Public MachineName As String
Public LogonServer As String
Public LogonDir As String
Public ReadDir As String
Public Const ComAreaDir = "ComArea\"
Public WorkDir As String
Public AuthDir As String
Public connect_status As Integer

Public gBoolStartingUp As Boolean
Public ASCII_CP_STRING As String, PASSBOOK_CLEAR_STRING As String
Public EBCDIC_CP_STRING As String

Public Strpin(15, 3) As String

Public cDebug As Integer
Public cBRANCH As String
Public cBRANCHIndex As String
Public cBRANCHName As String
Public cTERMINALID As String
Public cDepartment As String
Public cHEAD
Public cPOSTDATE As Date
Public cNextDateFlag

Public cTRNNum As Integer
Public cTRNCode As Long
Public cTRNTime As Double

Public cLaserDocumentsPrinter As String

Public cPassbookPrinter As Integer
Public cPrinterPort As Integer
Public cOCRPort As Integer
Public cListToPassbook As Integer
Public cPDC As String
Public cLogonServer As String
Public cUserName As String
Public cFullUserName As String
Public cHostUserName As String
Public cHostUserPassword As String
Public cJournalName As String
Public cBatchTotalsName As String

Public cPRINTERSERVER As String
Public cOCRREADERSERVER As String

Public cNewJournalType As Boolean
Public cUseCicsUserInfo As Boolean
Public cHasWinPanel As Boolean
Public cHasTellerTrnGroup As Boolean

Public cTELLERKEY As String * 1
Public cCHIEFKEY As String * 1
Public cCHIEFUserName As String
Public cMANAGERKEY As String * 1
Public cMANAGERUserName As String
Public cTELLERCHIEFKEY As String * 1
Public cTELLERMANAGERKEY As String * 1
Public cIRISAuthUserName As String
Public MachineList As New Collection
Public UserList As New Collection
Public UserKeysList As New Collection
Public IPList As New Collection
Public cANYKEY As String
Public Send0610 As String
Public SessID As Integer

Public EventLogWrite As Boolean
Public SendJournalWrite As Boolean
Public ReceiveJournalWrite As Boolean
Public SRJournal As Boolean

Public RequestFromMachine As String
Public RequestFromIP As String
Public ChiefRequest As Boolean
Public ManagerRequest As Boolean
Public AnyRequest As Boolean
Public SecretRequest As Boolean 'αιτηση για μυστικό chief teller
Public SecretValue As String    'μυστικός chief teller
Public KeyAccepted As Boolean

Public cEnableHiddenTransactions As Boolean

Public cBranchProfileName As String
Public cDefUserProfileName As String
Public cUserProfileName As String

Public cIRISUserName As String
Public cIRISComputerName As String
Public cIRISConnected As Boolean

Public cJournalTRNClass As Boolean 'ένδειξη αν υπάρχει στο journal η στήλη TRNClass
Public cTotalsTRNClass As Boolean 'ένδειξη αν υπάρχει στο Tbl_Totals, Tbl_Totals_Trace η στήλη TRNClass
Public cParamsTermUse As Boolean 'ένδειξη αν υπάρχει στο Tbl_Params, η στήλη TermUse

Public HasWinPanelConnection As Boolean

Public cSecretToken As String
Public cWebDavPath As String
Public cLocalEncryptedPath As String

Public CommunicationStarted As Boolean, ComTimerCounter As Integer, ComTimerCycle As Integer

Public Flag610 As Boolean, Flag620 As Boolean, Flag630 As Boolean
Public WorkEnvironment_ As String ' IRISEDUC για εκπαιδευτικό - IRISPROD για παραγωγή

Public NQCashierTicketID As String

Public Const cryptoPassword = "+sKib*drowSSaP-for=Shine"
    
Public LOCALUSERNAME As String
Public LOCALUSERPASSWORD As String
Public SQLSERVERUSERNAME As String
Public SQLSERVERUSERPASSWORD As String

Public cKMODEFlag As Boolean
Public cKMODEValue As String

Public WebLinks As New Collection

Public Type TRANSACTION_CONTROL_BLOCK
    BoolTransOk As Boolean
    curr_transaction As String      'Κωδικός συναλλαγής
    send_str As String              'String που θα σταλεί
    'initsend_str As String          'String που θα σταλεί πριν τη μετάφραση
    'send_str_length As Long         'μέγεθος -//-
    receive_str As String           'String απο host
    receive_str_length As Long      'μέγεθος -//-
    received_data As String
    read_again As Boolean
    MsgType As Long
    Ret1 As Long
    Ret2 As Long
    RetCode As Long
    LUADirection As Long
    TimeOut As Long
    ApplId As String
    com_debug As Long
    app_debug As Long
    send_convert As Long
    receive_convert As Long
    DecodeGreek As Long
    encodegreek As Long
    TransTerminating As Boolean
    CodePage As Integer
    inttime As Integer
End Type
    
Public Type PrinterText
    r As Integer
    c As Integer
    Text As String
End Type

Public cb As TRANSACTION_CONTROL_BLOCK
Public cbcomarea As cXmlComArea

Public ReceivedData As New Collection
Public G0Data As New Collection
'Public PrinterData As New Collection

Public Type IntegerPair
    p1 As Integer
    p2 As Integer
End Type

Public Type SelectionRow
    FldNo As Integer
    SelNo As Integer
    SelTxt As String
End Type

Public Type HelpLine
    LineCD As String
    LineText As String
End Type

'------------------------------------------------
Public HelpRetValue As String

Public TRNQueue As New Collection
Public TRNFldNoQueue As New Collection
Public TRNFldTextQueue As New Collection

Public gPanel As GlobalSPCPanel

Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal LPvalue As String) As Long
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub Main()
    
    GenWorkForm.Show
End Sub

Public Function UpdatexmlEnvironment(sHead As String, sBody As String)
    
    Dim envElm As IXMLDOMElement
    On Error Resume Next
    'If Left(sHead, 1) <> "-" Then
        If Not (xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))) Is Nothing) Then
            xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))).Text = sBody
        Else
            Set envElm = xmlEnvironment.createElement(UCase(Trim(sHead)))
            xmlEnvironment.documentElement.appendChild envElm
            envElm.Text = sBody
        End If
        
        Set envElm = Nothing
        
    'End If
End Function
Public Function UpdatexmlEnvironmentNode(eBody As IXMLDOMElement)
    On Error Resume Next
    If (xmlEnvironment.selectSingleNode("//" & Trim(eBody.nodename)) Is Nothing) Then
        ImportElement eBody, xmlEnvironment.documentElement
    End If
End Function

Public Function GetxmlEnvironment(sHead As String) As String
    
    Dim envElm As IXMLDOMElement
    On Error Resume Next
    'If Left(sHead, 1) <> "-" Then
        If Not (xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))) Is Nothing) Then
            GetxmlEnvironment = xmlEnvironment.selectSingleNode("//" & UCase(Trim(sHead))).Text
        Else
            GetxmlEnvironment = ""
        End If
        
    'End If
End Function

Public Function ImportElement(source As IXMLDOMElement, destinationparent)
    Dim newElm As IXMLDOMElement
    If destinationparent.nodeType = NODE_DOCUMENT Then
        Set newElm = destinationparent.createElement(source.baseName)
    Else
        Set newElm = destinationparent.ownerDocument.createElement(source.baseName)
    End If
    destinationparent.appendChild newElm
    Dim oldattr As IXMLDOMAttribute, newAttr As IXMLDOMAttribute
    For Each oldattr In source.Attributes
         Set newAttr = newElm.ownerDocument.createAttribute(oldattr.name)
         newAttr.value = oldattr.value
         newElm.Attributes.setNamedItem newAttr
    Next oldattr
    Dim sourcechild As IXMLDOMNode
    For Each sourcechild In source.childNodes
        If sourcechild.nodeType = NODE_ELEMENT Then
            ImportElement sourcechild, newElm
        ElseIf sourcechild.nodeType = NODE_TEXT Then
            newElm.Text = sourcechild.Text
        End If
    Next sourcechild
End Function

Public Sub UpdateChiefKey(Chief As String)
    UpdatexmlEnvironment UCase("ChiefTellerUsername"), UCase(Chief)
    cCHIEFUserName = Chief
End Sub

Public Sub UpdateManagerKey(Manager As String)
    UpdatexmlEnvironment UCase("ManagerUsername"), UCase(Manager)
    cMANAGERUserName = Manager
End Sub

Public Sub ShowStatusMessage(Message As String)
    Dim Control
    Dim controlfound As Boolean
    For Each Control In Screen.activeform.Controls
        If TypeOf Control Is StatusBar Then
            controlfound = True
            Control.Panels(1).Text = Message
            On Error Resume Next
            Control.SimpleText = Message
            On Error GoTo 0
            Exit For
        End If
    Next Control
    If (controlfound = False) Then MsgBox Message, vbOKOnly, "Ειδοποίηση"
End Sub

Public Function CreateTokenFile(filename As String) As String
    
    Dim token As String
    token = GetGuid
    
    On Error GoTo FileError
    Open filename For Output As #1
    Print #1, token
    Close #1
    
    CreateTokenFile = token
    Exit Function

FileError:
    CreateTokenFile = ""
    Exit Function
End Function

Public Function WebDavUpload(WebDavPath As String, filename As String) As Boolean
    
    On Error GoTo UpError
    
    Dim fso
    Dim infile
    Dim destinationfilename
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set infile = fso.GetFile(filename)
    destinationfilename = WebDavPath & "\" & infile.name
    
    infile.Copy destinationfilename
    
    WebDavUpload = True
    GoTo EndFunc

UpError:
    WebDavUpload = False
    GoTo EndFunc

EndFunc:
    Set infile = Nothing
    Set fso = Nothing

End Function

Public Function ReadSecretToken(WebDavPath As String, filename As String) As String

    On Error GoTo ReadError

    Dim fso
    Dim objTextFile
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = fso.OpenTextFile(WebDavPath & "\" & filename, 1)
    Dim content As String
    content = objTextFile.ReadAll
    objTextFile.Close
    
    ReadSecretToken = content
    GoTo EndFunc

ReadError:
    ReadSecretToken = ""
    GoTo EndFunc

EndFunc:
    Set objTextFile = Nothing
    Set fso = Nothing

End Function

