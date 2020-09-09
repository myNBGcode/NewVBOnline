VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form TRNFrm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Transaction Form"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer RefreshCom 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   240
      Top             =   2880
   End
   Begin MSComctlLib.ImageList CommandImages 
      Left            =   150
      Top             =   3630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TRNFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TRNFrm.frx":01D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TRNFrm.frx":03A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TRNFrm.frx":0576
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TRNFrm.frx":0748
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar CommandToolbar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   688
      ButtonWidth     =   3149
      ButtonHeight    =   688
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "CommandImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Εκτύπωση"
            Key             =   "F9"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ημερολόγιο"
            Key             =   "F10"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Σύνολα"
            Key             =   "F11"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Διαβίβαση"
            Key             =   "F12"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6270
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   20108
            MinWidth        =   20108
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "TRNFrm.frx":13CA
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "TRNFrm.frx":1616
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl ValidationControl 
      Left            =   120
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      Timeout         =   60000
   End
   Begin VB.Menu MNUItem 
      Caption         =   "MNUITem"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu MNUSub1 
         Caption         =   "MNUSub1"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem2 
      Caption         =   "MNUItem2"
      Visible         =   0   'False
      Begin VB.Menu MNUSub2 
         Caption         =   "MNUSub2"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem3 
      Caption         =   "MNUItem3"
      Visible         =   0   'False
      Begin VB.Menu MNUSub3 
         Caption         =   "MNUSub3"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem4 
      Caption         =   "MNUItem4"
      Visible         =   0   'False
      Begin VB.Menu MNUSub4 
         Caption         =   "MNUSub4"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem5 
      Caption         =   "MNUItem5"
      Visible         =   0   'False
      Begin VB.Menu MNUSub5 
         Caption         =   "MNUSub5"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem6 
      Caption         =   "MNUItem6"
      Visible         =   0   'False
      Begin VB.Menu MNUSub6 
         Caption         =   "MNUSub6"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem7 
      Caption         =   "MNUItem7"
      Visible         =   0   'False
      Begin VB.Menu MNUSub7 
         Caption         =   "MNUSub7"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem8 
      Caption         =   "MNUItem8"
      Visible         =   0   'False
      Begin VB.Menu MNUSub8 
         Caption         =   "MNUSub8"
         Index           =   0
      End
   End
   Begin VB.Menu MNUItem9 
      Caption         =   "MNUItem9"
      Visible         =   0   'False
      Begin VB.Menu MNUSub9 
         Caption         =   "MNUSub9"
         Index           =   0
      End
   End
End
Attribute VB_Name = "TRNFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Module28PoolLink As Boolean

Private xmlDocumentManager As cXMLDocumentManager
Private xmlTransformations As New Collection

Const fopNoOperation = 0
Const fopExitForm = 1           'Εξοδος απο την οθόνη ή μεταβαση σε επόμενη φάση ή παραμονή στην οθόνη αναλογα με το processphase, autoexit, restartflag
Const fopMoveNextPhase = 2
Const fopSendBuffer = 3
Const fopCloseForm = 4          'Εξοδος απο την οθόνη ανεξάρτητα παο processphase ή οτι άλλο

'ΝΑ ΓΙΝΕΤΕ ΕΛΕΓΧΟΣ ΓΙΑ ΜΗ ΥΠΑΡΚΤΟΥΣ ΑΘΡΟΙΣΤΕΣ

Dim trnXMLVersion As String
Dim trnXML As New MSXML2.DOMDocument
Dim rootnode As MSXML2.IXMLDOMElement
Dim trnNode As MSXML2.IXMLDOMElement
Dim stepsNode As MSXML2.IXMLDOMElement
Dim stepnode As MSXML2.IXMLDOMElement
Dim fieldsNode As MSXML2.IXMLDOMElement
Dim fieldNode As MSXML2.IXMLDOMElement
Dim labelsNode As MSXML2.IXMLDOMElement
Dim labelNode As MSXML2.IXMLDOMElement
Dim listsNode As MSXML2.IXMLDOMElement
Dim listNode As MSXML2.IXMLDOMElement
Dim gridsNode As MSXML2.IXMLDOMElement
Dim gridNode As MSXML2.IXMLDOMElement
Dim btnsNode As MSXML2.IXMLDOMElement
Dim btnNode As MSXML2.IXMLDOMElement
Dim chksNode As MSXML2.IXMLDOMElement
Dim chkNode As MSXML2.IXMLDOMElement
Dim cmbsNode As MSXML2.IXMLDOMElement
Dim cmbNode As MSXML2.IXMLDOMElement
Dim chrNode As MSXML2.IXMLDOMElement
Dim chrsNode As MSXML2.IXMLDOMElement

Public StartTime
Public EndTime
Public StartTickCount
Public EndTickCount
Public trn_key As String

Dim SpecialKey As String                    'Κλειδι με το οποίο θα φύγει η συναλλαγή αν πατηθεί Ctrl-C ή Ctrl-M
Dim docPrinting As Boolean

Public SPCPanel

Public SelectedTRN As Integer, BonusScale As Integer, BonusRegPhase As Integer, BonusRegPos As Integer
Public TRNOk As Boolean
Private ExitPhaseStarted As Boolean        'χρειάζεται για να απενεργοποιεί τον έλεγχο πεδίων στο Esc
Private DisableControlLostFocus As Boolean 'χρειάζεται για να απενεργοποιεί τον έλεγχο πεδίων στο φόρτωμα της οθόνης

Public CloseTransactionFlag As Boolean

'----------------------------------------------------------
' ενδείξεις για τα τμήματα κώδικα που περιέχει η συναλλαγή
'----------------------------------------------------------
Private StartupCodeFlag As Boolean
Private FormValidationFlag As Boolean
Private BeforeOutFlag As Boolean
Private AfterInFlag As Boolean
Private CommunicationErrorFlag As Boolean
Private BeforeActionFlag As Boolean
Private AfterActionFlag As Boolean
Private AfterKeyFlag As Boolean
Private BeforePrintFlag As Boolean
Private BeforeDocumentFlag As Boolean
Private AfterDocumentFlag As Boolean
Private LostFocusFlag As Boolean
'----------------------------------------------------------

Public EditFieldsCount As Integer   'ενημερώνει το textbox για τον αριθμό των ενεργών πεδίων της οθόνης
Private LastValidChk As Boolean, LastChkControl As GenTextBox

Public ActiveTextBox As GenTextBox
Public ActiveListBox As GenListBox
Public ActiveSpread As GenSpread

Public NamedControls As New Collection
Public fields As New Collection, NamedFields As New Collection
Public OutFields As New Collection
Public DocFields As New Collection

Public Buttons As New Collection
Public Checks As New Collection
Public Combos As New Collection

Public Labels As New Collection
Public Lists As New Collection
Public Spreads As New Collection
Public Charts As New Collection
Public Browsers As New Collection
Public RichTextBoxes As New Collection

Public TabbedControls As New Collection

Public StartupFailed As Boolean, StartupError As String
Public ValidOk As Boolean, ValidationError As String
Public ChangeFocusOk As Boolean, ChangeFocusError As String
Public BeforeOutFailed As Boolean, BeforeOutError As String
Public AfterInFailed As Boolean, AfterInError As String
Public CommunicationErrorFailed As Boolean, CommunicationErrorError As String
Public DisableEnterKey As Boolean
Public CancelCommunicationFlag As Boolean
Public DisableWriteJournal As Boolean
Public CancelPrintFlag As Boolean
Public PrintPromptMessage As String
Public SkipCommConfirmation As Boolean 'Απενεργοποιεί την ερώτηση Ναι/Οχι στην επικοινωνία

Private EncodeGreekflag As Boolean
Private FldLengthInBuffer As Boolean
Public AutoExitFlag As Boolean
Public RestartEditFlag As Boolean
Public TotalName As String, TotalPos As Integer
Private SendStarted As Boolean 'ένδειξη άν έχει πατηθει F12 ή έχει αρχίσει το send για να αποφύγουμε διπλό send
Private HideSendFromJournal As Boolean, HideReceiveFromJournal As Boolean 'ένδειξη ακύρωσης καταγραφής buffer στο ημερολόγιο
Public SkipKeyChk As Boolean 'Απενεργοποιεί τον έλεγχο κλειδιού
Public HiddenFlag As Boolean 'Ένδειξη αν η συναλλαγή φαίνεται στο menu

Public KeyProcessStarted As Boolean

'-----------------------------------------------------------

Public ListData As Collection 'λιστα με τα strings που έλαβε το τερματικό από το ΚΜ στη διάρκεια της συναλλαγής
Public ListG0 As Collection   'λιστα με τα G0 strings που έλαβε το τερματικό από το ΚΜ στη διάρκεια της συναλλαγής

'backup πεδια για χρήση μεσα απο τις συναλλαγές
Public BackValue1 As String
Public BackValue2 As String
Public BackValue3 As String
Public BackValue4 As String
Public BackValue5 As String

'Public ListLine As Integer, PageLine As Integer, LineOffset As Integer
'Public Counter1 As Integer, Counter2 As Integer, Counter3 As Integer
'Public Counter4 As Integer, Counter5 As Integer
'Public sum1, sum2, sum3, sum4, sum5, sum6, sum7, sum8, sum9, sum10 As Double

'-----------------------------------------------------------
Private TrnCode(10) As String, MaxPhaseNum As Integer
Public CurrAction As Integer, NextAction As Integer, processphase As Integer
Public QTrn As Integer

Public TrnBuffers As Buffers, AppLevelOutBuffer As String, AppLevelInBuffer As String, OpAfterHandling As Integer, ActivateOp As Integer
Public TrnVariables As New Collection

Public paramnames As String, Params As String, AppBuffersPos As Long, AppVariablesPos As Long
Public PNameArray, PArray 'array με τα ονοματα των παραμέτρων τησ οθόνης και array με τις παραμέτρους

Public OwnerForm As TRNFrm
Public AppRSPos As Long, AppSPPos As Long, AppRS_SPos As Long
Public AppCRSPos As Long


Private aTotalEntries As TotalEntries
'Public GroupUsersDoc As MSXML2.DOMDocument
Public IRISAuthError As String
Public UseIRISUpdateFiles As Boolean

Public amsgmemberconstructor As msgmemberwsconstructor
Public amsgwrapperconstructor As msgwrapperwsconstructor

Public DisableTRNCounterUpdate As Boolean

Public Sub DoEvents_()
    DoEvents
End Sub


Public Function GetxmlEnvironment() As MSXML2.DOMDocument30
    Set GetxmlEnvironment = xmlEnvironment
End Function

Public Function GetBankName(BankCode) As String
    GetBankName = GetBankName_(CInt(BankCode))
End Function

Public Sub CLSendCheques(SDate As Date, EDate As Date)
End Sub

Public Function ISOTOCURR(ByVal inUnit) As String
     ISOTOCURR = ISOTOCURR_(CStr(inUnit))
End Function

Public Function CURRTOISO(ByVal inUnit) As String
     CURRTOISO = CURRTOISO_(CStr(inUnit))
End Function

Public Function gSqlDate(indate) As String
    On Error GoTo OnError
    gSqlDate = "'" & Year(indate) & "-" & Month(indate) & "-" & Day(indate) & "'"
    Exit Function
OnError: gSqlDate = ""
End Function

Public Function gDateF8(indate) As String
    On Error GoTo OnError
    gDateF8 = Right("00" & CStr(Day(indate)), 2) & "/" & _
             Right("00" & CStr(Month(indate)), 2) & "/" & _
             Right("0000" & CStr(Year(indate)), 4)
    Exit Function
OnError: gDateF8 = ""
End Function

Public Function gDateU8(indate) As String
    On Error GoTo OnError
    gDateU8 = Right("00" & CStr(Day(indate)), 2) & _
             Right("00" & CStr(Month(indate)), 2) & _
             Right("0000" & CStr(Year(indate)), 4)
    Exit Function
OnError: gDateU8 = ""
End Function

Public Function gFormat(FormatString, inParams)
    gFormat = gFormat_(FormatString, inParams)
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
End Function

Public Sub ShowIRISMessages(inMessageView)
    ShowIRISMessages_ inMessageView
End Sub

Public Function bIIf(Expression, TruePart, FalsePart)
    If Expression Then bIIf = TruePart Else bIIf = FalsePart
End Function

Public Function Format_(invalue, InFormat As String)
    Format_ = format(invalue, InFormat)
End Function

Public Property Get TrnVariable(inName As String)
    Dim i As Integer
    TrnVariable = ""
    For i = 1 To TrnVariables.count
        If UCase(TrnVariables.item(i).name) = UCase(inName) Then
            TrnVariable = TrnVariables.item(i).value
            Exit Property
        End If
    Next i
End Property

Public Function TrnVariableDouble(inName As String) As Double
    Dim ares As String
    ares = TrnVariable(inName)
    TrnVariableDouble = 0
    On Error GoTo TrnVariableDoubleError
    TrnVariableDouble = CDbl(ares)
    Exit Function
TrnVariableDoubleError:
End Function

Public Function TrnVariableInteger(inName As String) As Long
    Dim ares As String
    ares = TrnVariable(inName)
    TrnVariableInteger = 0
    On Error GoTo TrnVariableIntegerError
    TrnVariableInteger = CDbl(ares)
    Exit Function
TrnVariableIntegerError:
End Function

Public Property Let TrnVariable(inName As String, invalue)
    Dim i As Integer
    For i = 1 To TrnVariables.count
        If UCase(TrnVariables.item(i).name) = UCase(inName) Then
            TrnVariables.item(i).value = invalue
            Exit Property
        End If
    Next i
    Dim aVariable As New VariableEntry
    aVariable.name = inName
    aVariable.value = invalue
    TrnVariables.add aVariable
    
    If UCase(inName) = "CANCELCOMMUNICATIONFLAG" Then
        If UCase(invalue) = UCase("true") Then
            CancelCommunicationFlag = True
        ElseIf UCase(invalue) = UCase("false") Then
            CancelCommunicationFlag = False
        End If
    End If
End Property

Public Property Get AppVariable(inName As String)
    AppVariable = AppVariable_(inName)
End Property

Public Property Let AppVariable(inName As String, invalue)
    AppVariable_(inName) = invalue
End Property

Public Function GetInCur() As String
'ΛΙΣΤΑ ΜΕ IN ΝΟΜΙΣΜΑΤΑ
    GetInCur = GetInCur_
End Function

Public Function NVLDouble(invalue, retValue As Double) As Double
    NVLDouble = NVLDouble_(invalue, retValue)
End Function

Public Function NVLInteger(invalue, retValue As Integer) As Long
    NVLInteger = NVLInteger_(invalue, retValue)
End Function

Public Function NVLString(invalue, retValue As String) As String
    NVLString = NVLString_(invalue, retValue)
End Function

Public Function NVLBoolean(invalue, retValue As Boolean) As Boolean
    NVLBoolean = NVLBoolean_(invalue, retValue)
End Function

Public Function NVLDate(invalue, retValue As Date) As Date
    NVLDate = NVLDate_(invalue, retValue)
End Function

Public Function GetWorkEnvironment() As String
    GetWorkEnvironment = WorkEnvironment_
End Function

Public Function GetComputerName() As String
    GetComputerName = MachineName
End Function

Public Function ChkValidIBAN(IBANaccount) As Integer
    ChkValidIBAN = ChkValidIBAN_(CStr(IBANaccount))
End Function


Public Function CreateIBAN(branch, account) As String
    CreateIBAN = CreateIBAN_(CStr(branch), CStr(account))
End Function

Public Function FormatIBAN(IBAN As String) As String
    FormatIBAN = FormatIBAN_(IBAN)
End Function

Public Sub Disconnect()
   DISCONNECT_
End Sub

Public Sub DisableTabForControl(sender)
    'sender.TabStop = False
End Sub

Public Sub EnableTabForControl(sender)
    'sender.TabStop = False
End Sub

Public Sub UnLockPrinter()
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then SPCPanel.UnLockPrinter
End Sub

Public Sub sbShowCommStatus(ByVal bActive As Boolean)
    Status.Panels(2).Visible = bActive
    Status.Panels(3).Visible = Not bActive
    If bActive Then Screen.MousePointer = vbDefault Else Screen.MousePointer = vbArrowHourglass
End Sub

Public Function CalcCd1(Acc, Digits) As Integer
    CalcCd1 = CalcCd1_(CStr(Acc), CInt(Digits))
End Function

Public Function CalcCd2(Acc) As Integer
    CalcCd2 = CalcCd2_(CStr(Acc))
End Function

Public Function CalcCd1_4330(Acc, Digits) As Integer
    CalcCd1_4330 = CalcCd1_4330_(CStr(Acc), CInt(Digits))
End Function

Public Function ChkTaxID(inNum) As Boolean
    ChkTaxID = ChkTaxID_(CStr(inNum))
End Function

Public Function ChkFldType(invalue, inValidationCode) As Boolean
    ChkFldType = ChkFldType_(CStr(invalue), CInt(inValidationCode))
End Function

Public Function ChkGenBankCheque(inCheque) As Boolean
    ChkGenBankCheque = ChkGenBankCheque_(CStr(inCheque))
End Function

Public Function ChkETECheque(inNum) As Boolean
'Έλεγχος αριθμού επιταγής ΕΤΕ
On Error GoTo ErrorPos
    ChkETECheque = ChkETECheque_(CLng(inNum))
    
    Exit Function
ErrorPos:
    ChkETECheque = False
End Function

Public Function CurVer() As Long
    CurVer = cVersion ' 23/08/2000
End Function

Public Function EUROText() As String
    EUROText = EUROText_
End Function

Public Function EUROText2002() As String
    EUROText2002 = EUROText2002_
End Function

Public Function GRDText() As String
    GRDText = GRDText_
End Function

Public Function EURORate() As Double
    EURORate = EURORate_
End Function

Public Function EUROAmount(inAmount) As String
    EUROAmount = EUROAmount_(inAmount)
End Function

Public Function EUROAmount2002(inAmount) As String
    EUROAmount2002 = EUROAmount2002_(inAmount)
End Function

Public Function GRDAmount(inAmount) As String
    GRDAmount = GRDAmount_(inAmount)
End Function

Public Function EUROAmount5(inAmount) As String
'Format ποσου για το πρόγραμμα 5
    EUROAmount5 = EUROAmount5_(inAmount)
End Function

Public Function GetDefUserProfileName() As String
    GetDefUserProfileName = cDefUserProfileName
End Function

'Public Function GetGroupUsersDoc() As MSXML2.DOMDocument
''Επιστρέφει XML Document με τους users του καταστήματος
'    Set GetGroupUsersDoc = GetGroupUsersDoc_
'End Function

Public Function TELLERName() As String
    TELLERName = "Τ :" & cUserName
End Function

Public Function CHIEFTELLERName() As String
    CHIEFTELLERName = IIf(Trim(cCHIEFUserName) <> "", "CT:" & cCHIEFUserName, "")
End Function

Public Function MANAGERName() As String
    MANAGERName = IIf(Trim(cMANAGERUserName) <> "", "M :" & cMANAGERUserName, "")
End Function

Public Function ChkChiefTeller() As Boolean
    ChkChiefTeller = isChiefTeller
End Function

Public Function ChkManager() As Boolean
    ChkManager = isManager
End Function

Public Function GetIRISAuth() As String
GetIRISAuth = ""
    KeyAccepted = False
    IRISSelKeyFrm.Show vbModal, Me
    
'    Load KeyWarning: Set KeyWarning.owner = Me: KeyWarning.Show vbModal, Me
    If Not KeyAccepted Then Exit Function
    GetIRISAuth = cIRISAuthUserName
End Function

Public Function GetChiefTellerKey() As Boolean
    Set SelKeyFrm.owner = Me: ChiefRequest = True
    SelKeyFrm.Show vbModal, Me
    If Not KeyAccepted Then GetChiefTellerKey = False: Exit Function
    GetChiefTellerKey = True ': isChiefTeller = True
    eJournalWrite "Εγκριση " & "Chief Teller απο:" & cCHIEFUserName
        SaveJournal

End Function

Public Function GetManagerKey() As Boolean
    Set SelKeyFrm.owner = Me: ManagerRequest = True
    SelKeyFrm.Show vbModal, Me
    If Not KeyAccepted Then GetManagerKey = False: Exit Function
    GetManagerKey = True ': isManager = True
    eJournalWrite "Εγκριση " & "Manager απο:" & cMANAGERUserName
        SaveJournal

End Function

Public Sub DisableSendFromJournal()
   HideSendFromJournal = True
End Sub

Public Sub EnableSendFromJournal()
   HideSendFromJournal = False
End Sub

Public Sub DisableReceiveFromJournal()
   HideReceiveFromJournal = True
End Sub

Public Sub EnableReceiveFromJournal()
   HideSendFromJournal = False
End Sub

Public Sub SetTRNCode(inPhase As Integer, inTrnCode As String)
    TrnCode(inPhase) = inTrnCode
End Sub

Public Sub ClearTotalEntries()
'καθαρίζει τη λίστα εγγραφών για τους αθροιστές
    aTotalEntries.ClearEntries_
End Sub

Public Sub AddDBTotalEntry(inTotalName, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    aTotalEntries.AddDBEntry_ inTotalName, inAmount
End Sub

Public Sub AddCRTotalEntry(inTotalName, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    aTotalEntries.AddCREntry_ inTotalName, inAmount
End Sub

Public Sub AddCurDBTotalEntry(inTotalName, inCurrency, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    aTotalEntries.AddCurDBEntry_ inTotalName, inCurrency, inAmount
End Sub

Public Sub AddCurCrTotalEntry(inTotalName, inCurrency, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    aTotalEntries.AddCurCrEntry_ inTotalName, inCurrency, inAmount
End Sub

Public Sub StoreTotalEntries()
'ενημερώνει τους αθροιστές από τη λίστα αθροιστών
    aTotalEntries.StoreEntries_
End Sub

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
    On Error GoTo errorpos1
'    DoEvents
    If cTRNCode = 0 Or cTRNCode = 610 Or cTRNCode = 620 _
    Or cTRNCode = 630 Or cTRNCode = 632 Then
        GenWorkForm.vStatus.Panels(1).Text = sMessage
    Else
        Status.Panels(1).Text = sMessage
    End If
    
    Status.Refresh
    Exit Sub
errorpos1:
    Call Runtime_error("Write Status Message", Err.number, Err.description)
End Sub

Public Function fnReadStatusMessage() As String
    On Error GoTo errorpos1
    If cTRNCode = 0 Then
        fnReadStatusMessage = GenWorkForm.vStatus.Panels(1).Text
    Else
        fnReadStatusMessage = Status.Panels(1).Text
    End If

    Exit Function
errorpos1:
    Call Runtime_error("Read Status Message", Err.number, Err.description)
End Function

Public Function PrepareSendOut() As Boolean
'1.Κάνει το Validation της οθόνης
'2.Εκτελεί το BeforeOutScript
'3.προετοιμασία buffer για αποστολή στο ΚΜ
'επιστρέφει TRUE αν η διαδικασία (μέ την καταγραφή στο ημερολόγιο)
'ολοκληρωθεί χωρίς πρόβλημα
    PrepareSendOut = PrepareSendOut_(processphase)
End Function

Public Function SendOut() As Boolean
'αποστολή στο ΚΜ
'Επιστρέφει TRUE αν η επικοινωνία ολοκληρωθεί χωρίς λάθος
    SendOut = SendOut_(processphase)
End Function

Public Sub SetOutBuffer(outString)
'Δημιουργεί το string που θα σταλεί από τα καθαρά data
    SetOutBuffer_ CStr(outString), processphase
End Sub

Public Sub ReadIn()
'1.Εκτελεί το AfterInScript
'2.μεταφέρει στα πεδία το περιεχόμενο του buffer που στάλθηκε από το ΚΜ
'3.και καταγράφει στο ημερολόγιο τα πεδία
'επιστρέφει TRUE αν η μεταφορά ολοκληρωθεί χωρίς πρόβλημα
    ReadIn_ processphase
End Sub

Public Sub ReadBuffer()
'έχει αντικατασταθεί από το ReadIn
'Διατηρείται μόνο για συμβατότητα με τις συναλλαγές που τυχόν έχει χρησημοποιειθεί
    ReadIn_ processphase
End Sub


Public Function StrPad(PString, PLength, Optional PChar, Optional PLftRgt) As String
' η πιό αξιόλογη ρουτίνα της εφαρμογής. Αν δεν ξέρεις τι κάνει πάτα Alt-F4
    StrPad = StrPad_(CStr(PString), CInt(PLength), PChar, PLftRgt)
End Function

Public Sub xClearDoc()
' καθαρισμός περιεχομένου παραστατικού
    xClearDoc_
End Sub

Public Sub xSetDocLine(inLineNo, inLineData)
' ανάθεση τιμής σε γραμμή παραστατικού
    xSetDocLine_ CInt(inLineNo), CStr(inLineData)
End Sub

Public Sub xSetInDocLine(inLineNo, inLineData, inX, inW, inAlign)
' ανάθεση τιμής σε περιοχή παραστατικού
' Align: "L" ή "R"
    xSetInDocLine_ CInt(inLineNo), CStr(inLineData), CInt(inX), CInt(inW), CStr(inAlign)
End Sub

Public Sub xPrintDoc(Optional inPrompt)
' εκτύπωση τρέχουσας μορφής παραστατικού
    If IsMissing(inPrompt) Then inPrompt = "Εκτύπωση Παραστατικού"
    xPrintDoc_ Me, inPrompt
End Sub

Public Function GetPassbookAmount(inAmount As Double) As String
' μορφοποίηση ποσού για εκτύπωση σε βιβλιάριο (συμπληρωμένο με *)
    GetPassbookAmount = GetPassbookAmount_(CDbl(inAmount))
End Function

Public Sub PrintPassbookLine(inAccount, inTrnDate, inTrnCode, inTrnAmount1, _
    inTrnAmount2, fromLine, fromAmount)
'inTrnAmount1: ποσο καταθεσης
'inTrnAmount2: ποσο αναληψη
'Dim apanel As SPCPanelX
'Set apanel = SPCPanel
Dim inTerm As String
inTerm = ""
'Εκτύπωση μιάς μόνο γραμμής στο βιβλιάριο για κατάθεση ανάληψη ή ενημέρωση
    PrintSinglePassbookLine_ Me, CStr(inAccount), CStr(inTrnDate), CInt(inTrnCode), CDbl(inTrnAmount1), _
        CDbl(inTrnAmount2), CInt(fromLine), CDbl(fromAmount), CStr(inTerm)
End Sub

Public Sub PrintPassbook(inAccount As String, inTrnType As Integer, inTrnCode As String, inTrnAmount As Double, _
    fromLine As Integer, fromAmount As Double, Optional PrintEUROText As Integer)
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

' εκτύπωση βιβλιαρίου πρόγραμμα 1
    If (cVersion < 20010101) Or IsMissing(PrintEUROText) Then
        PrintPassbook_ Me, CStr(inAccount), CInt(inTrnType), CStr(inTrnCode), CDbl(inTrnAmount), CInt(fromLine), CDbl(fromAmount)
    Else
        PrintPassbook_ Me, CStr(inAccount), CInt(inTrnType), CStr(inTrnCode), CDbl(inTrnAmount), CInt(fromLine), CDbl(fromAmount), PrintEUROText
    End If
End Sub

Public Sub PrintPassbook5(inAccount As String, inTrnType As Integer, _
    inTrnCode As String, inTrnAmount As Double, inTrnDRXAmount As Double, fromLine As Integer, fromAmount As Double, Optional inTrnEuroFinalAmount)
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

'εκτύπωση βιβλιαρίου πρόγραμμα 5
    If (cVersion < 20010101) Or IsMissing(inTrnEuroFinalAmount) Then
        PrintPassbook5_ Me, CStr(inAccount), CInt(inTrnType), CStr(inTrnCode), CDbl(inTrnAmount), CDbl(inTrnDRXAmount), CInt(fromLine), CDbl(fromAmount)
    Else
        PrintPassbook5_ Me, CStr(inAccount), CInt(inTrnType), CStr(inTrnCode), CDbl(inTrnAmount), CDbl(inTrnDRXAmount), CInt(fromLine), CDbl(fromAmount), CDbl(inTrnEuroFinalAmount)
    End If
End Sub

Public Function AddCheque_0(inPostDate, inGroup, inAccount, inCheque, inAmount, inChequeDate) As Integer
' εγγραφή στο αρχείο επιταγών για επιταγή ΕΤΕ
' επιστρέφει 0 αν αποτύχει ή τον αριθμό γραμμής διαφορετικά
    AddCheque_0 = 0
End Function

Public Function AddCheque_1(inPostDate, inGroup, inBank, inbranch, inAccount, inCheque, inAmount, inChequeDate) As Integer
' εγγραφή στο αρχείο επιταγών για επιταγή ξένης τράπεζας
' επιστρέφει 0 αν αποτύχει ή τον αριθμό γραμμής διαφορετικά
    AddCheque_1 = 0
End Function

Public Sub AddCheque_2(inPostDate, inGroup, inDelGroup, inLine)
End Sub

Public Function AddCheque_4(inPostDate, inGroup, inBank, inbranch, inAccount, inCheque, inAmount, inChequeDate) As Integer
' εγγραφή στο αρχείο επιταγών για επιταγή ξένης τράπεζας
' επιστρέφει 0 αν αποτύχει ή τον αριθμό γραμμής διαφορετικά
    AddCheque_4 = 0
End Function

Public Function AddCheque_5(inPostDate, inGroup, inAccount, inCheque, inAmount, inChequeDate) As Integer
' εγγραφή στο αρχείο επιταγών για επιταγή ΕΤΕ
' επιστρέφει 0 αν αποτύχει ή τον αριθμό γραμμής διαφορετικά
    AddCheque_5 = 0
End Function

Public Function ChkBankAcount(inBank, inbranch, inAcc, Optional inChequeType) As Boolean
'έλεγχος check digit λογαριασμού τράπεζας
    ChkBankAcount = ChkBankAccount_(CStr(inBank), CStr(inbranch), CStr(inAcc), CInt(inChequeType))
End Function

Public Function ChkBankCheque(inBank, inbranch, inAcc, inCheque, Optional inChequeType) As Boolean
'έλεγχος check digit επιταγής τράπεζας
    ChkBankCheque = ChkBankCheque_(CStr(inBank), CStr(inbranch), CStr(inAcc), CStr(inCheque), CInt(inChequeType))
End Function

Public Sub SendCheques()
'αποστολή των επιταγών
End Sub
    

Public Sub SetInBuffer(inString As String)
' μεταβολή του buffer λήψης στοιχείων από ΚΜ
    cb.received_data = inString
End Sub

Public Function GetWorkDir() As String
' επιστρέφει το NetWork directory
    GetWorkDir = WorkDir
End Function

Public Function GetReadDir() As String
' επιστρέφει το VbRead directory
    GetReadDir = ReadDir
End Function

Public Function GetPostDate_U6() As String
' ημερομηνία λειτουργίας σε DDMMYY μορφή
    GetPostDate_U6 = format(cPOSTDATE, "DDMMYY")
End Function

Public Function GetPostDate_U8() As String
' ημερομηνία λειτουργίας σε DDMMYYYY μορφή
    GetPostDate_U8 = format(cPOSTDATE, "DDMMYYYY")
End Function

Public Function GetPostDate_F8() As String
' ημερομηνία λειτουργίας σε DD/MM/YY μορφή
    GetPostDate_F8 = format(cPOSTDATE, "DD/MM/YY")
End Function

Public Function GetPostDate_F10() As String
' ημερομηνία λειτουργίας σε DD/MM/YYYY μορφή
    GetPostDate_F10 = format(cPOSTDATE, "DD/MM/YYYY")
End Function

Public Function GetFullDatabaseName() As String
' Connection string για τη βάση που συντηρεί ημερολόγιο, αθροιστές, batch, επιταγές, στοιχεία τρίτων
    GetFullDatabaseName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & WorkDir & "VBWork97.mdb"
End Function

Public Function GetTerminalID() As String
' ταυτότητα τερματικού
    GetTerminalID = cTERMINALID
End Function

Public Function GetTrnNum() As String
' αριθμός τρέχουσας συναλλαγής
    GetTrnNum = StrPad_(CStr(cTRNNum), 3, "0", "L")
End Function

Public Function GetBranchCode() As String
' κωδικός καταστήματος
    GetBranchCode = StrPad_(CStr(IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH)), 3, "0", "L")
End Function

Public Function GetBranchName() As String
' ονομασία καταστήματος
    GetBranchName = cBRANCHName
End Function

Public Function GetAmountText2002(inAmount) As String
' ποσό ολογράφως
    GetAmountText2002 = Amount_Str2002(CDbl(inAmount))
End Function
Public Function GetAmountTextGen(inAmount, Optional pFlag1 As Variant, Optional pFlag2 As Variant)
  'ποσό ολογράφως
  'pFlag2 αν τυπώνει ΕΥΡΩ Η ΔΡΧ
  'pFlag1 ΘΗΛ Η ΟΥΔ. ΝΟΜΙΣΜ.
  If IsMissing(pFlag1) Then pFlag1 = False
  If IsMissing(pFlag2) Then pFlag2 = False
  GetAmountTextGen = Amount_Str2002(CDbl(inAmount), pFlag1, pFlag2)
End Function

Public Function GetAmountText(inAmount) As String
' ποσό ολογράφως
    GetAmountText = Amount_str(CDbl(inAmount), True)
End Function

Public Function GetChequeAmountText(inAmount) As String
' ποσό ολογράφως ειδικά για επιταγές συναλλαγή 6000
    GetChequeAmountText = Cheque_Amount_str_(CDbl(inAmount), False)
End Function

Public Function UpdateIRISTotals() As Integer
' σύνολο αθροιστή IRIS
    UpdateIRISTotals = UpdateIRISTotals_(Me)
End Function

Public Function ClearIRISTotals() As Integer
' σύνολο αθροιστή IRIS
    ClearIRISTotals = ClearIRISTotals_(Me)
End Function

Public Function GetIRISTotal(TotalName) As Double
' σύνολο αθροιστή IRIS
    GetIRISTotal = GetIRISTotal_(CStr(TotalName))
End Function

Public Function GetBranchIRISTotal(TotalName) As Double
' σύνολο αθροιστή IRIS στο κατάστημα
    GetBranchIRISTotal = GetBranchIRISTotal_(CStr(TotalName))
End Function

Public Function GetTotal(TotalName) As Double
' σύνολο αθροιστή
    GetTotal = GetTotal_(CStr(TotalName))
End Function

Public Function GetBranchTotal(TotalName) As Double
' σύνολο αθροιστή για το κατάστημα
    GetBranchTotal = GetBranchTotal_(CStr(TotalName))
End Function

Public Function GetCurTotal(TotalName, Cur) As Double
' σύνολο αθροιστή σε νόμισμα
    GetCurTotal = GetCurTotal_(CStr(TotalName), CInt(Cur))
End Function

Public Function GetDBTotal(TotalName, Optional term, Optional pDate) As Double
' σύνολο χρέωσης αθροιστή
    If IsMissing(term) Then term = ""
    If IsMissing(pDate) Then pDate = cPOSTDATE
    If term = "" Then
        GetDBTotal = GetDBTotal_(CStr(TotalName))
    Else
        GetDBTotal = GetDBTotalTerm_(CStr(TotalName), CStr(term), CDate(pDate))
    End If
End Function

Public Function GetBranchDBTotal(TotalName) As Double
' σύνολο χρέωσης αθροιστή για το κατάστημα
    GetBranchDBTotal = GetBranchDBTotal_(CStr(TotalName))
End Function

Public Function GetCRTotal(TotalName, Optional term, Optional pDate) As Double
' σύνολο πίστωσης αθροιστή
    If IsMissing(term) Then term = ""
    If IsMissing(pDate) Then pDate = cPOSTDATE
    If term = "" Then
        GetCRTotal = GetCRTotal_(CStr(TotalName))
    Else
        GetCRTotal = GetCRTotalTerm_(CStr(TotalName), CStr(term), CDate(pDate))
    End If
End Function

Public Function GetBranchCRTotal(TotalName) As Double
' σύνολο πίστωσης αθροιστή για το κατάστημα
    GetBranchCRTotal = GetBranchCRTotal_(CStr(TotalName))
End Function

Public Sub SetDBTotal(TotalName, aValue)
' ανάθεση στο συνόλο χρέωσης αθροιστή
End Sub

Public Sub SetCRTotal(TotalName, aValue)
' ανάθεση στο συνόλο πίστωσης αθροιστή
End Sub

Public Sub AddDBTotal(TotalName, aValue)
' προσθέτει στο σύνολο χρέωσης αθροιστή
End Sub

Public Sub AddCRTotal(TotalName, aValue)
' προσθέτει στο σύνολο πίστωσης αθροιστή
End Sub

Public Function GetCurDBTotal(TotalName, Cur, Optional term) As Double
' σύνολο χρέωσης αθροιστή σε νόμισμα
    If IsMissing(term) Then term = ""
    GetCurDBTotal = GetCurDBTotal_(CStr(TotalName), CInt(Cur), CStr(term))
End Function

Public Function GetCurCRTotal(TotalName, Cur, Optional term) As Double
    If IsMissing(term) Then term = ""
    GetCurCRTotal = GetCurCRTotal_(CStr(TotalName), CInt(Cur), CStr(term))
End Function

Public Sub SetCurDBTotal(TotalName, Cur, aValue)
' ανάθεση στο συνόλο χρέωσης αθροιστή σε νόμισμα
End Sub

Public Sub SetCurCRTotal(TotalName, Cur, aValue)
' ανάθεση στο συνόλο πίστωσης αθροιστή σε νόμισμα
End Sub

Public Sub AddCurDBTotal(TotalName, Cur, aValue)
' πρόσθεση στο συνόλο χρέωσης αθροιστή σε νόμισμα
End Sub

Public Sub AddCurCRTotal(TotalName, Cur, aValue)
' πρόσθεση στο συνόλο πίστωσης αθροιστή σε νόμισμα
End Sub

Public Function GetNextCur(TotalName, inCur, Optional inTerm) As Integer
' επιστρέφει το επόμενο νόμισμα με υπόλοιπο για συγκεκριμένο αθροιστή
' η πρώτη κλήση γίνεται με νόμισμα 0
' αν δεν βρεθεί επόμενο επιστρέφει -1
    If IsMissing(inTerm) Then inTerm = ""
    GetNextCur = GetNextCur_(CStr(TotalName), CInt(inCur), CStr(inTerm))
End Function

'Public Function GetNextCur(TotalName, inCur) As Integer
'' επιστρέφει το επόμενο νόμισμα με υπόλοιπο για συγκεκριμένο αθροιστή
'' η πρώτη κλήση γίνεται με νόμισμα 0
'' αν δεν βρεθεί επόμενο επιστρέφει -1
'    GetNextCur = GetNextCur_(CStr(TotalName), CInt(inCur))
'End Function
'
Public Sub GetAllTotals()
' ενημέρωση του πίνακα αθροιστών
    Call fnDisplayTotals(GenWorkForm.TotalsGrid)
End Sub

Public Sub ClearTotals()
' μηδενισμός αθροιστών
    Call fnDisplayTotals(GenWorkForm.TotalsGrid)
End Sub

Public Sub ClearBranchTotals()
' μηδενισμός αθροιστών καταστήματος
End Sub

Public Sub ClearTotal(inTotal)
' μηδενισμός αθροιστή δραχμών
    Call fnDisplayTotals(GenWorkForm.TotalsGrid)
End Sub

Public Sub CopyTotalsToBranch()
' αντιγράφει τα σύνολα του τερματικού στο αρχείο συνόλων του καταστήματος
End Sub

Public Sub ClearAllTotal(TotalName)
' πλήρης μηδενισμός αθροιστή νομίσματος (σε όλα τα νομίσματα)
    Call fnDisplayTotals(GenWorkForm.TotalsGrid)
End Sub

Public Sub ClearCurTotal(inTotal, inCurrency)
' μηδενισμός αθροιστή νομίσματος σε συγκεκριμένο νόμισμα
    Call fnDisplayTotals(GenWorkForm.TotalsGrid)
End Sub

Public Sub WriteJournal(amessage)
' εγγραφή μυνήματος στο αρχείο ημερολογίου
    eJournalWrite CStr(amessage)
    SaveJournal
End Sub

Public Sub WriteJournalFinal()
' εγγραφή ένδειξης τέλους συναλλαγής στο αρχείο ημερολογίου
    eJournalWriteFinal Me
End Sub

Public Sub SetKey(aKey)
' αλλαγή του απαιτούμενου από τη συναλλαγή κλειδιού
    SpecialKey = CStr(aKey)
    trn_key = CStr(aKey)
End Sub

Public Function Read3rdList() As Boolean
' ενημέρωση του αρχείου στοιχείων τρίτων από το buffer λήψης στοιχείων από ΚΜ
    Read3rdList = False
End Function

Public Function GetBatchList() As Boolean
' επιστρέφει στο ListData collection τα περιεχόμενα του batch
    GetBatchList = GetBatchList_(Me)
End Function

Public Function ClearBatchList(lastline As Long) As Boolean
' διαγράφει τις εγγραφές του batch με sn <= LastLine
    ClearBatchList = ClearBatchList_(CLng(lastline))
End Function
Public Function DisplayTotalsTrace(indate As Date, inTerminalID As String, inTrnCode As String, _
                                    inTotalFrom As Double, inTotalTo As Double, _
                                    inTotal As String, inTotalsGrid As Object) As Boolean
    DisplayTotalsTrace = False
End Function

Public Function UpdateBonusStats(inPostDate) As Boolean
' ενημερώνει τον πίνακα bonus για τους teller
    UpdateBonusStats = False
End Function

Public Function MonthBonusStats(fromPostDate, toPostDate) As Boolean
' επιστρέφει σύνολα χρόνου και βαθμών για ta bonus του μήνα
    MonthBonusStats = False
End Function
Public Function EnableAllBranchTrn(EnableFlag As Boolean, GroupName As String) As Boolean
'ενεργοποιεί όλες τις συναλλαγές για το κατάστημα
    EnableAllBranchTrn = False
End Function

Public Function EnableBranchTrn(TrnCD) As Boolean
'ενεργοποιεί τη συναλλαγή για το κατάστημα
    EnableBranchTrn = False
End Function

Public Function EnableGroupTrn(TrnCD, GroupName) As Boolean
'ενεργοποιεί τη συναλλαγή για group χρηστών
    EnableGroupTrn = False
End Function

Public Function EnableUserTrn(TrnCD, UserName) As Boolean
'ενεργοποιεί τη συναλλαγή για χρηστη
    EnableUserTrn = False
End Function

Public Function DisableBranchTrn(TrnCD) As Boolean
'απενεργοποιεί τη συναλλαγή για το κατάστημα
    DisableBranchTrn = False
End Function

Public Function DisableGroupTrn(TrnCD, GroupName) As Boolean
'απενεργοποιεί τη συναλλαγή για group χρηστών
    DisableGroupTrn = False
End Function

Public Function DisableUserTrn(TrnCD, UserName) As Boolean
'απενεργοποιεί τη συναλλαγή για χρηστη
    DisableUserTrn = False
End Function

Public Function EnableTerminalForBonus(inTerminalID) As Boolean
'προσθήκη τερματικού στο σύστημα bonus
    EnableTerminalForBonus = False
End Function

Public Function SetUserProfile(inUName, inProfile) As Boolean
    SetUserProfile = False
End Function

Public Function DisableTerminalForBonus(inTerminalID) As Boolean
'διαγραφή τερματικού από το σύστημα bonus
    DisableTerminalForBonus = False
End Function

Public Function WriteErrorMessage(inMessage) As Boolean
'Τυπώνει στο ημερολόγιο και στη γραμμή μηνυμάτων κάποιο μήνυμα
    sbWriteStatusMessage inMessage
    eJournalWriteAll Me, CStr(inMessage) ', CStr(cTRNCode), cTRNNum
    WriteErrorMessage = True
    SaveJournal
End Function

Public Sub BackupWorkDB()
'κανει backup τη βάση του ημερολογίου στο WorkDir (Network)
End Sub

Public Function GetADOConnection() As ADODB.Connection

End Function

Public Function GetADORecordset(inCmd, Optional inCursorType, Optional inLockType) As ADODB.Recordset

End Function

Public Function ExecADOCommand(aCommandStr) As Integer

End Function

Public Function ExecTradeStoredProcedure(aStoredProc) As ADODB.Recordset
    Dim aRecNo As Long
    On Error GoTo TradeError
    Set ExecTradeStoredProcedure = trade_db.Execute(CStr(aStoredProc), aRecNo, adCmdText)
    Exit Function
TradeError:
       Set ExecTradeStoredProcedure = Nothing
       Exit Function
End Function

Public Function GetTradeRecordSet(ByVal cmdStr As String) As ADODB.Recordset
       Dim adoRecs As New ADODB.Recordset
       On Error GoTo TradeError
       adoRecs.open cmdStr, trade_db, adOpenKeyset, adLockOptimistic
'       If Not (adoRecs.BOF Or adoRecs.EOF) Then
          Set GetTradeRecordSet = adoRecs: Exit Function
'       End If
TradeError:
       Set GetTradeRecordSet = Nothing
       Exit Function
End Function
'Public Function CalcTradeComm(inComm, inAmount) As Double
'    CalcTradeComm = CalcTradeComm_(CInt(inComm), CDbl(inAmount))
'End Function
'Public Function CalcTradeCommDiscount(inCs, inComm, inAmount, inAmountCalc) As Double
'    CalcTradeCommDiscount = CalcTradeCommDiscount_(CDbl(inCs), CInt(inComm), CDbl(inAmount), CDbl(inAmountCalc))
'End Function
'Public Function CheckTradeAccRange(inGen, inSpc) As Boolean
'    CheckTradeAccRange = CheckTradeAccRange_(CInt(inGen), CDbl(inSpc))
'End Function
'pa

Public Function AppRS_S() As Collection
    Set AppRS_S = GenWorkForm.AppRS_S
End Function

Public Function AddAppRS_S(inName, Optional inConnection) As XMLRecordsetView
    Dim aview As XMLRecordsetView
    
    Set aview = New XMLRecordsetView
    aview.name = inName
    On Error GoTo invalidRS
    If IsMissing(inConnection) Then
        aview.prepare Nothing, Nothing
    ElseIf UCase(inConnection) = "VBTRADE" Then
        aview.prepare VBTradeSLink, Nothing
    End If
    
    GenWorkForm.AppRS_S.add aview, inName
    Set AddAppRS_S = aview
    Exit Function
invalidRS:
    Dim astr As String, aMsg As Integer
    astr = Err.description: aMsg = Err.number
    Set AddAppRS_S = Nothing
    MsgBox "ΛΑΘΟΣ (" & aMsg & "). " & astr, vbCritical, "On Line Εφαρμογή"
    
End Function

Public Function AddAppRecordset(inName, inCmd, Optional inConnection, Optional inCursorType, Optional inLockType) As ADODB.Recordset
    Set AddAppRecordset = Nothing
End Function

Public Function AddFTFilaRecordset(inName, inFilter, Optional inSort) As ADODB.Recordset
    Set AddFTFilaRecordset = AddFTFilaRecordset_(inName, inFilter, inSort)
End Function

Public Function AddAppStoredProcedure(inName, inCmd, Optional inConnection) As ADODB.command
    Set AddAppStoredProcedure = Nothing
End Function

Public Function AppRecordsetByName(inName) As ADODB.Recordset
    Set AppRecordsetByName = AppRecordsetByName_(inName)
End Function

Public Function AppRSEntryByName(inName) As RecordsetEntry
    Set AppRSEntryByName = AppRSEntryByName_(inName)
End Function

Public Function AppRecordsetByIndex(inIdx) As ADODB.Recordset
    Set AppRecordsetByIndex = AppRecordsetByIndex_(inIdx)
End Function

Public Sub FreeAppRecordset(inName)
    FreeAppRecordset_ inName
End Sub

Public Sub FreeAppCRecordset(inName)
    FreeAppCRecordset_ inName
End Sub

Public Function AppStoredProcedureByName(inName) As ADODB.command
    Set AppStoredProcedureByName = Nothing
End Function

Public Function AppStoredProcedureByIndex(inIdx) As ADODB.command
    Set AppStoredProcedureByIndex = Nothing
End Function

Public Sub FreeAppStoredProcedure(inName)
End Sub

'Public Function LastOutMessage() As String
'    LastOutMessage = cb.initsend_str
'End Function

Public Function RemoveChar(inString, inchar) As String
Dim i As Integer, bstr As String
    For i = 1 To Len(CStr(inString))
        If Mid(CStr(inString), i, 1) <> CStr(inchar) Then bstr = bstr & Mid(CStr(inString), i, 1)
    Next i
    RemoveChar = bstr
End Function

Public Sub SEND()
'Public Sub SEND(Sender As Control)
Dim res As Boolean
    If SendStarted Then Exit Sub
    If Not (ActiveTextBox Is Nothing) Then res = ControlLostFocus(ActiveTextBox) Else res = True
    If res Then
        SendStarted = True: NextAction = taSend_Buffer: ProcessLoop: SendStarted = False
    End If
End Sub

Public Function WriteTotals() As Boolean
If Trim(TotalName) <> "" Then
    If TotalPos = 1 Then AddDBTotal_ TotalName, 100 _
    Else If TotalPos = 2 Then AddCRTotal_ TotalName, 100
End If
Dim i As Integer
For i = 1 To fields.count
    If fields(i).TotalName <> "" Then
        If fields(i).TotalPos = 1 Then AddDBTotal_ fields(i).TotalName, fields(i).AsDouble _
        Else If fields(i).TotalPos = 2 Then AddCRTotal_ fields(i).TotalName, fields(i).AsDouble
    End If
Next i
WriteTotals = True
End Function

Public Function ControlLostFocus(sender As GenTextBox) As Boolean
On Error GoTo BypassError
Dim i As Integer, vcount As Integer
Dim astr As String, bstr As String
    
    If Not (LastChkControl Is Nothing) Then
        If (Not LastValidChk) And Not (LastChkControl Is sender) Then ControlLostFocus = True: Exit Function
    End If
    
    If ExitPhaseStarted _
    Or CurrAction = taExit_Form Or CurrAction = taStay_In_Form _
    Or DisableControlLostFocus Then
        ControlLostFocus = True
        Exit Function
    End If
    
    Set LastChkControl = sender
    
    
    astr = sender.Text
    LastValidChk = sender.ChkValidL1(processphase)
    If Not LastValidChk Then
        sender.SetFocus: sender.ControlText = astr: Beep
    Else
        ChangeFocusOk = True: ChangeFocusError = ""
        If LostFocusFlag Then ValidationControl.Run "LostFocus_Script", sender
        If Not ChangeFocusOk Then
            LastValidChk = False
            Beep
            If ChangeFocusError = "" Then
                
            Else
                sbWriteStatusMessage ChangeFocusError
            End If
            sender.SetFocus: sender.ControlText = astr
        Else
'            sbWriteStatusMessage ""
        End If
    End If

    ControlLostFocus = LastValidChk
BypassError:
'do nothing
End Function

Public Function GetFormatedFld(inFldName As String) As String
On Error GoTo ErrorValue
    GetFormatedFld = fields(inFldName).FormatedText
    Exit Function
ErrorValue:
    GetFormatedFld = "Error!!!"
End Function

Public Sub SetDefaultFocus(inPhase As Integer)
'Επιλέγει το πρώτο active πεδίο της οθόνης
Dim i As Integer, foundflag As Boolean, selectedOrder As Long
    selectedOrder = 999999999
    Dim astr As String
    astr = fnReadStatusMessage
    foundflag = False
    For i = 1 To fields.count
        If fields(i).IsVisible(inPhase) And fields(i).IsEditable(inPhase) Then
            If fields(i).TTabIndex < selectedOrder Then
                fields(i).SetAsActive: 'Exit Sub:
                selectedOrder = fields(i).TTabIndex
            End If
        End If
    Next i
    sbWriteStatusMessage astr
    If Not foundflag Then Exit Sub
    
    For i = 1 To fields.count
        If fields(i).Visible Then fields(i).SetAsActive: Exit For
    Next i
    sbWriteStatusMessage astr
End Sub

Private Sub MoveItem(item, Optional NoPrompt As Boolean)
    If item.IsVisible(processphase) Then
        item.Visible = True
        ScaleMode = vbCharacters: item.Left = item.ScrLeft: item.width = item.ScrWidth
        ScaleMode = vbTwips: item.Top = item.ScrTop: item.height = item.ScrHeight
        
        If TypeOf item Is shine.GenBrowser Then
        ElseIf TypeOf item Is shine.GenRichTextBox Then
        ElseIf TypeOf item Is shine.GenLabel Then
        Else
            If IsMissing(NoPrompt) Then
                If Not (item.Prompt Is Nothing) Then item.Prompt.Visible = True
            Else
                If Not NoPrompt Then
                    On Error Resume Next
                    If Not (item.Prompt Is Nothing) Then item.Prompt.Visible = True
                End If
            End If
        End If
    Else
        item.Visible = False
        If IsMissing(NoPrompt) Then
            If Not (item.Prompt Is Nothing) Then item.Prompt.Visible = False
        Else
            If Not NoPrompt Then
                If Not (item.Prompt Is Nothing) Then item.Prompt.Visible = False
            End If
        End If
    End If
End Sub

Public Sub RefreshView()
Dim i As Long
On Error GoTo 0
    EditFieldsCount = 0
    For i = 1 To fields.count
        MoveItem fields(i)
        If fields(i).IsVisible(processphase) Then
            If CurrAction <> taStay_In_Form Then
                fields(i).HandleEdit (processphase)
                fields(i).TabStop = fields(i).IsEditable(processphase) Or fields(i).Tabbed
                If fields(i).IsEditable(processphase) Then EditFieldsCount = EditFieldsCount + 1
            Else
                fields(i).SetEditableNoRefresh processphase, False: fields(i).TabStop = False Or fields(i).Tabbed
            End If
        End If
    Next i
    For i = 1 To Combos.count
        MoveItem Combos(i)
        If Combos(i).IsVisible(processphase) Then
            If CurrAction <> taStay_In_Form Then
                Combos(i).HandleEdit (processphase)
                Combos(i).TabStop = Combos(i).IsEditable(processphase) 'Or Combos(i).Tabbed
                If Combos(i).IsEditable(processphase) Then EditFieldsCount = EditFieldsCount + 1
            Else
                Combos(i).SetEditableNoRefresh processphase, False: Combos(i).TabStop = False 'Or Combos(i).Tabbed
            End If
        End If
    Next i
    For i = 1 To Checks.count
        MoveItem Checks(i), True
        If Checks(i).IsVisible(processphase) Then
            If CurrAction <> taStay_In_Form Then
                Checks(i).HandleEdit (processphase)
                Checks(i).TabStop = Checks(i).IsEditable(processphase)
                If Checks(i).IsEditable(processphase) Then EditFieldsCount = EditFieldsCount + 1
            Else
                Checks(i).SetEditableNoRefresh processphase, False: Checks(i).TabStop = False
            End If
        End If
    Next i
    
    For i = 1 To Buttons.count
        MoveItem Buttons(i), True
        Buttons(i).TabStop = False
    Next i
    
    For i = 1 To Charts.count
        MoveItem fields(i)
    Next i
    
    For i = 1 To Browsers.count
        MoveItem Browsers(i)
    Next i
    
    
    For i = 1 To RichTextBoxes.count
        MoveItem RichTextBoxes(i)
    Next i
    
    For i = 1 To Labels.count
        MoveItem Labels(i), True
    Next i
    
    For i = 1 To Lists.count
        MoveItem Lists(i)
        Lists(i).TabStop = False
    Next i
    
    For i = 1 To Spreads.count
        MoveItem Spreads(i)
    Next i
    
End Sub

Private Sub ShowFields(inPhase As Integer, Optional aSetDefaultFocus)
Dim i As Integer, SetDefaultFocusFlag As Boolean

If IsMissing(aSetDefaultFocus) Then SetDefaultFocusFlag = False Else SetDefaultFocusFlag = aSetDefaultFocus
LockWindowUpdate (Me.hwnd)

Dim BackStatus As String
BackStatus = fnReadStatusMessage
On Error GoTo 0
    RefreshView
    
'    EditFieldsCount = 0
'    For i = 1 To Fields.count
'        MoveItem Fields(i)
'        If Fields(i).IsVisible(inPhase) Then
'            If CurrAction <> taStay_In_Form Then
'                Fields(i).HandleEdit (inPhase)
'                Fields(i).TabStop = Fields(i).IsEditable(inPhase) Or Fields(i).Tabbed
'                If Fields(i).IsEditable(inPhase) Then EditFieldsCount = EditFieldsCount + 1
'            Else
'                Fields(i).SetEditableNoRefresh inPhase, False: Fields(i).TabStop = False Or Fields(i).Tabbed
'            End If
'        End If
'    Next i
'
'    For i = 1 To Checks.count
'        MoveItem Checks(i), True
'        If Checks(i).IsVisible(inPhase) Then
'            If CurrAction <> taStay_In_Form Then
'                Checks(i).HandleEdit (inPhase)
'                Checks(i).TabStop = Checks(i).IsEditable(inPhase)
'                If Checks(i).IsEditable(inPhase) Then EditFieldsCount = EditFieldsCount + 1
'            Else
'                Checks(i).SetAsReadOnly: Checks(i).TabStop = False
'            End If
'        End If
'    Next i
'
'    For i = 1 To Combos.count
'        MoveItem Combos(i)
'        If Combos(i).IsVisible(inPhase) Then
'            If CurrAction <> taStay_In_Form Then
'                Combos(i).HandleEdit (inPhase)
'                Combos(i).TabStop = Combos(i).IsEditable(inPhase)
'                If Combos(i).IsEditable(inPhase) Then EditFieldsCount = EditFieldsCount + 1
'            Else
'                Combos(i).SetEditable inPhase, False: Combos(i).TabStop = False
'            End If
'        End If
'    Next i
'
'    For i = 1 To Buttons.count
'        MoveItem Buttons(i), True
'    Next i
'
'    For i = 1 To Charts.count
'        MoveItem Charts(i)
'    Next i
'
'    For i = 1 To Browsers.count
'        MoveItem Browsers(i), True
'    Next i
'
'    For i = 1 To Labels.count
'        MoveItem Labels(i), True
'    Next i
    
    If SetDefaultFocusFlag Then SetDefaultFocus inPhase
 
Dim listfocused As Boolean
listfocused = False
    For i = 1 To Lists.count
        'MoveItem Lists(i)
        If Lists(i).IsVisible Then
            If Not listfocused And CurrAction = taStay_In_Form Then _
                Lists(i).TabStop = True: listfocused = True
        End If
        'Lists(i).TabStop = False
    Next i
    For i = 1 To Spreads.count
        'MoveItem Spreads(i)
        If Spreads(i).IsVisible Then
            If Not listfocused And CurrAction = taStay_In_Form Then _
                Spreads(i).TabStop = True: Spreads(i).SetFocus: listfocused = True
        End If
    Next i
    
    sbWriteStatusMessage BackStatus
    LockWindowUpdate (0&)
End Sub

Private Sub InitForm()
Dim i As Integer, aStatus As Boolean, Line As Integer
On Error GoTo ExitPoint
    Line = 100
    AppBuffersPos = GenWorkForm.AppBuffers.BufferNum
    AppRSPos = GenWorkForm.AppRS.count
    AppVariablesPos = GenWorkForm.AppVariables.count
    AppSPPos = GenWorkForm.AppSP.count
    AppRS_SPos = GenWorkForm.AppRS_S.count
    AppCRSPos = GenWorkForm.AppCRS.count

    Line = 200
    CloseTransactionFlag = True
    If cTRNCode = "9857" Then aStatus = True Else aStatus = ChkProfileAccessNew(cTRNCode)
    
    If Not aStatus Then
        CloseTransactionFlag = True:
        MsgBox "Δεν Επιτρέπεται η χρήση της συναλλαγής: " & (cTRNCode), vbCritical
        Exit Sub
    End If
        
    Line = 300
    Set TrnBuffers = New Buffers: TrnBuffers.name = cTRNCode & "_AppBuffers"
    
    Set aTotalEntries = New TotalEntries
    CurrAction = taNo_Action
    DisableControlLostFocus = False
    
    CloseTransactionFlag = False
    SpecialKey = ""
    SelectedTRN = cTRNCode
    cCHIEFUserName = ""
    cMANAGERUserName = ""
    
    Status.Panels(1).width = Round(Status.width * 0.9)
   
    Left = GenWorkForm.Left
    Top = GenWorkForm.Top
    width = GenWorkForm.width
    height = GenWorkForm.height
    ScaleMode = vbCharacters
   
    ExitPhaseStarted = False
    StartupCodeFlag = False
    FormValidationFlag = False
    BeforeOutFlag = False
    AfterInFlag = False
    BeforeActionFlag = False
    AfterActionFlag = False
    AfterKeyFlag = False
    CommunicationErrorFlag = False
    DisableEnterKey = False
    CancelCommunicationFlag = False
    CancelPrintFlag = False
    BeforePrintFlag = False
    BeforeDocumentFlag = False
    AfterDocumentFlag = False
    HideSendFromJournal = False
    HideReceiveFromJournal = False
    SkipKeyChk = False
    SkipCommConfirmation = False
    
    PrintPromptMessage = "Εισαγωγή Παραστατικού"
    
    Line = 400
    'Set ActiveTextBox = Nothing
    'If IsEmpty(ActiveTextBox) <> Null Then
    '    Set ActiveListBox = Nothing
    'End If
    'Set ActiveSpread = Nothing
    
    Set ListData = ReceivedData
    Set ListG0 = ListG0
       
    ValidationControl.TimeOut = -1
    ValidationControl.AddCode "Const No_Action = 0"
    ValidationControl.AddCode "Const Get_Input = 200"
    ValidationControl.AddCode "Const Send_Buffer = 201"
    ValidationControl.AddCode "Const Print_Document = 202"
    ValidationControl.AddCode "Const Exit_Form = 203"
    ValidationControl.AddCode "Const Escape_Form = 204"
    ValidationControl.AddCode "Const Stay_In_Form = 205"
    
    ValidationControl.AddCode "Const fopNoOperation = 0"
    ValidationControl.AddCode "Const fopExitForm = 1"
    ValidationControl.AddCode "Const fopMoveNextPhase = 2"
    ValidationControl.AddCode "Const fopSendBuffer = 3"
    ValidationControl.AddCode "Const fopCloseForm = 4"
    
    
    ValidationControl.AddCode "Const Teller_Key = """ & cTELLERKEY & """"
    ValidationControl.AddCode "Const Chief_Key = """ & cCHIEFKEY & """"
    ValidationControl.AddCode "Const Manager_Key = """ & cMANAGERKEY & """"
    ValidationControl.AddCode "Const TellerChief_Key = """ & cTELLERCHIEFKEY & """"
    ValidationControl.AddCode "Const TellerManager_Key = """ & cTELLERMANAGERKEY & """"
    
    
'ADO Constants
    ValidationControl.AddCode "Const adOpenDynamic = " & CStr(adOpenDynamic)
    ValidationControl.AddCode "Const adOpenForwardOnly = " & CStr(adOpenForwardOnly)
    ValidationControl.AddCode "Const adOpenKeyset = " & CStr(adOpenKeyset)
    ValidationControl.AddCode "Const adOpenStatic = " & CStr(adOpenStatic)
    ValidationControl.AddCode "Const adOpenUnspecified = " & CStr(adOpenUnspecified)
    
    ValidationControl.AddCode "Const adLockBatchOptimistic = " & CStr(adLockBatchOptimistic)
    ValidationControl.AddCode "Const adLockOptimistic = " & CStr(adLockOptimistic)
    ValidationControl.AddCode "Const adLockPessimistic = " & CStr(adLockPessimistic)
    ValidationControl.AddCode "Const adLockReadOnly = " & CStr(adLockReadOnly)
    ValidationControl.AddCode "Const adLockUnspecified = " & CStr(adLockUnspecified)
    
    ValidationControl.AddCode "Const adAsyncExecute = " & CStr(adAsyncExecute)
    ValidationControl.AddCode "Const adAsyncFetch = " & CStr(adAsyncFetch)
    ValidationControl.AddCode "Const adAsyncFetchNonBlocking = " & CStr(adAsyncFetchNonBlocking)
    ValidationControl.AddCode "Const adExecuteNoRecords = " & CStr(adExecuteNoRecords)
'    ValidationControl.AddCode "Const adExecuteStream = " & CStr(adExecuteStream)
'    ValidationControl.AddCode "Const adExecuteRecord = " & CStr(adExecuteRecord)
    ValidationControl.AddCode "Const adOptionUnspecified = " & CStr(adOptionUnspecified)
    
    ValidationControl.AddCode "Const adCmdUnspecified = " & CStr(adCmdUnspecified)
    ValidationControl.AddCode "Const adCmdText = " & CStr(adCmdText)
    ValidationControl.AddCode "Const adCmdTable = " & CStr(adCmdTable)
    ValidationControl.AddCode "Const adCmdStoredProc = " & CStr(adCmdStoredProc)
    ValidationControl.AddCode "Const adCmdUnknown = " & CStr(adCmdUnknown)
    ValidationControl.AddCode "Const adCmdFile = " & CStr(adCmdFile)
    ValidationControl.AddCode "Const adCmdTableDirect = " & CStr(adCmdTableDirect)
    
    ValidationControl.AddCode "Const  vbKeyF1= " & CStr(vbKeyF1)
    ValidationControl.AddCode "Const  vbKeyF2= " & CStr(vbKeyF2)
    ValidationControl.AddCode "Const  vbKeyF3= " & CStr(vbKeyF3)
    ValidationControl.AddCode "Const  vbKeyF4= " & CStr(vbKeyF4)
    ValidationControl.AddCode "Const  vbKeyF5= " & CStr(vbKeyF5)
    ValidationControl.AddCode "Const  vbKeyF6= " & CStr(vbKeyF6)
    ValidationControl.AddCode "Const  vbKeyF7= " & CStr(vbKeyF7)
    ValidationControl.AddCode "Const  vbKeyF8= " & CStr(vbKeyF8)
    ValidationControl.AddCode "Const  vbKeyF9= " & CStr(vbKeyF9)
    ValidationControl.AddCode "Const  vbKeyF10= " & CStr(vbKeyF10)
    ValidationControl.AddCode "Const  vbKeyF11= " & CStr(vbKeyF11)
    ValidationControl.AddCode "Const  vbKeyF12= " & CStr(vbKeyF12)
    ValidationControl.AddCode "Const  vbKeyEscape= " & CStr(vbKeyEscape)
'    ValidationControl.AddCode "Const  = " & CStr()
'    ValidationControl.AddCode "Const  = " & CStr()
'    ValidationControl.AddCode "Const  = " & CStr()
    
'-----------------
    
    Line = 500
    ValidationControl.AddObject "Form", Me, True
    ValidationControl.AddObject "Buffers", TrnBuffers, True
    ValidationControl.AddObject "AppBuffers", GenWorkForm.AppBuffers, True
    
    ValidationControl.AddObject "AppRecordset", GenWorkForm.AppRS, True
    ValidationControl.AddObject "AppStoredProcedure", GenWorkForm.AppSP, True
    
    Line = 600
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
        Set SPCPanel = CreateObject("SPCPanelXControl.SPCPanelX")
        SPCPanel.host = cPRINTERSERVER
        'SPCPanel.Port = 999
        SPCPanel.Port = cPrinterPort
    End If
    
    Line = 700
    TRNOk = False
    
    LastValidChk = True: Set LastChkControl = Nothing
    
    QTrn = 0
    
    KeyProcessStarted = False
    CloseTransactionFlag = False
    RefreshCom.Enabled = True
    Exit Sub

ExitPoint:
    MsgBox "InitForm Line: " & Line & Err.number & Err.description, vbCritical, "ΛΑΘΟΣ"
    CloseTransactionFlag = True
End Sub


Private Sub CommandToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, foundflag As Boolean
    If Button.key = "F9" Then
        SendKeys "{F9}"
    ElseIf Button.key = "F10" Then
'        If cNewJournalType = False Then
'            eJournalFrm.Show vbModal, Me
'        Else
            Dim aTRNHandler As New L2TrnHandler
            aTRNHandler.ExecuteForm "9989"
            aTRNHandler.CleanUp
            Set aTRNHandler = Nothing
'        End If
    ElseIf Button.key = "F11" Then
        Dim bTRNHandler As New L2TrnHandler
        bTRNHandler.ExecuteForm "9747"
        bTRNHandler.CleanUp
        Set bTRNHandler = Nothing
    ElseIf Button.key = "F12" Then
        ActiveControl.FinalizeEdit: SEND
    ElseIf Button.key = "Ctrl-L" Then
        If SessID < 99 Then
            SessID = SessID + 1
        Else
            SessID = 1
        End If
        
        Dim aSelectFrm As New SelectTRNFrm
        aSelectFrm.Show vbModal, Me
        Set aSelectFrm = Nothing
        
        If SessID > 1 Then
            SessID = SessID - 1
        End If
    End If
End Sub


Private Sub Form_Activate()
    If ActivateOp = fopSendBuffer Then
        Form_KeyDown vbKeyF12, 0
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, foundflag As Boolean
    
'On Error GoTo ScriptError
    On Error Resume Next
    Dim aFldName As String
    aFldName = UCase(ActiveControl.name)
    On Error GoTo ScriptError
    
    
    If KeyProcessStarted Then Exit Sub
    KeyProcessStarted = True
    
'    If OpAfterHandling = fopExitForm Then KeyCode = 0: Exit Sub
    OpAfterHandling = fopNoOperation
    If AfterKeyFlag Then KeyCode = ValidationControl.Run("AfterKey_Script", aFldName, KeyCode, Shift)
    
'    If OpAfterHandling = fopExitForm Then Unload Me: Exit Sub Else OpAfterHandling = fopNoOperation
    If OpAfterHandling = fopExitForm Then
        CurrAction = taExit_Form:
        Unload Me:
        Exit Sub
    ElseIf OpAfterHandling = fopSendBuffer Then
        OpAfterHandling = fopNoOperation
        NextAction = taSend_Buffer:
        KeyProcessStarted = False
        ProcessLoop
        Exit Sub
    ElseIf OpAfterHandling = fopMoveNextPhase Then
        OpAfterHandling = fopNoOperation
        CurrAction = taExit_Form:
        'ProcessLoop: Exit Sub
        
        
                If processphase = MaxPhaseNum Then
                    If Not DisableWriteJournal Then eJournalWriteFinal Me
                    If AutoExitFlag Then
                        Unload Me:
                    ElseIf RestartEditFlag Then
                        NextAction = taGet_Input: processphase = 1
                        ShowFields processphase, True
                        NextAction = taSend_Buffer
                        
                    Else
                        NextAction = taStay_In_Form
                    End If
                Else
                    processphase = processphase + 1: NextAction = taGet_Input: CurrAction = taGet_Input
                    ShowFields processphase, True
                    NextAction = taSend_Buffer
                    
                End If
            KeyProcessStarted = False
            Exit Sub
        
    End If
    
    KeyProcessStarted = False

On Error GoTo 0
    If KeyCode = vbKeyF10 Then
'        If cNewJournalType = False Then
'            eJournalFrm.Show vbModal, Me
'        Else
            Dim aTRNHandler As New L2TrnHandler
            aTRNHandler.ExecuteForm "9989"
            aTRNHandler.CleanUp
            Set aTRNHandler = Nothing
'        End If
    ElseIf KeyCode = vbKeyF11 Then
        Dim bTRNHandler As New L2TrnHandler
        bTRNHandler.ExecuteForm "9747"
        bTRNHandler.CleanUp
        Set bTRNHandler = Nothing
    ElseIf KeyCode = vbKeyF12 Then
        If (Shift And vbCtrlMask) > 0 And cDebug > 0 Then
            GetInputFromRTF
        Else
            If Not (ActiveControl Is Nothing) Then
                On Error Resume Next
                ActiveControl.FinalizeEdit
                On Error GoTo 0
            End If
            SEND
        End If
    ElseIf KeyCode = 65 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-a
        KeyCode = 0
        Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = GenWorkForm.AppBuffers
        BufferViewer.Show vbModal, Me
        Unload BufferViewer
    ElseIf KeyCode = 66 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-b
        KeyCode = 0
        Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = TrnBuffers
        BufferViewer.Show vbModal, Me
        Unload BufferViewer
    ElseIf KeyCode = 76 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-l
        If SessID < 99 Then
            SessID = SessID + 1
        Else
            SessID = 1
        End If
        
        KeyCode = 0
        Dim aSelectFrm As New SelectTRNFrm
        aSelectFrm.Show vbModal, Me
        Set aSelectFrm = Nothing
    
        If SessID > 1 Then
            SessID = SessID - 1
        End If
    ElseIf KeyCode = 86 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-v
        On Error Resume Next
        SPCPanel.ShowServer: KeyCode = 0
    ElseIf KeyCode = 72 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-h
        On Error Resume Next
        SPCPanel.HideServer: KeyCode = 0
    ElseIf KeyCode = 84 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-t
        On Error Resume Next
        SPCPanel.StartPrinter: KeyCode = 0
    ElseIf KeyCode = 80 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-p
        On Error Resume Next
        SPCPanel.StopPrinter: KeyCode = 0
    ElseIf KeyCode = 67 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-c
        KeyCode = 0
        SpecialKey = cCHIEFKEY:
        On Error Resume Next
        ActiveControl.FinalizeEdit:
        On Error GoTo 0
        SEND
    ElseIf KeyCode = 77 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-m
        KeyCode = 0
        SpecialKey = cTELLERMANAGERKEY
        On Error Resume Next
        ActiveControl.FinalizeEdit:
        On Error GoTo 0
        SEND
    ElseIf KeyCode = 83 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-s
        MailSlotFrm.Show vbModal, Me
        'SendToWinPopUp "u34000 Marinos Giorgos", "n0yyy0051", "test message"
' parms: PopFrom: user or computer that sends the message
' PopTo: computer that receives the  message
' MsgText: the text of the message to send
    ElseIf KeyCode = vbKeyEscape Then
        If CurrAction = taGet_Input Or CurrAction = taStay_In_Form Or CurrAction = taExit_Form Then
            DoEvents
            NextAction = taEscape_Form
            ProcessLoop
        End If
    ElseIf (KeyCode = vbKeyDown) Then
        If ((Shift And vbAltMask) > 0) Then
        foundflag = False
        If Lists.count > 0 And CurrAction <> taStay_In_Form Then
            For i = 1 To Lists.count
                If Lists(i).IsVisible Then
                    DoEvents
                    KeyCode = 0
                    Lists(i).SetFocus
                    foundflag = True
                    Exit For
                End If
            Next i
            If Not foundflag Then
            For i = 1 To Spreads.count
                If Spreads(i).IsVisible Then
                    DoEvents
                    KeyCode = 0
                    Spreads(i).SetFocus
                    foundflag = True
                    Exit For
                End If
            Next i
            End If
        End If
        End If
    ElseIf (KeyCode = vbKeyUp) Then
        If ((Shift And vbAltMask) > 0) Then
        If fields.count > 0 And CurrAction <> taStay_In_Form Then
            For i = 1 To fields.count
                If fields(i).IsVisible(processphase) And fields(i).IsEditable(processphase) Then
                    DoEvents
                    fields(i).SetFocus
                    Exit For
                End If
            Next i
        End If
        End If
    End If
    
Exit Sub

ScriptError:
MsgBox "ERRERR"
Call NBG_LOG_MsgBox("Error :" & CStr(Err.number) & Err.description & "-" & CStr(ValidationControl.error.number) & ValidationControl.error.description, True)
MsgBox "ERRERR2"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        If Lists.count > 0 Then
            Lists(1).PrintLines
            Lists(1).UnLockPrinter Me

            sbWriteStatusMessage ""
            KeyCode = 0
        ElseIf Spreads.count = 1 Then
            Spreads(1).PrintLines
            Spreads(1).UnLockPrinter Me

            sbWriteStatusMessage ""
            KeyCode = 0
        End If
    End If
End Sub

Private Sub Form_Resize()
    Status.Panels(1).width = Round(Status.width * 0.9)
End Sub

Public Function CreateBrowser(name As String, Left As Long, Top As Long, width As Long, height As Long) As GenBrowser
    
    Dim abrowser As Variant
    Set abrowser = Controls.add("Shine.Genbrowser", name)
    abrowser.Initialize Me, ValidationControl, name, Left, Top, width, height
    Browsers.add abrowser, name
    
    'browser.Visible = True
    Set CreateBrowser = abrowser
End Function

Public Function CreateRichTextBox(name As String, Left As Long, Top As Long, width As Long, height As Long) As GenRichTextBox
    
    Dim aRichTextBox As Variant
    Set aRichTextBox = Controls.add("Shine.GenRichTextBox", name)
    aRichTextBox.Initialize Me, ValidationControl, name, Left, Top, width, height
    RichTextBoxes.add aRichTextBox, name
    
    Set CreateRichTextBox = aRichTextBox
End Function

Private Sub PrepareFromXML()
If CloseTransactionFlag Then Exit Sub
Dim aelm As IXMLDOMAttribute

On Error GoTo InvalidTransactionData

trnXML.preserveWhiteSpace = True
trnXML.Load (ReadDir & CStr(SelectedTRN) & ".xml")

trnXMLVersion = "1"
If trnXML.documentElement.Attributes.length > 0 Then
    Set aelm = trnXML.documentElement.Attributes.getNamedItem("VER")
    If Not (aelm Is Nothing) Then trnXMLVersion = aelm.value
End If


Set trnNode = trnXML.documentElement.selectSingleNode("TRN")
Set stepsNode = trnXML.documentElement.selectSingleNode("STEP")
Set listsNode = trnXML.documentElement.selectSingleNode("LIST")
Set gridsNode = trnXML.documentElement.selectSingleNode("GRID")

Caption = CStr(cTRNCode) & " - " & NodeStringFld(trnNode, "Name", trnModel) & "  (" & NodeStringFld(trnNode, "VersionTime", trnModel) & ")" & "V" & trnXMLVersion

Select Case NodeIntegerFld(trnNode, "RequiredKey", trnModel)
Case tkNoKey: trn_key = " "
Case tkTellerKey: trn_key = cTELLERKEY
Case tkChiefKey: trn_key = cCHIEFKEY
Case tkManagerKey: trn_key = cMANAGERKEY
Case tkTellerChiefKey: trn_key = cTELLERCHIEFKEY
Case tkTellerManagerKey: trn_key = cTELLERMANAGERKEY
End Select

EncodeGreekflag = NodeBooleanFld(trnNode, "EncodeGreek", trnModel)
AutoExitFlag = NodeBooleanFld(trnNode, "AutoExit", trnModel)
RestartEditFlag = NodeBooleanFld(trnNode, "RestartEdit", trnModel)
FldLengthInBuffer = NodeBooleanFld(trnNode, "FldLengthInBuffer", trnModel)
TotalName = NodeStringFld(trnNode, "TotalName", trnModel)
TotalPos = NodeIntegerFld(trnNode, "TotalPos", trnModel)
BonusScale = NodeIntegerFld(trnNode, "BonusScale", trnModel)
BonusRegPhase = NodeIntegerFld(trnNode, "BonusRegPhase", trnModel)
BonusRegPos = NodeIntegerFld(trnNode, "BonusRegPos", trnModel)

HiddenFlag = NodeBooleanFld(trnNode, "Hidden", trnModel)
If HiddenFlag And Not cEnableHiddenTransactions Then CloseTransactionFlag = True: Exit Sub

Dim Phase As Integer
Dim i As Integer, k As Integer, l As Integer
Dim afld As Variant, aFldCD As Integer, FldName As String
Dim aLbl As Variant, aLblCD As Integer, LblName As String
Dim aBtn As Variant, aBtnCD As Integer, BtnName As String
Dim aChk As Variant, aChkCD As Integer, ChkName As String
Dim aCmb As Variant, aCmbCD As Integer, CmbName As String
Dim aChr As Variant, aChrCD As Integer, ChrName As String
Dim foundflag As Boolean
Dim astr

If trnXMLVersion <> "1" Then
    On Error GoTo AllScriptError
    astr = NodeStringFld(trnNode, "AllScript", trnModel)
    If astr <> "" Then ValidationControl.AddCode astr
    For i = 1 To ValidationControl.Procedures.count
        Select Case UCase(ValidationControl.Procedures.item(i).name)
        Case UCase("STARTUP_SCRIPT"):  StartupCodeFlag = True
        Case UCase("Validation_Script"):  FormValidationFlag = True
        Case UCase("BeforeOut_Script"):  BeforeOutFlag = True
        Case UCase("AfterIn_Script"):  AfterInFlag = True
        Case UCase("CommunicationError_Script"):  CommunicationErrorFlag = True
        Case UCase("BeforeAction_Script"):  BeforeActionFlag = True
        Case UCase("AfterAction_Script"):  AfterActionFlag = True
        Case UCase("AfterKey_Script"):  AfterKeyFlag = True
        Case UCase("BeforePrint_Script"):  BeforePrintFlag = True
        Case UCase("BeforeDocumentPrint_Script"): BeforeDocumentFlag = True
        Case UCase("AfterDocumentPrint_Script"): AfterDocumentFlag = True
        Case UCase("LostFocus_Script"):  LostFocusFlag = True
        End Select
    Next i
Else
    On Error GoTo StartUpScriptError
    astr = NodeStringFld(trnNode, "StartupScript", trnModel)
    If astr <> "" Then
        StartupCodeFlag = True
        ValidationControl.AddCode "Public Sub StartUp_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo ValidationScriptError
    astr = NodeStringFld(trnNode, "ValidationScript", trnModel)
    If astr <> "" Then
        FormValidationFlag = True
        ValidationControl.AddCode "Public Sub Validation_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo BeforeOutScriptError
    astr = NodeStringFld(trnNode, "BeforeOutScript", trnModel)
    If astr <> "" Then
        BeforeOutFlag = True
        ValidationControl.AddCode "Public function BeforeOut_Script " & vbCrLf & astr & vbCrLf & _
            "BeforeOut_Script = TRUE" & vbCrLf & "End function"
    End If
    On Error GoTo AfterInScriptError
    astr = NodeStringFld(trnNode, "AfterInScript", trnModel)
    If astr <> "" Then
        AfterInFlag = True
        ValidationControl.AddCode "Public Sub AfterIn_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo CommunicationErrorScriptError
    astr = NodeStringFld(trnNode, "CommunicationErrorScript", trnModel)
    If astr <> "" Then
        CommunicationErrorFlag = True
        ValidationControl.AddCode "Public Sub CommunicationError_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo BeforeActionError
    astr = NodeStringFld(trnNode, "BeforeActionScript", trnModel)
    If astr <> "" Then
        BeforeActionFlag = True
        ValidationControl.AddCode "Public Sub BeforeAction_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo AfterActionError
    astr = NodeStringFld(trnNode, "AfterActionScript", trnModel)
    If astr <> "" Then
        AfterActionFlag = True
        ValidationControl.AddCode "Public Sub AfterAction_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    On Error GoTo AfterKeyError
    astr = NodeStringFld(trnNode, "AfterKeyScript", trnModel)
    If astr <> "" Then
        AfterKeyFlag = True
        ValidationControl.AddCode "Public function AfterKey_Script (inFldName, inKeyCode, inShift)" & vbCrLf & _
            astr & vbCrLf & "AfterKey_Script = inKeyCode" & vbCrLf & "End function"
    End If
    
    On Error GoTo BeforePrintError
    astr = NodeStringFld(trnNode, "BeforePrintScript", trnModel)
    If astr <> "" Then
        BeforePrintFlag = True
        ValidationControl.AddCode "Public Sub BeforePrint_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    
    On Error GoTo BeforeDocumentError
    astr = NodeStringFld(trnNode, "BeforeDocumentScript", trnModel)
    If astr <> "" Then
        BeforeDocumentFlag = True
        ValidationControl.AddCode "Public Sub BeforeDocumentPrint_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    
    On Error GoTo AfterDocumentError
    astr = NodeStringFld(trnNode, "AfterDocumentScript", trnModel)
    If astr <> "" Then
        AfterDocumentFlag = True
        ValidationControl.AddCode "Public Sub AfterDocumentPrint_Script " & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    
    On Error GoTo LostFocusError
    astr = NodeStringFld(trnNode, "LostFocusScript", trnModel)
    If astr <> "" Then
        LostFocusFlag = True
        ValidationControl.AddCode "Public Sub LostFocus_Script (Sender)" & vbCrLf & astr & vbCrLf & "End Sub"
    End If
    
    On Error GoTo AddRoutinesError
    'astr = NodeStringFld(trnNode, "AddRoutines", trnModel)
    'If astr <> "" Then
    '    ValidationControl.AddCode astr
    'End If
End If
On Error GoTo GenErrorHandling:
'Πληροφορίες για τις φάσεις της συναλλαγής

For i = NamedControls.count To 1 Step -1: NamedControls.Remove (i): Next i

For i = fields.count To 1 Step -1: fields.Remove (i): Next i
For i = NamedFields.count To 1 Step -1: NamedFields.Remove (i): Next i
For i = Labels.count To 1 Step -1: Labels.Remove (i): Next i
For i = DocFields.count To 1 Step -1: DocFields.Remove (i): Next i

For i = 0 To stepsNode.childNodes.length - 1
    If i > 9 Then: Exit For
    Set stepnode = stepsNode.childNodes.item(i)
        
    TrnCode(i + 1) = NodeStringFld(stepnode, "TrnNum", stepModel)
    If TrnCode(i + 1) = "" Then
        TrnCode(i + 1) = NodeIntegerFld(trnNode, "CD", trnModel)
    End If
    MaxPhaseNum = NodeIntegerFld(stepnode, "StepCD", stepModel)

On Error Resume Next
    fieldsNode = Nothing
    Set fieldsNode = stepnode.selectSingleNode("FIELDS")
    If Not (fieldsNode Is Nothing) Then
On Error GoTo FieldLoadError:
'        If Not (fieldsNode.children Is Nothing) Then
        If fieldsNode.childNodes.length > 0 Then
            For k = 0 To fieldsNode.childNodes.length - 1
                Set fieldNode = fieldsNode.childNodes.item(k)
                aFldCD = NodeIntegerFld(fieldNode, "FldNo", fldModel)
                FldName = "Fld" & StrPad_(CStr(aFldCD), 3, "0", "L") 'NodeStringFld(fieldNode, "Name", fldModel)
                If i = 0 Then
                    Set afld = Controls.add("Shine.GenTextBox", FldName)
                    fields.add afld, FldName
                Else
                    Set afld = fields.item(FldName)
                End If
                afld.InitializeFromXML Me, ValidationControl, fieldNode, i + 1
                
                If i = 0 And afld.FldName2 <> "" Then NamedFields.add afld, afld.FldName2
                If i = 0 And afld.FldName2 <> "" Then NamedControls.add afld, afld.FldName2
                
                If i = 0 And afld.TTabIndex > 0 And i = 0 Then TabbedControls.add afld
                
                If (afld.DocX * afld.DocY * afld.DocWidth * afld.DocHeight) _
                   + (afld.TitleX * afld.TitleY * afld.TitleWidth * afld.TitleHeight) <> 0 Then
                   DocFields.add afld
                End If
            Next k
        End If
'        End If
    End If
    
On Error Resume Next
    btnsNode = Nothing
    Set btnsNode = stepnode.selectSingleNode("BUTTONS")
    If Not (btnsNode Is Nothing) Then
On Error GoTo FieldLoadError:
'        If Not (btnsNode.children Is Nothing) Then
        If btnsNode.childNodes.length > 0 Then
            For k = 0 To btnsNode.childNodes.length - 1
                Set btnNode = btnsNode.childNodes.item(k)
                
                aBtnCD = NodeIntegerFld(btnNode, "BtnNo", btnModel)
                BtnName = "Btn" & StrPad_(CStr(aBtnCD), 3, "0", "L")
                'NodeStringFld(BtnNode, "Name", BtnModel)
                If i = 0 Then
                    Set aBtn = Controls.add("Shine.GenBtn", BtnName)
                    Buttons.add aBtn, BtnName
                Else
                    Set aBtn = Buttons.item(BtnName)
                End If
                aBtn.InitializeFromXML Me, ValidationControl, btnNode, i + 1
                If i = 0 And aBtn.BtnName2 <> "" Then NamedControls.add aBtn, aBtn.BtnName2
                'If i = 0 And aBtn.TTabIndex > 0 And i = 0 Then TabbedControls.Add aBtn
                
            Next k
        End If
'        End If
    End If
    
On Error Resume Next
    chrsNode = Nothing
    Set chrsNode = stepnode.selectSingleNode("CHARTS")
    If Not (chrsNode Is Nothing) Then
On Error GoTo ChartLoadError:
'        If Not (chrsNode.children Is Nothing) Then
        If chrsNode.childNodes.length > 0 Then
            For k = 0 To chrsNode.childNodes.length - 1
                Set chrNode = chrsNode.childNodes.item(k)
                
                aChrCD = NodeIntegerFld(chrNode, "ChartNo", chrModel)
                ChrName = "Chr" & StrPad_(CStr(aChrCD), 3, "0", "L")
                'NodeStringFld(BtnNode, "Name", BtnModel)
                If i = 0 Then
                    Set aChr = Controls.add("Shine.GenChart", ChrName)
                    Charts.add aChr, ChrName
                Else
                    Set aChr = Charts.item(ChrName)
                End If
                aChr.InitializeFromXML Me, ValidationControl, chrNode, i + 1
                If i = 0 And aChr.ChrName2 <> "" Then NamedControls.add aChr, aChr.ChrName2
                If i = 0 And aChr.TTabIndex > 0 And i = 0 Then TabbedControls.add aChr
                
            Next k
        End If
'        End If
    End If
    
    
On Error Resume Next
    chksNode = Nothing
    Set chksNode = stepnode.selectSingleNode("CHECKS")
    If Not (chksNode Is Nothing) Then
On Error GoTo FieldLoadError:
'        If Not (chksNode.children Is Nothing) Then
        If chksNode.childNodes.length > 0 Then
            For k = 0 To chksNode.childNodes.length - 1
                Set chkNode = chksNode.childNodes.item(k)
                
                aChkCD = NodeIntegerFld(chkNode, "ChkNo", chkModel)
                ChkName = "Chk" & StrPad_(CStr(aChkCD), 3, "0", "L")
                'NodeStringFld(BtnNode, "Name", BtnModel)
                If i = 0 Then
                    Set aChk = Controls.add("Shine.GenCheck", ChkName)
                    Checks.add aChk, ChkName
                Else
                    Set aChk = Checks.item(ChkName)
                End If
                aChk.InitializeFromXML Me, ValidationControl, chkNode, i + 1
                If i = 0 And aChk.CHKName2 <> "" Then NamedControls.add aChk, aChk.CHKName2
                If i = 0 And aChk.TTabIndex > 0 And i = 0 Then TabbedControls.add aChk
                
            Next k
        End If
'        End If
    End If
    
On Error Resume Next
    cmbsNode = Nothing
    Set cmbsNode = stepnode.selectSingleNode("COMBOS")
    If Not (cmbsNode Is Nothing) Then
On Error GoTo FieldLoadError:
'        If Not (cmbsNode.children Is Nothing) Then
        If cmbsNode.childNodes.length > 0 Then
            For k = 0 To cmbsNode.childNodes.length - 1
                Set cmbNode = cmbsNode.childNodes.item(k)
                
                aCmbCD = NodeIntegerFld(cmbNode, "CmbNo", cmbModel)
                CmbName = "CMB" & StrPad_(CStr(aCmbCD), 3, "0", "L")
                'NodeStringFld(BtnNode, "Name", BtnModel)
                If i = 0 Then
                    Set aCmb = Controls.add("Shine.GenCombo", CmbName)
                    Combos.add aCmb, CmbName
                Else
                    Set aCmb = Combos.item(CmbName)
                End If
                aCmb.InitializeFromXML Me, ValidationControl, cmbNode, i + 1
                If i = 0 And aCmb.CMBName2 <> "" Then NamedControls.add aCmb, aCmb.CMBName2
                If i = 0 And aCmb.TTabIndex > 0 And i = 0 Then TabbedControls.add aCmb

            Next k
        End If
'        End If
    End If
    
On Error Resume Next
    labelsNode = Nothing
    Set labelsNode = stepnode.selectSingleNode("LABELS")
    If Not (labelsNode Is Nothing) Then
On Error GoTo LabelLoadError:
        Set labelsNode = stepnode.selectSingleNode("LABELS")
'        If Not (labelsNode.children Is Nothing) Then
            If labelsNode.childNodes.length > 0 Then
                For k = 0 To labelsNode.childNodes.length - 1
                    Set labelNode = labelsNode.childNodes.item(k)
                    aLblCD = NodeIntegerFld(labelNode, "LabelNo", lblModel)
                    LblName = "LBL" & StrPad_(CStr(aLblCD), 3, "0", "L")
                    If i = 0 Then
                        Set aLbl = Controls.add("Shine.GenLabel", LblName)
                        Labels.add aLbl, LblName
                    Else
On Error Resume Next
                        Set aLbl = Labels.item(LblName)
On Error GoTo LabelLoadError:
                    End If
                    aLbl.InitializeFromXML Me, ValidationControl, labelNode, i + 1
                    If i = 0 And aLbl.LabelName <> "" Then NamedControls.add aLbl, aLbl.LabelName
                Next k
            End If
'        End If
    End If
Next i


MaxPhaseNum = stepsNode.childNodes.length

On Error GoTo ListLoadError:
'Πληροφορίες για τα ListBoxes της συναλλαγής
Dim aLst As Variant
Dim LstNo As Integer, LstName As String

For i = Lists.count To 1 Step -1
    Lists.Remove (i)
Next i

If Not (listsNode Is Nothing) Then
'If Not (listsNode.children Is Nothing) Then
On Error GoTo ListLoadError:
    If listsNode.childNodes.length > 0 Then
        For i = 0 To listsNode.childNodes.length - 1
            Set listNode = listsNode.childNodes.item(i)
            LstNo = NodeIntegerFld(listNode, "LstNo", listModel)
            LstName = "Lst" & StrPad_(CStr(LstNo), 3, "0", "L")
'            LstName = NodeStringFld(listNode, "Name", listModel)
            Set aLst = Controls.add("Shine.GenListBox", LstName)
            Lists.add aLst, LstName
            aLst.InitializeFromXML Me, ValidationControl, listNode
            If aLst.LstName2 <> "" Then NamedControls.add aLst, aLst.LstName2
            If aLst.TTabIndex > 0 Then TabbedControls.add aLst
        Next i
    End If
'End If
End If

On Error GoTo GridLoadError:
'Πληροφορίες για τα Grids της συναλλαγής
Dim aSpread As Variant
Dim SprdNo As Integer, SprdName As String

For i = Spreads.count To 1 Step -1
    Spreads.Remove (i)
Next i
If Not (gridsNode Is Nothing) Then
'If Not (gridsNode.children Is Nothing) Then
On Error GoTo GridLoadError:
    If gridsNode.childNodes.length > 0 Then
        For i = 0 To gridsNode.childNodes.length - 1
            Set gridNode = gridsNode.childNodes.item(i)
            SprdNo = NodeIntegerFld(gridNode, "SprdNo", gridModel)
            SprdName = "Spd" & StrPad_(CStr(SprdNo), 3, "0", "L") 'NodeStringFld(gridNode, "Name", gridModel)
            
            Set aSpread = Controls.add("Shine.GenSpread", SprdName)
            Spreads.add aSpread, SprdName
            aSpread.InitializeFromXML Me, ValidationControl, gridNode
            If aSpread.SprdName2 <> "" Then NamedControls.add aSpread, aSpread.SprdName2
            If aSpread.TTabIndex > 0 Then TabbedControls.add aSpread
        Next i
    End If
End If
'End If

Dim aItem
If TabbedControls.count > 0 Then
SortAgain:
    k = 0
    For i = 1 To TabbedControls.count - 1
        If TabbedControls(i).TTabIndex > TabbedControls(i + 1).TTabIndex Then
            Set aItem = TabbedControls(i)
            TabbedControls.add aItem, , , i + 1
            TabbedControls.Remove (i)
            k = k + 1
        End If
    Next i
    If k > 0 Then GoTo SortAgain
    For i = 1 To TabbedControls.count
        TabbedControls(i).TabIndex = i
    Next i
End If
Set aItem = Nothing
'TabbedControls(1).SetFocus
    


'Στοιχεία Πεδίων απο προηγούμενη συναλλαγή
If TRNQueue.count > 0 Then
    For i = 1 To TRNFldNoQueue.count
        For k = 1 To fields.count
            If CStr(fields(k).FldNo) = TRNFldNoQueue(i) Then fields(k).Text = TRNFldTextQueue(i)
        Next k
    Next i
    For i = TRNQueue.count To 1 Step -1: TRNQueue.Remove (i): Next i
    For i = TRNFldNoQueue.count To 1 Step -1: TRNFldNoQueue.Remove (i): Next i
    For i = TRNFldTextQueue.count To 1 Step -1: TRNFldTextQueue.Remove (i): Next i
End If
If NodeIntegerFld(trnNode, "QTrn", trnModel) Then QTrn = NodeIntegerFld(trnNode, "QTrn", trnModel)

ExitPoint:
Exit Sub
GenErrorHandling:
Call NBG_LOG_MsgBox("Error :" & Err.description, True)
GoTo ExitPoint
HandleScriptErrors:
Call NBG_LOG_MsgBox("Error :" & ValidationControl.error.description, True)
GoTo ExitPoint
AllScriptError:
Call NBG_LOG_MsgBox("Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
StartUpScriptError:
Call NBG_LOG_MsgBox("Startup Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
ValidationScriptError:
Call NBG_LOG_MsgBox("Validation Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
BeforeOutScriptError:
Call NBG_LOG_MsgBox("Before Send Buffer Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
AfterInScriptError:
Call NBG_LOG_MsgBox("After Receive Buffer Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
CommunicationErrorScriptError:
Call NBG_LOG_MsgBox("After Receive Buffer Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
BeforeActionError:
Call NBG_LOG_MsgBox("Before Action Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
AfterActionError:
Call NBG_LOG_MsgBox("After Action Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
AfterKeyError:
Call NBG_LOG_MsgBox("After Key Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
LostFocusError:
Call NBG_LOG_MsgBox("Lost Focus Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
AfterDocumentError:
Call NBG_LOG_MsgBox("After Document Print Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
BeforePrintError:
Call NBG_LOG_MsgBox("Before Print Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
BeforeDocumentError:
Call NBG_LOG_MsgBox("Before Document Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
AddRoutinesError:
Call NBG_LOG_MsgBox("Additional Routines Code parse Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
StartUpExecError:
Call NBG_LOG_MsgBox("Startup Code run Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
FieldLoadError:
Call NBG_LOG_MsgBox("Field Data Loading Error :" & FldName & vbCrLf & ValidationControl.error.description, True)
Resume Next
LabelLoadError:
Call NBG_LOG_MsgBox("Label Data Loading Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
ListLoadError:
Call NBG_LOG_MsgBox("List Data Loading Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
GridLoadError:
Call NBG_LOG_MsgBox("Grid Data Loading Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
ChartLoadError:
Call NBG_LOG_MsgBox("Chart Data Loading Error :" & vbCrLf & ValidationControl.error.description, True)
Resume Next
InvalidTransactionData:
    CloseTransactionFlag = True
GoTo ExitPoint
HelpLoadError:
End Sub

Private Sub Form_Load()

DisableTRNCounterUpdate = False

On Error GoTo GenErrorHandling
Status.Panels(2).Visible = GenWorkForm.vStatus.Panels(2).Visible
Status.Panels(3).Visible = GenWorkForm.vStatus.Panels(3).Visible
InitForm

'If CloseTransactionFlag Then GoTo before_exit_sub
If CloseTransactionFlag Then GoTo ExitPoint

PrepareFromXML
If CloseTransactionFlag Then GoTo before_exit_sub

On Error GoTo StartUpExecError
OpAfterHandling = fopNoOperation: ActivateOp = fopNoOperation:
If StartupCodeFlag Then ValidationControl.Run "StartUp_Script"
If OpAfterHandling = fopExitForm Then CloseTransactionFlag = True: GoTo before_exit_sub

If UseIRISUpdateFiles Then prepareIRISUpdate

'το καθαρισμα του buffer γίνεται μετά την εκτέλεση του startup_script
'ώστε να είναι δυνατή η μεταφορά στοιχείων σε λίστα της επόμενης συναλλαγής
Dim i As Integer
If ReceivedData.count > 0 Then For i = ReceivedData.count To 1 Step -1: ReceivedData.Remove i: Next i
If G0Data.count > 0 Then For i = G0Data.count To 1 Step -1: G0Data.Remove i: Next i

'If GenWorkForm.HookKeyboard Then SendKeys GenWorkForm.KeyboardString: GenWorkForm.HookKeyboard = False

On Error GoTo GenErrorHandling
NextAction = taGet_Input
processphase = 1
'If Fields.count + Lists.count + Spreads.count = 0 Then UCImage.Visible = True

If OpAfterHandling = fopSendBuffer Then NextAction = taSend_Buffer

ProcessLoop

before_exit_sub:
If CloseTransactionFlag Then Unload Me
GoTo ExitPoint
GenErrorHandling:
Call NBG_LOG_MsgBox("Error :" & Err.description, True)
GoTo ExitPoint
StartUpExecError:
Call NBG_LOG_MsgBox("Startup Code run Error :" & vbCrLf & _
    ValidationControl.error.description, True)
GoTo ExitPoint
HelpLoadError:

ExitPoint:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (TrnBuffers Is Nothing) Then
    TrnBuffers.ClearAll
    Set TrnBuffers = Nothing
End If

ExitPhaseStarted = True
RefreshCom.Enabled = False
'StopWritePipe = True
'StopReadPipe = True
DoEvents
cTRNCode = 0

Dim i As Integer
If CurrAction = taExit_Form And QTrn <> 0 Then
    For i = TRNQueue.count To 1 Step -1: TRNQueue.Remove i: Next i
    TRNQueue.add QTrn
    For i = 1 To fields.count
        If fields(i).QFldNo <> 0 Then
            TRNFldNoQueue.add item:=CStr(fields(i).QFldNo)
            TRNFldTextQueue.add item:=fields(i).Text
        End If
    Next i
    LockWindowUpdate (Me.hwnd)
    
'For i = Fields.count To 1 Step -1: Fields.Remove (i): Next i
'For i = OutFields.count To 1 Step -1: OutFields.Remove (i): Next i
'For i = Lists.count To 1 Step -1: Lists.Remove (i): Next i
'For i = Spreads.count To 1 Step -1: Spreads.Remove (i): Next i
    
    
'    LoadNewTransaction CStr(QTrn)
'    Cancel = 1
'    Exit Sub
End If

For i = fields.count To 1 Step -1: fields.Remove (i): Next i
For i = OutFields.count To 1 Step -1: OutFields.Remove (i): Next i
For i = Lists.count To 1 Step -1: Lists.Remove (i): Next i
For i = Spreads.count To 1 Step -1: Spreads.Remove (i): Next i
For i = Buttons.count To 1 Step -1: Buttons.Remove (i): Next i
For i = Checks.count To 1 Step -1: Checks.Remove (i): Next i
For i = Combos.count To 1 Step -1: Combos.Remove (i): Next i
For i = TabbedControls.count To 1 Step -1: TabbedControls.Remove (i): Next i
    
Set aTotalEntries = Nothing
If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then Set SPCPanel = Nothing
For i = TrnVariables.count To 1 Step -1
    Set TrnVariables.item(i).value = Nothing
    TrnVariables.Remove (i)
Next i

While GenWorkForm.AppBuffers.BufferNum > AppBuffersPos
    i = AppBuffersPos
    For i = AppBuffersPos + 1 To GenWorkForm.AppBuffers.BufferNum
        If Not (GenWorkForm.AppBuffers.ByIndex(i) Is Nothing) Then
            GenWorkForm.AppBuffers.FreeBuffer GenWorkForm.AppBuffers.ByIndex(i).name
            Exit For
        End If
    Next i
Wend


While GenWorkForm.AppRS.count > AppRSPos
    Dim ars As ADODB.Recordset
    For i = GenWorkForm.AppRS.count To AppRSPos + 1 Step -1
        Set ars = GenWorkForm.AppRS.item(i).rs
        On Error Resume Next: ars.Close: On Error GoTo 0
        Set GenWorkForm.AppRS.item(i).rs = Nothing
        GenWorkForm.AppRS.Remove (i)
        Set ars = Nothing
    Next i
Wend

While GenWorkForm.AppRS_S.count > AppRS_SPos
    For i = GenWorkForm.AppRS_S.count To AppRS_SPos + 1 Step -1
        AppRS_S.Remove (i)
    Next i
Wend

While GenWorkForm.AppSP.count > AppSPPos
    Dim acm As ADODB.command
    For i = GenWorkForm.AppSP.count To AppSPPos + 1 Step -1
        Set acm = GenWorkForm.AppSP.item(i).cm
        Set GenWorkForm.AppSP.item(i).cm = Nothing
        GenWorkForm.AppSP.Remove (i)
        Set acm = Nothing
    Next i
Wend

While GenWorkForm.AppVariables.count > AppVariablesPos
    For i = GenWorkForm.AppVariables.count To AppVariablesPos + 1 Step -1
        GenWorkForm.AppVariables.Remove (i)
    Next i
Wend

While GenWorkForm.AppCRS.count > AppCRSPos
   For i = GenWorkForm.AppCRS.count To AppCRSPos + 1 Step -1
      GenWorkForm.AppCRS.Remove (i)
   Next
Wend


KeyProcessStarted = False
End Sub

Private Function Validate(inPhase As Integer) As Boolean

Dim i As Integer, res As Boolean, failed As Boolean
Dim astr As String
    failed = False
    For i = 1 To fields.count
        If fields(i).IsEditable(inPhase) Then
            res = fields(i).ChkValid(inPhase)
            If Not res Then
                failed = True
                Beep
                Exit For
            End If
        End If
    Next i
    
    If (Not failed) And FormValidationFlag Then
        ValidOk = True
        ValidationError = ""
        ValidationControl.Run "Validation_Script"
        If Not ValidOk Then
            failed = True
            Beep
            If ValidationError = "" Then
                sbWriteStatusMessage "Απέτυχε ο ελεγχος πεδίων"
            Else
                sbWriteStatusMessage ValidationError
            End If
        End If
    End If
    Validate = Not failed
End Function

Public Function UpdateTrnNum()
    UpdateTrnNum_
End Function

Public Sub UpdateScreen()
    ShowFields (processphase)
End Sub

Private Sub PrepareDoc(inPhase As Integer)
Dim i As Integer, aX As Long, aY As Long, aW As Long, aH As Long
Dim aTX As Long, aTY As Long, aTW As Long, aTH As Long
Dim docText As String, docTitle As String, docSize As Integer, DocAlign As Integer, DocMask As String
Dim astr As String, bstr As String
If cPassbookPrinter = 0 Then Exit Sub
For i = 0 To DocumentLines - 1: DocLines(i) = String(255, " "): Next i
LastDocLine = 0
For i = 1 To DocFields.count
    With DocFields(i)
        aX = .DocX: aY = .DocY: aW = .DocWidth: aH = .DocHeight
        aTX = .TitleX: aTY = .TitleY: aTW = .TitleWidth: aTH = .TitleHeight
        docText = .FormatedText: docTitle = .Title: DocAlign = .DocAlign: DocMask = .DocMask
    End With
    If aX * aY * aW * aH <> 0 Then
        astr = DocLines(aY - 1): bstr = docText
        If DocAlign = 1 Then bstr = StrPad_(docText, CInt(aW), " ", "R") _
        Else If DocAlign = 2 Then bstr = StrPad_(docText, CInt(aW), " ", "L")
        bstr = Left$(bstr, aW): docSize = Len(bstr)
        astr = Left$(astr, aX - 1) & bstr & Right$(astr, 255 - aX - docSize + 1)
        DocLines(aY - 1) = astr
        If aY - 1 > LastDocLine Then LastDocLine = aY - 1
    End If
    If aTX * aTY * aTW * aTH <> 0 Then
        astr = DocLines(aTY - 1): bstr = docTitle
        bstr = StrPad_(docTitle, CInt(aTW), " ", "R")
        bstr = Left$(bstr, aTW): docSize = Len(bstr)
        astr = Left$(astr, aTX - 1) & bstr & Right$(astr, 255 - aTX - docSize + 1)
        DocLines(aTY - 1) = astr
        If aY - 1 > LastDocLine Then LastDocLine = aTY - 1
    End If
Next i
End Sub

Public Sub PrintDoc(inPhase As Integer)

    sbWriteStatusMessage "Εκτύπωση Παραστατικών...."
    docPrinting = True
    
    If BeforeDocumentFlag Then ValidationControl.Run "BeforeDocumentPrint_Script"
    If Not CancelPrintFlag And cPassbookPrinter <> 0 Then PrepareDoc inPhase: PrintDocLines_ Me
    If AfterDocumentFlag Then ValidationControl.Run "AfterDocumentPrint_Script"
    
    docPrinting = False
    sbWriteStatusMessage ""
    
End Sub

Public Sub ProcessLoop()
StartProcessLoop:
Do
    CurrAction = NextAction
    
    On Error GoTo BeforeActionError
    OpAfterHandling = fopNoOperation
    If BeforeActionFlag And (Not docPrinting) Then ValidationControl.Run "BeforeAction_Script"
    If OpAfterHandling = fopExitForm Then
        CurrAction = taExit_Form:
        Unload Me:
        Exit Sub
    End If
    
    On Error GoTo GenericError
    
    If Not docPrinting Then
        Select Case CurrAction
            Case taGet_Input
                ShowFields processphase, True
                NextAction = taSend_Buffer
                Exit Do
            Case taSend_Buffer
                If CommunicationStarted Then
                    NextAction = taGet_Input
                Else
                    SetDefaultFocus processphase    'Για να μην αλλάξει το focus σε περίπτωση λάθους οπότε θα
                                                    'γίνει το NextAction = taGet_Input και χαθεί το μύνημα λάθους
                    DoEvents
                    OpAfterHandling = fopNoOperation
                    If Not CancelCommunicationFlag Then
                        NextAction = taPrint_Document
                        If Not PrepareSendOut_(processphase) Then
                            NextAction = taGet_Input
                        Else
                            If CancelCommunicationFlag Then
                                NextAction = taPrint_Document
                            Else
                                If Flag610 And Flag620 And Flag630 Then
                                    If Not SendOut_(processphase) Then NextAction = taGet_Input _
                                    Else ReadIn_ (processphase)
                                Else
                                    If Not Flag610 Then
                                        sbWriteStatusMessage "ΔΕΝ ΔΟΘΗΚΕ 0610"
                                    ElseIf Not Flag620 Then
                                        sbWriteStatusMessage "ΔΕΝ ΔΟΘΗΚΕ 0620"
                                    ElseIf Not Flag630 Then
                                        sbWriteStatusMessage "ΔΕΝ ΔΟΘΗΚΕ 0630"
                                    End If
                                    NextAction = taGet_Input
                                End If
                            End If
                        End If
                    Else
                        If Not PrepareSendOut_(processphase) Then
                            NextAction = taGet_Input
                        Else
                            ReadIn_ (processphase): NextAction = taPrint_Document
                        End If
                    End If
                    If OpAfterHandling = fopSendBuffer Then
                        NextAction = taSend_Buffer
                    ElseIf OpAfterHandling = fopExitForm Then
                        NextAction = taExit_Form
                    ElseIf OpAfterHandling = fopCloseForm Then
                        NextAction = taEscape_Form
                    End If
                End If
            Case taPrint_Document
                If processphase = MaxPhaseNum Then WriteTotals
                
                If BeforePrintFlag Then BeforePrint_ (processphase)
                If processphase = MaxPhaseNum Then PrintDoc (processphase)
                NextAction = taExit_Form
            Case taExit_Form
                If processphase = MaxPhaseNum Then
'                    WriteTotals
                    If Not DisableWriteJournal Then eJournalWriteFinal Me
                    If AutoExitFlag Then
                        Exit Do
                    ElseIf RestartEditFlag Then
                        NextAction = taGet_Input: processphase = 1
                    Else
                        NextAction = taStay_In_Form
                    End If
                Else: processphase = processphase + 1: NextAction = taGet_Input
                End If
            Case taEscape_Form: Exit Do
            Case taStay_In_Form
                ShowFields (processphase)
                NextAction = taExit_Form
                If Not RestartEditFlag Then CancelCommunicationFlag = True
                Exit Do
        End Select
    End If
    On Error GoTo AfterActionError
    If AfterActionFlag And (Not docPrinting) Then ValidationControl.Run "AfterAction_Script"
Loop

Dim i
If CurrAction = taExit_Form Then
'    If QTrn <> 0 Then
'        For i = 1 To Fields.count
'            If Fields(i).QFldNo <> 0 Then
'                TRNFldNoQueue.Add Item:=CStr(Fields(i).QFldNo)
'                TRNFldTextQueue.Add Item:=Fields(i).Text
'            End If
'        Next i
'        LoadNewTransaction CStr(QTrn)
'        GoTo StartProcessLoop
'    Else
        Unload Me
'    End If
ElseIf CurrAction = taEscape_Form Then
    Unload Me
End If

Exit_Point:
Exit Sub
BeforeActionError:
Call NBG_LOG_MsgBox("Before Action Error :" & vbCrLf & ValidationControl.error.description, True)
GoTo Exit_Point
AfterActionError:
Call NBG_LOG_MsgBox("After Action Error :" & vbCrLf & ValidationControl.error.description, True)
GoTo Exit_Point
GenericError:
Call NBG_LOG_MsgBox("Error :" & vbCrLf & Err.description, True)
GoTo Exit_Point

End Sub

Private Sub ClearControls()
Dim i As Integer

ValidationControl.Reset
'For i = Labels.Count To 1 Step -1
'    Labels.Remove (i)
'Next i
For i = fields.count To 1 Step -1
    Controls.Remove fields(i).LabelName
    Controls.Remove fields(i).FldName
    fields.Remove (i)
Next i
For i = Labels.count To 1 Step -1
    Controls.Remove Labels(i).LabelName
    Labels.Remove (i)
Next i
For i = Lists.count To 1 Step -1
    Controls.Remove Lists(i).LabelName
    Controls.Remove Lists(i).LstName
    Lists.Remove (i)
Next i
For i = Spreads.count To 1 Step -1
    Controls.Remove Spreads(i).LabelName
    Controls.Remove Spreads(i).SprdName
    Spreads.Remove (i)
Next i

End Sub

Private Sub LoadNewTransaction(inTrnNo As String)


'Dim hwnd As Long

' Get a handle to the main application window.
'hwnd = FindWindow("TRNFrm", 0&)

' Lock the main applicationwindow.
' Prevents PowerPoint from redrawing the main window.
LockWindowUpdate (Me.hwnd)


cTRNCode = inTrnNo
SelectedTRN = CInt(inTrnNo)
ClearControls

InitForm
If CloseTransactionFlag Then GoTo before_exit_sub
PrepareFromXML
If CloseTransactionFlag Then GoTo before_exit_sub

On Error GoTo StartUpExecError
If StartupCodeFlag Then
    ValidationControl.Run "StartUp_Script"
End If

On Error GoTo GenErrorHandling
NextAction = taGet_Input
processphase = 1

'If Fields.count + Lists.count + Spreads.count = 0 Then
'    UCImage.Visible = True
'End If
'ProcessLoop

'Στοιχεία Πεδίων απο προηγούμενη συναλλαγή
Dim i As Integer, k As Integer
If TRNFldNoQueue.count > 0 Then
    For i = 1 To TRNFldNoQueue.count
        For k = 1 To fields.count
            If CStr(fields(k).FldNo) = TRNFldNoQueue(i) Then
                fields(k).Text = TRNFldTextQueue(i): Exit For
            End If
        Next k
    Next i
    For i = TRNFldNoQueue.count To 1 Step -1
        TRNFldNoQueue.Remove (i)
    Next i
    For i = TRNFldTextQueue.count To 1 Step -1
        TRNFldTextQueue.Remove (i)
    Next i
End If


' Unlocks the main application window. Whenever you use the
' LockWindowUpdate API to lock a window, you must call
' LockWindowUpdate again (with a Null parameter 0&) to allow the
' window to redraw.
LockWindowUpdate (0&)

before_exit_sub:
If CloseTransactionFlag Then
    Unload Me
End If

ExitPoint:
Exit Sub
GenErrorHandling:
Call NBG_LOG_MsgBox("Error :" & Err.description, True)
GoTo ExitPoint
StartUpExecError:
Call NBG_LOG_MsgBox("Startup Code run Error :" & vbCrLf & ValidationControl.error.description, True)
GoTo ExitPoint
HelpLoadError:

End Sub

Public Sub WriteJournalBeforeSend()
    WriteJournalBeforeSend_ processphase
End Sub

Private Function WriteJournalBeforeSend_(inPhase As Integer) As Boolean
'καταγραφή στο ημερολόγιο των πεδίων που θα σταλούν στο buffer
'επιστρέφει TRUE αν η καταγραφή ολοκληρωθεί χωρίς πρόβλημα
WriteJournalBeforeSend_ = False
Dim res As Boolean, FailedCount As Integer, SuccedCount As Integer, i As Integer

On Error GoTo GenericError
FailedCount = 0: SuccedCount = 0
For i = 1 To fields.count
    If fields(i).IsJournalBeforeOut(inPhase) Then
        If fields(i).WriteEJournal(inPhase, TrnCode(inPhase)) Then SuccedCount = SuccedCount + 1 _
        Else FailedCount = FailedCount + 1
    End If
Next i
For i = 1 To Combos.count
    If Combos(i).IsJournalBeforeOut(inPhase) Then
        If Combos(i).WriteEJournal(inPhase, TrnCode(inPhase)) Then SuccedCount = SuccedCount + 1 _
        Else FailedCount = FailedCount + 1
    End If
Next i


If FailedCount = 0 Then WriteJournalBeforeSend_ = True
SaveJournal

Exit_Point:
    Exit Function
GenericError:
    Call NBG_LOG_MsgBox("Error :" & vbCrLf & Err.code & Err.description, True)
WriteJournalBeforeSend_ = False
    GoTo Exit_Point
End Function

Private Function PrepareSendOut_(inPhase As Integer) As Boolean
'προετοιμασία του buffer για αποστολή στο ΚΜ
'επιστρέφει TRUE αν η διαδικασία (μέ την καταγραφή στο ημερολόγιο)
'ολοκληρωθεί χωρίς πρόβλημα
On Error GoTo GenericError
PrepareSendOut_ = False
sbWriteStatusMessage ""

'Clear Receive Data Buffers
Dim i As Integer
If ReceivedData.count > 0 Then For i = ReceivedData.count To 1 Step -1: ReceivedData.Remove i: Next i
If G0Data.count > 0 Then For i = G0Data.count To 1 Step -1: G0Data.Remove i: Next i


If Not Validate(inPhase) Then Exit Function

On Error GoTo BeforeOutError
BeforeOutFailed = False: BeforeOutError = ""

If BeforeOutFlag Then If Not ValidationControl.Run("BeforeOut_Script") Then Exit Function
On Error GoTo GenericError

Me.Enabled = False

Dim outString As String, CodePart As String
Dim aPos As String, foundflag As Boolean, afld As Object

For i = OutFields.count To 1 Step -1: OutFields.Remove (i): Next i

'Sort στα πεδία με τη σειρά που πρέπει να μπούν στο buffer
On Error GoTo 0
For i = 1 To fields.count
    If fields(i).GetOutBuffPos(inPhase) > 0 And fields(i).GetOutBuffLength(inPhase) > 0 Then
        aPos = CStr(fields(i).GetOutBuffPos(inPhase))
        aPos = StrPad_(aPos, 4, "0", "L")
        OutFields.add item:=fields(i), key:=aPos
    End If
Next i
For i = 1 To Combos.count
    If Combos(i).GetOutBuffPos(inPhase) > 0 And Combos(i).GetOutBuffLength(inPhase) > 0 Then
        aPos = CStr(Combos(i).GetOutBuffPos(inPhase))
        aPos = StrPad_(aPos, 4, "0", "L")
        OutFields.add item:=Combos(i), key:=aPos
    End If
Next i
On Error GoTo GenericError

foundflag = True
While foundflag
    foundflag = False
    For i = 1 To OutFields.count - 1
        If OutFields(i).GetOutBuffPos(inPhase) > OutFields(i + 1).GetOutBuffPos(inPhase) Then
            Set afld = OutFields(i + 1)
            OutFields.Remove (i + 1)
            OutFields.add item:=afld, Before:=i
            foundflag = True
        End If
    Next i
Wend
'Τέλος στο Sort

'Δημιουργία του string για το ΚΜ
Dim DataLength As Integer
Dim DataPart  As String
outString = ""

On Error GoTo PrepareErrorPos1
For i = 1 To OutFields.count
    OutFields(i).FormatBeforeOut
    If OutFields(i).IsOptional(inPhase) And OutFields(i).OutText = "" Then

    Else
        If OutFields(i).GetOutCodeEx(inPhase) <> "" Then
            CodePart = Right("00" & OutFields(i).GetOutCodeEx(inPhase), 2)
        ElseIf OutFields(i).GetOutCode(inPhase) > 0 Then
            CodePart = StrPad_(CStr(OutFields(i).GetOutCode(inPhase)), 2, "0", "L")
        Else
            CodePart = ""
        End If
        
        If FldLengthInBuffer Then
            CodePart = CodePart & StrPad_(Len(Trim(OutFields(i).OutText)), 4, "0", "L")
            CodePart = CodePart & Trim(OutFields(i).OutText)
        Else
            DataLength = OutFields(i).GetOutBuffLength(inPhase) - Len(CodePart)
            
            If Len(OutFields(i).OutText) > DataLength Then
                DataPart = Left$(OutFields(i).OutText, DataLength)
                    'Left part για τον Λογαριασμό δανείου όταν κόβεται το δευτερο CD
            Else
                If OutFields(i).EditType = etTEXT _
                Or OutFields(i).EditType = etNONE Then
                    DataPart = StrPad_(OutFields(i).OutText, DataLength, " ", "R")
                Else
                    DataPart = StrPad_(OutFields(i).OutText, DataLength, "0", "L")
                End If
            End If
            CodePart = CodePart & DataPart
        End If
        outString = outString & CodePart
    End If
    
Next i
On Error GoTo GenericError

'If Not DisableWriteJournal And Not DisableTRNCounterUpdate Then
If Not DisableWriteJournal Then
    If Not DisableTRNCounterUpdate Then
        UpdateTrnNum
    End If
    'καταγραφή στο ημερολόγιο
    If Not WriteJournalBeforeSend_(inPhase) Then PrepareSendOut_ = False: Exit Function
End If

If DisableTRNCounterUpdate Then
    DisableTRNCounterUpdate = False
End If

If Trim(SpecialKey) <> "" Then trn_key = SpecialKey
'Ελεγχος κλειδιών
ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False

If Not SkipKeyChk Then
    If Not isChiefTeller And ((trn_key = cCHIEFKEY) Or (trn_key = cTELLERCHIEFKEY)) Then _
        Set SelKeyFrm.owner = Me: ChiefRequest = True
    If Not isManager And ((trn_key = cMANAGERKEY) Or (trn_key = cTELLERMANAGERKEY)) Then _
        Set SelKeyFrm.owner = Me: ManagerRequest = True
End If

'Εγκριση
If ChiefRequest Or ManagerRequest Then
    SelKeyFrm.Show vbModal, Me
    If Not KeyAccepted Then PrepareSendOut_ = False: Exit Function
    eJournalWrite "Εγκριση " & IIf(ChiefRequest, "Chief Teller απο:" & cCHIEFUserName, "Manager από:" & cMANAGERUserName)
    SaveJournal
Else
    If ((trn_key = cCHIEFKEY) Or (trn_key = cTELLERCHIEFKEY)) Then
        ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = Me: KeyWarning.Show vbModal, Me
        If Not KeyAccepted Then
            PrepareSendOut_ = False
            Exit Function
        Else
            If isChiefTeller Then UpdateChiefKey cUserName
        End If
    ElseIf ((trn_key = cMANAGERKEY) Or (trn_key = cTELLERMANAGERKEY)) Then
        ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = Me: KeyWarning.Show vbModal, Me
        If Not KeyAccepted Then
            PrepareSendOut_ = False
            Exit Function
        Else
            If isManager Then UpdateManagerKey cUserName
        End If
    End If
End If

'τελικό string για το buffer
SetOutBuffer_ outString, inPhase
'cb.send_str = StrPad_(CStr(TRNNum(inPhase)), 4, "0", "L") & cHEAD & cb.trn_key & _
'    StrPad_(CStr(cTRNNum), 3, "0", "L") & outString
'cb.send_str_length = Len(cb.send_str)
'cb.curr_transaction = StrPad_(CStr(SelectedTRN), 4, "0", "L")


Me.Enabled = True
PrepareSendOut_ = True

Exit_Point:
Exit Function

GenericError:
    Call NBG_LOG_MsgBox("Error :" & vbCrLf & Err.code & Err.description, True)
    PrepareSendOut_ = False
    GoTo Exit_Point
BeforeOutError:
    Call NBG_LOG_MsgBox("Before Action Error :" & vbCrLf & ValidationControl.error.description, True)
PrepareSendOut_ = False
    GoTo Exit_Point
PrepareErrorPos1:
    Call NBG_LOG_MsgBox("Error PrepareSendOut Field:" & CStr(i) & vbCrLf & Err.code & Err.description, True)
End Function

Private Sub SetOutBuffer_(outString As String, inPhase As Integer)
'Δημιουργεί το string που θα σταλεί από τα καθαρά data
    cb.send_str = StrPad_(CStr(TrnCode(inPhase)), 4, "0", "L") & cHEAD & trn_key & _
        StrPad_(CStr(cTRNNum), 3, "0", "L") & outString
    'cb.send_str_length = Len(cb.send_str)
    cb.curr_transaction = StrPad_(CStr(SelectedTRN), 4, "0", "L")
End Sub

Private Function SendOut_(inPhase As Integer) As Boolean
'Αποστολή buffer στο ΚΜ
'Επιστρέφει TRUE αν η επικοινωνία ολοκληρωθεί χωρίς λάθος
SendOut_ = False
sbWriteStatusMessage ""

Dim com_status As Integer

Screen.MousePointer = vbHourglass
Me.Enabled = False

'Clear Receive Data Buffers
Dim i As Integer
If ReceivedData.count > 0 Then For i = ReceivedData.count To 1 Step -1: ReceivedData.Remove i: Next i
If G0Data.count > 0 Then For i = G0Data.count To 1 Step -1: G0Data.Remove i: Next i

'Clear Reset Key Flag
ResetKey = False

If EncodeGreekflag Then cb.DecodeGreek = 1: cb.encodegreek = 1 _
Else cb.DecodeGreek = 0: cb.encodegreek = 0

cb.LUADirection = 1: cb.send_convert = 1

CommunicationStarted = True
Dim aFlag As Boolean, bflag As Boolean
aFlag = SendJournalWrite: bflag = ReceiveJournalWrite
If HideSendFromJournal Then SendJournalWrite = False
If HideReceiveFromJournal Then ReceiveJournalWrite = False
If DisableWriteJournal Then SendJournalWrite = False: ReceiveJournalWrite = False

StartTime = Time
StartTickCount = GetTickCount
cTRNTime = 0

com_status = SNAPool_Communicate(Module28PoolLink)

EndTime = Time
EndTickCount = GetTickCount

Me.Enabled = True
Screen.MousePointer = vbDefault
SendJournalWrite = aFlag: ReceiveJournalWrite = bflag

Dim acount As Integer
acount = ListData.count
acount = ReceivedData.count

If (com_status = COM_OK) And cb.BoolTransOk Then
    On Error GoTo 0
    TRNOk = True: SendOut_ = True
Else
    On Error GoTo 0
    If CommunicationErrorFlag Then ValidationControl.Run "CommunicationError_Script"
    SendOut_ = False
End If

CommunicationStarted = False
End Function

Private Function WriteJournalAfterReceive_(inPhase As Integer) As Boolean
'καταγραφή στο ημερολόγιο των πεδίων που έχουν παραληφθεί από το ΚΜ
'επιστρέφει TRUE αν η καταγραφή ολοκληρωθεί χωρίς πρόβλημα
WriteJournalAfterReceive_ = False
Dim i As Integer, res As Boolean
For i = 1 To fields.count
    If fields(i).IsJournalAfterIN(inPhase) Then _
        res = fields(i).WriteEJournal(inPhase, TrnCode(inPhase))
Next i
SaveJournal
WriteJournalAfterReceive_ = True
End Function

Public Property Get received_data() As String
    received_data = cb.received_data
End Property

Public Property Let received_data(value As String)
    cb.received_data = value
End Property

Private Sub ReadIn_(inPhase As Integer)
'μεταφέρει στα πεδία το περιεχόμενο του buffer που στάλθηκε από το ΚΜ
'και καταγράφει στο ημερολόγιο τα πεδία
'επιστρέφει TRUE αν η μεταφορά ολοκληρωθεί χωρίς πρόβλημα
If AfterInFlag Then ValidationControl.Run "AfterIn_Script"

Me.Enabled = False
Dim inString As String
Dim res As Boolean
inString = cb.received_data
Dim i As Integer, CodePart As String
If Not CancelCommunicationFlag Then
    For i = 1 To fields.count
        With fields(i)
            If .GetInBuffLength(inPhase) > 0 Then
                If Len(inString) >= .GetInBuffPos(inPhase) Then
                    .InText = RTrim(Mid(inString, .GetInBuffPos(inPhase), .GetInBuffLength(inPhase)))
                Else
                    .InText = ""
                End If
                .FormatAfterIn
            End If
        End With
    Next i
    
    For i = 1 To Lists.count: res = Lists(i).ReadLines(): Next i
    For i = 1 To Spreads.count: res = Spreads(i).ReadLines(): Next i
End If
If Not DisableWriteJournal Then res = WriteJournalAfterReceive_(inPhase)
Me.Enabled = True

End Sub

Private Sub BeforePrint_(inPhase As Integer)
'Εκτελεί τον κώδικα του BeforePrintScript αφού έχουν φορτωθεί τα πεδία από το buffer
'και πρίν την εκτύπωση του παραστατικού
'Δεν περιλαμβάνεται στο block κλειδώματος του εκτυπωτή
    On Error Resume Next
    If BeforePrintFlag Then ValidationControl.Run "BeforePrint_Script"
End Sub

Private Sub GetInputFromRTF()
Dim astr As String, bstr As String, i As Integer
    astr = GenWorkForm.TrnInputBox.Text
    bstr = "PH" & StrPad_(CStr(processphase), 3, "0", "L") & vbCrLf
    i = InStr(1, astr, bstr, vbBinaryCompare)
    If i > 0 Then astr = Right(astr, Len(astr) - i - Len(bstr) + 1)
    i = InStr(1, astr, bstr, vbBinaryCompare)
    If i > 0 Then astr = Left(astr, i - 1)
    
    For i = ListData.count To 1 Step -1: ListData.Remove (i): Next i
    i = InStr(1, astr, vbCrLf, vbBinaryCompare)
    While i > 0
        ListData.add Left(astr, i - 1)
        astr = Right(astr, Len(astr) - i - 1)
        i = InStr(1, astr, vbCrLf, vbBinaryCompare)
    Wend
    If ListData.count > 0 Then cb.received_data = ListData.item(ListData.count) Else cb.received_data = ""
    ReadIn_ (processphase)
    NextAction = taPrint_Document
    ProcessLoop
End Sub

Private Sub MNUItem_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MNUItem(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub1_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub1(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub2_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub2(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub3_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub3(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub4_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub4(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub5_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub5(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub6_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub6(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub7_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub7(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub8_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub8(Index).Tag, 0, 0, aFlag
End Sub

Private Sub MnuSub9_Click(Index As Integer)
    Dim aFlag As Boolean
    aFlag = False
    HandleEvent MnuSub9(Index).Tag, 0, 0, aFlag
End Sub

Private Sub RefreshCom_Timer()
    Status.Panels(2).Visible = GenWorkForm.vStatus.Panels(2).Visible
    Status.Panels(3).Visible = GenWorkForm.vStatus.Panels(3).Visible
End Sub

Public Function AsciiToEbcdic(inputStr As String) As String
    AsciiToEbcdic = AsciiToEbcdic_(inputStr)
End Function

Public Function EbcdicToAscii(inputStr As String) As String
    EbcdicToAscii = EbcdicToAscii_(inputStr)
End Function

Public Function DecimalToHPS(invalue As Double, Digits As Long, positive As Boolean) As String
    DecimalToHPS = DecimalToHPS_(invalue, Digits)
End Function

Public Function HPSToDecimal(invalue As String) As Double
    HPSToDecimal = HPSToDecimal_(invalue)
End Function

Public Function IntToHps(ByVal InputInt As Long) As String
    IntToHps = IntToHps_(InputInt)
End Function

Public Function HpsToInt(ByVal InputInt As String) As Integer
    HpsToInt = HpsToInt_(InputInt)
End Function

Public Function SmallToHps(ByVal InputInt As Long) As String
    SmallToHps = SmallToHps_(InputInt)
End Function

Public Function HPSSEND(inputStr As String) As Integer
    HPSSEND = HPSSEND_(Me, inputStr)
End Function

Public Function HPSRECEIVE() As Integer
    HPSRECEIVE = HPSRECEIVE_(Me)
End Function

Public Function GetInBuffer() As String
    GetInBuffer = cb.receive_str
End Function

Public Function SQLServerName() As String

End Function

Public Sub WriteRegistry(RootKey, key, Variable, value)
Dim aSPCPanel
    Set aSPCPanel = CreateObject("SPCPanelXControl.SPCPanelX")
    aSPCPanel.host = MachineName
    'aSPCPanel.Port = 999
    aSPCPanel.Port = cPrinterPort
    aSPCPanel.Connect
'    aSPCPanel.PrintText ("EX$REGWRITE" & RootKey & vbCrLf & Key & vbCrLf & Variable & vbCrLf & Value & vbCrLf)
    aSPCPanel.SendText "PRINTEX$REGWRITE" & RootKey & vbCrLf & key & vbCrLf & Variable & vbCrLf & value & vbCrLf, True
    'aSPCPanel.DISCONNECT
    'aSPCPanel = Nothing
End Sub

Public Function ReadOCR() As String
Dim aSPCPanel, Result As String
    Set aSPCPanel = CreateObject("SPCPanelXControl.SPCPanelX")
    If (cOCRREADERSERVER = "") Then
        aSPCPanel.host = cPRINTERSERVER
    Else
        aSPCPanel.host = cOCRREADERSERVER
    End If
    aSPCPanel.Port = cOCRPort
    'aSPCPanel.Port = 999
    aSPCPanel.Connect
'    aSPCPanel.PrintText ("EX$REGWRITE" & RootKey & vbCrLf & Key & vbCrLf & Variable & vbCrLf & Value & vbCrLf)
    
    ReadOCR = aSPCPanel.SendTextB("PRINTEX$READOCR" & vbCrLf, True)
    
    'aSPCPanel.DISCONNECT
    'aSPCPanel = Nothing
End Function


Public Function StructBuildFromFields_() As Boolean
Dim i As Long, OutText As String
StructBuildFromFields_ = False
On Error GoTo 0
For i = 1 To fields.count
    If Trim(fields(i).HPSOutStruct) <> "" And Trim(fields(i).HPSOutPart) <> "" Then
        fields(i).FormatBeforeOut
        TrnBuffers.SetPart fields(i).HPSOutStruct, fields(i).HPSOutPart, 1, fields(i).OutText
        
        
    End If
Next i
StructBuildFromFields_ = True
End Function

Public Function HPSPrepareForSend_(inPhase As Integer) As Boolean
'προετοιμασία του buffer για αποστολή στο ΚΜ
'επιστρέφει TRUE αν η διαδικασία (μέ την καταγραφή στο ημερολόγιο)
'ολοκληρωθεί χωρίς πρόβλημα
On Error GoTo GenericError
HPSPrepareForSend_ = False
sbWriteStatusMessage ""

'Clear Receive Data Buffers
Dim i As Integer
If ReceivedData.count > 0 Then For i = ReceivedData.count To 1 Step -1: ReceivedData.Remove i: Next i
If G0Data.count > 0 Then For i = G0Data.count To 1 Step -1: G0Data.Remove i: Next i

If Not Validate(inPhase) Then Exit Function

On Error GoTo BeforeOutError
BeforeOutFailed = False: BeforeOutError = ""

If BeforeOutFlag Then If Not ValidationControl.Run("BeforeOut_Script") Then Exit Function
On Error GoTo GenericError

Me.Enabled = False

Dim outString As String, CodePart As String
Dim aPos As String, foundflag As Boolean, afld As Object

'Δημιουργία του string για το ΚΜ
If Not StructBuildFromFields_ Then HPSPrepareForSend_ = False: Exit Function

UpdateTrnNum
'καταγραφή στο ημερολόγιο
If Not WriteJournalBeforeSend_(inPhase) Then HPSPrepareForSend_ = False: Exit Function

If Trim(SpecialKey) <> "" Then trn_key = SpecialKey
'Ελεγχος κλειδιών
ChiefRequest = False: SecretRequest = False: ManagerRequest = False: KeyAccepted = False

If Not SkipKeyChk Then
    If Not isChiefTeller And ((trn_key = cCHIEFKEY) Or (trn_key = cTELLERCHIEFKEY)) Then _
        Set SelKeyFrm.owner = Me: ChiefRequest = True
    If Not isManager And ((trn_key = cMANAGERKEY) Or (trn_key = cTELLERMANAGERKEY)) Then _
        Set SelKeyFrm.owner = Me: ManagerRequest = True
End If

'Εγκριση
If ChiefRequest Or ManagerRequest Then
    SelKeyFrm.Show vbModal, Me
    If Not KeyAccepted Then HPSPrepareForSend_ = False: Exit Function
    eJournalWrite "Εγκριση " & IIf(ChiefRequest, "Chief Teller απο:" & cCHIEFUserName, "Manager από:" & cMANAGERUserName)
    SaveJournal
Else
    If ((trn_key = cCHIEFKEY) Or (trn_key = cTELLERCHIEFKEY)) Then
        ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = Me: KeyWarning.Show vbModal, Me
        If Not KeyAccepted Then
            HPSPrepareForSend_ = False
            Exit Function
        Else
            If isChiefTeller Then UpdateChiefKey cUserName
        End If
        
    ElseIf ((trn_key = cMANAGERKEY) Or (trn_key = cTELLERMANAGERKEY)) Then
        ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = Me: KeyWarning.Show vbModal, Me
        If Not KeyAccepted Then
            HPSPrepareForSend_ = False
            Exit Function
        Else
            If isManager Then UpdateManagerKey cUserName
        End If
    End If
End If

'τελικό string για το buffer
SetOutBuffer_ outString, inPhase
'cb.send_str = StrPad_(CStr(TRNNum(inPhase)), 4, "0", "L") & cHEAD & cb.trn_key & _
'    StrPad_(CStr(cTRNNum), 3, "0", "L") & outString
'cb.send_str_length = Len(cb.send_str)
'cb.curr_transaction = StrPad_(CStr(SelectedTRN), 4, "0", "L")


Me.Enabled = True
HPSPrepareForSend_ = True

Exit_Point:
Exit Function

GenericError:
    Call NBG_LOG_MsgBox("Error :" & vbCrLf & Err.code & Err.description, True)
    HPSPrepareForSend_ = False
    GoTo Exit_Point
BeforeOutError:
    Call NBG_LOG_MsgBox("Before Action Error :" & vbCrLf & ValidationControl.error.description, True)
HPSPrepareForSend_ = False
    GoTo Exit_Point

End Function

Public Function StructDefine(aStructName As String, aStructDesc As String) As Integer
    StructDefine = TrnBuffers.DefineBuffer(aStructName, aStructName, aStructDesc)
End Function

Public Function StructClear(aStructName As String) As Integer
    'StructClear = ClearStruct_(aStructName)
End Function

Public Function StructGetIN(aStruct As String) As String
    StructGetIN = TrnBuffers.GetIn(aStruct)
End Function

Public Function StructSetIN(aStruct As String, aValue As String) As Integer
    TrnBuffers.SetIn aStruct, aValue
End Function

Public Function StructSetPart(aStruct As String, apart As String, aValue, Optional position) As Integer
    If IsMissing(position) Then position = 1
    TrnBuffers.SetPart aStruct, apart, CLng(position), aValue
End Function

Public Function StructGetINPart(aStruct As String, apart As String, Optional position)
    If IsMissing(position) Then position = 1
    StructGetINPart = TrnBuffers.GetInPart(aStruct, apart, CLng(position))
End Function

Public Function StructGetPart(aStruct As String, apart As String, Optional position)
    If IsMissing(position) Then position = 1
    StructGetPart = TrnBuffers.GetPart(aStruct, apart, CLng(position))
End Function


Public Sub debug_()
    Dim i As Integer
    i = 0
End Sub

Public Sub DimTRNVariable(VariableName, Optional value)
    If IsMissing(value) Then Set value = Nothing
    ValidationControl.ExecuteStatement "DIM " & VariableName
End Sub

Public Sub BuildIRISStruct(StructureName, Optional StructureAlias, Optional hidden As Boolean)
    
'    If Not cIRISConnected Then
'        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (BuildIRISStruct)", vbCritical: Exit Sub
'    End If
On Error GoTo GenError
    
    If IsMissing(StructureAlias) Then StructureAlias = StructureName
    If IsMissing(hidden) Then hidden = False
    
    If Trim(StructureAlias) = "" Then StructureAlias = StructureName
    
    Dim aDesc As String, bDesc As String, res As Long, AliasCopy As String, astructid As String
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
    
    Dim aPos As Integer, neststruct As String
    aDesc = UCase(aDesc): bDesc = aDesc
    aPos = InStr(1, aDesc, "STRUCT ")
    While aPos > 0
        aDesc = Right(aDesc, Len(aDesc) - aPos + 1)
        aDesc = Right(aDesc, Len(aDesc) - 7)
        aPos = InStr(1, aDesc, " ")
        neststruct = Trim(Left(aDesc, aPos))
        
        If Not TrnBuffers.Exists(neststruct) Then BuildIRISStruct neststruct, neststruct, True
        
        aPos = InStr(1, aDesc, "STRUCT ")
    Wend
    
    res = TrnBuffers.DefineBuffer(CStr(StructureAlias), CStr(astructid), bDesc, CStr(StructureName), Not hidden)
    If res > -1 Then
        If IsNumeric(Left(StructureAlias, 1)) Then AliasCopy = "_" & StructureAlias Else AliasCopy = StructureAlias
        ValidationControl.ExecuteStatement "DIM " & AliasCopy
        ValidationControl.ExecuteStatement "set " & AliasCopy & "= Trnbuffers.ByName(""" & StructureAlias & """)"
    End If
    Exit Sub
GenError:
    MsgBox "Πρόβλημα στη δήλωση της δομής: " & StructureAlias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "ΛΑΘΟΣ"
End Sub

Public Sub BuildIRISAppStruct(StructureName, Optional StructureAlias, Optional hidden As Boolean)
    If IsMissing(hidden) Then hidden = False
    DatabaseMdl.BuildIRISAppStruct StructureName, StructureAlias, Not hidden
End Sub

Public Sub BuildStructFromDB(tablename, StructureName, Optional StructureAlias, Optional hidden As Boolean)
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (BuildStructFromDB)", vbCritical: Exit Sub
    End If
On Error GoTo GenError
    
    If IsMissing(StructureAlias) Then StructureAlias = StructureName
    If IsMissing(hidden) Then hidden = False
    
    If Trim(StructureAlias) = "" Then StructureAlias = StructureName
    
    Dim aDesc As String, bDesc As String, res As Long, AliasCopy As String
    aDesc = xmlCRAStructures.documentElement.selectSingleNode(StructureName).Text
    
    Dim aPos As Integer, neststruct As String
    aDesc = UCase(aDesc): bDesc = aDesc
    aPos = InStr(1, aDesc, "STRUCT ")
    While aPos > 0
        aDesc = Right(aDesc, Len(aDesc) - aPos + 1)
        aDesc = Right(aDesc, Len(aDesc) - 7)
        aPos = InStr(1, aDesc, " ")
        neststruct = Trim(Left(aDesc, aPos))
        
        If Not TrnBuffers.Exists(neststruct) Then BuildStructFromDB tablename, neststruct, , True
        
        aPos = InStr(1, aDesc, "STRUCT ")
    Wend
    
    res = TrnBuffers.DefineBuffer(CStr(StructureAlias), CStr(StructureAlias), bDesc, CStr(StructureName), Not hidden)
    If res > -1 Then
        If IsNumeric(Left(StructureAlias, 1)) Then AliasCopy = "_" & StructureAlias Else AliasCopy = StructureAlias
        ValidationControl.ExecuteStatement "DIM " & AliasCopy
        ValidationControl.ExecuteStatement "set " & AliasCopy & "= Trnbuffers.ByName(""" & StructureAlias & """)"
    End If
    Exit Sub
GenError:
    MsgBox "Πρόβλημα στη δήλωση της δομής: " & StructureAlias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "ΛΑΘΟΣ"
End Sub

Public Function BuildAppStructFromDB(tablename, StructureName, Alias, Optional hidden As Boolean) As Boolean
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (BuildAppStructFromDB)", vbCritical: Exit Function
    End If
    
    BuildAppStructFromDB = False
    On Error GoTo GenError
    If GenWorkForm.AppBuffers.Exists(Alias) Then Exit Function
    If IsMissing(hidden) Then hidden = False
    
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
        
        If Not GenWorkForm.AppBuffers.Exists(neststruct) Then BuildAppStructFromDB tablename, neststruct, neststruct, True
        
        aPos = InStr(1, aDesc, "STRUCT ")
    Wend
    
    GenWorkForm.AppBuffers.DefineBuffer CStr(Alias), CStr(Alias), bDesc, CStr(StructureName), Not hidden
    BuildAppStructFromDB = True
    Exit Function
GenError:
    MsgBox "Πρόβλημα στη δήλωση της δομής: " & Alias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "ΛΑΘΟΣ"
End Function

Public Sub FreeAppStruct(StructureName)
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (FreeAppStruct)", vbCritical: Exit Sub
    End If
    
    GenWorkForm.AppBuffers.FreeBuffer CStr(StructureName)
End Sub


Public Sub HandleEvent(SenderName, code, State, ByRef UnloadFlag)
    On Error GoTo ScriptError
'    If OpAfterHandling = fopExitForm Then Exit Sub
    OpAfterHandling = fopNoOperation
    Me.Enabled = False
    If AfterKeyFlag Then ValidationControl.Run "AfterKey_Script", UCase(SenderName), code, State
    Me.Enabled = True
    
'    If OpAfterHandling = fopExitForm Then Unload Me Else OpAfterHandling = fopNoOperation
    If OpAfterHandling = fopExitForm Then Unload Me: UnloadFlag = True
    If OpAfterHandling = fopSendBuffer Then
        If AutoExitFlag Then
            UnloadFlag = True
        End If
        NextAction = taSend_Buffer: ProcessLoop
    End If
    
    Exit Sub
ScriptError:
Call NBG_LOG_MsgBox("Error :" & CStr(Err.number) & Err.description & "-" & CStr(ValidationControl.error.number) & ValidationControl.error.description, True)
    
End Sub

Public Sub LinkToTrn(inTrnCD, AppLevelOutBuffer, AppLevelInBuffer)
Dim oldTrnCD As String, anewTrnFrm As New TRNFrm, oldFlag As Boolean
    
    oldTrnCD = cTRNCode: cTRNCode = inTrnCD: oldFlag = cEnableHiddenTransactions: cEnableHiddenTransactions = True
    With anewTrnFrm
        .AppLevelOutBuffer = AppLevelOutBuffer
        .AppLevelInBuffer = AppLevelInBuffer
        .Show vbModal, Me
    End With
    cTRNCode = oldTrnCD: cEnableHiddenTransactions = oldFlag
    Set anewTrnFrm = Nothing
End Sub

Public Sub LinkToTrnV2(inTrnCD, Params)
Dim oldTrnCD As String, anewTrnFrm As New TRNFrm, oldFlag As Boolean
    oldTrnCD = cTRNCode: cTRNCode = inTrnCD: oldFlag = cEnableHiddenTransactions: cEnableHiddenTransactions = True
    With anewTrnFrm
        Set .OwnerForm = Me
        .Params = Params
        On Error Resume Next: .Show vbModal, Me
    End With
    cTRNCode = oldTrnCD: cEnableHiddenTransactions = oldFlag
    Set anewTrnFrm = Nothing
End Sub

Public Sub LinkToTrnV3(inTrnCD, PArray)
Dim oldTrnCD As String, anewTrnFrm As New TRNFrm, oldFlag As Boolean
    oldTrnCD = cTRNCode: cTRNCode = inTrnCD: oldFlag = cEnableHiddenTransactions: cEnableHiddenTransactions = True
    Dim astr As String
    astr = TypeName(PArray)
    With anewTrnFrm
        Set .OwnerForm = Me
        .PArray = PArray
        On Error Resume Next: .Show vbModal, Me
    End With
    cTRNCode = oldTrnCD: cEnableHiddenTransactions = oldFlag
    Set anewTrnFrm = Nothing
End Sub

Public Property Get param(aname As String)
Dim i As Integer, k As Long, l As Long, astr As String
    If paramnames <> "" Then
        i = InStr(UCase(Trim(paramnames)), UCase(Trim(aname)))
        If i > 0 Then
            If i = 1 Then
                i = InStr(Params, ",")
                If i = 0 Then param = UCase(Trim(Params)) Else param = UCase(Trim(Left(Params, i - 1)))
                Exit Property
            Else
                astr = Left(paramnames, i - 1)
                k = 0: l = InStr(astr, ",")
                While l > 0
                    k = k + 1: l = InStr(l + 1, astr, ",")
                Wend
                l = 0
                For i = 1 To k
                    l = InStr(l + 1, Params, ",")
                    If l = 0 Then param = "": Exit Property
                Next i
                k = InStr(l + 1, Params, ",")
                If k = 0 Then
                    param = UCase(Trim(Right(Params, Len(Params) - l))): Exit Property
                Else
                    param = UCase(Trim(Mid(Params, l + 1, k - l - 1))): Exit Property
                End If
            End If
        Else
            param = ""
        End If
    Else
        On Error Resume Next
        If Not (PNameArray = Empty) Then
        For i = LBound(PNameArray) To UBound(PNameArray)
            If UCase(aname) = UCase(PNameArray(i)) Then
                If UCase(TypeName(PArray(i))) = UCase("Object") Then
                    Set param = PArray(i): Exit For
                Else
                    param = PArray(i): Exit For
                End If
            End If
        Next i
        End If
    End If
End Property

Public Property Let param(aname As String, value)
Dim i As Integer, k As Long, l As Long, astr As String
    On Error Resume Next
    If Not (PNameArray = Empty) Then
        For i = LBound(PNameArray) To UBound(PNameArray)
            If UCase(aname) = UCase(PNameArray(i)) Then
                If UCase(TypeName(value)) = UCase("Object") Then
                    Set PArray(i) = value: Exit For
                Else
                    PArray(i) = value: Exit For
                End If
            End If
        Next i
    End If
End Property

Public Function GetCRASetValue(SetName As String, Optional SetCode)
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (GetCRASetValue)", vbCritical: Exit Function
    End If
    
    If UCase(SetName) = "NAMEFORMAT" Then 'απο 7759
        If SetCode >= 1 And SetCode <= 5 Then
            GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            Select Case SetCode
            Case 1
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΠΩΝΥΜΟ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΟΝΟΜΑ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΟΝΟΜΑ ΠΑΤΕΡΑ": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΟΝΟΜΑ ΜΗΤΕΡΑΣ": .ByName("Flag", 4).value = 1
                    .ByName("LBL1", 5).value = "ΟΝΟΜΑΤΕΠΩΝΥΜΟ ΣΥΖΥΓΟΥ": .ByName("Flag", 5).value = 1
                    .ByName("LBL1", 6).value = "ΓΕΝΟΣ": .ByName("Flag", 6).value = 1
                    .ByName("LBL1", 7).value = "ΚΥΡΙΟΣ/ΚΥΡΙΑ": .ByName("Flag", 7).value = 1
                    .ByName("LBL1", 8).value = "ΠΡΟΣΦΩΝΗΣΗ": .ByName("Flag", 8).value = 1
                End With
            Case 2
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΑΤΟΜΙΚΗ ΕΠΙΧΕΙΡΗΣΗ": .ByName("Flag", 1).value = 1
                End With
            Case 3
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΠΩΝΥΜΙΑ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΕΠΩΝΥΜΙΑ 2": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΕΠΩΝΥΜΙΑ 3": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΝΟΜΙΚΗ ΜΟΡΦΗ": .ByName("Flag", 4).value = 1
                End With
            Case 4
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ 2": .ByName("Flag", 2).value = 1
                End With
            Case 5
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΠΩΝΥΜΙΑ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΕΠΩΝΥΜΙΑ 2": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΕΠΩΝΥΜΙΑ 3": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΕΙΔΟΣ ΜΟΝΑΔΑΣ": .ByName("Flag", 4).value = 1
                    .ByName("LBL1", 5).value = "ΤΟΠΟΣ": .ByName("Flag", 5).value = 1
                End With
            End Select
            Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
        End If
    ElseIf UCase(SetName) = "NAMEFORMAT2" Then ' απο 7735
            GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            Select Case SetCode
            Case 900002
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΠΩΝΥΜΟ-ONOMA-ΕΠΩΝΥΜΙΑ": .ByName("Flag", 1).value = 1
                End With
            Case 900004
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ 2": .ByName("Flag", 2).value = 1
                End With
            Case 900006
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΑΤΟΜΙΚΗ ΕΠΙΧΕΙΡΗΣΗ": .ByName("Flag", 1).value = 1
                End With
            End Select
            Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "NAMEPREFIX" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
        With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
            .ByName("LBL1", 1).value = "ΚΑΜΙΑ": .ByName("Flag", 1).value = 1
            .ByName("LBL1", 2).value = "ΓΙΑΤΡΟΣ": .ByName("Flag", 2).value = 1
            .ByName("LBL1", 3).value = "ΚΑΘΗΓΗΤΗΣ": .ByName("Flag", 3).value = 1
            .ByName("LBL1", 4).value = "ΒΟΥΛΕΥΤΗΣ": .ByName("Flag", 4).value = 1
            .ByName("LBL1", 5).value = "ΝΑΥΑΡΧΟΣ": .ByName("Flag", 5).value = 1
            .ByName("LBL1", 6).value = "ΑΡΧΙΕΠΙΣΚΟΠΟΣ": .ByName("Flag", 6).value = 1
            .ByName("LBL1", 7).value = "ΣΤΡΑΤΗΓΟΣ": .ByName("Flag", 7).value = 1
            .ByName("LBL1", 8).value = "ΑΙΔΕΣΙΜΟΤΑΤΟΣ": .ByName("Flag", 8).value = 1
            .ByName("LBL1", 9).value = "ΠΑΤΕΡ": .ByName("Flag", 9).value = 1
            .ByName("LBL1", 10).value = "ΠΤΕΡΑΡΧΟΣ": .ByName("Flag", 10).value = 1
            .ByName("LBL1", 11).value = "ΥΠΟΥΡΓΟΣ": .ByName("Flag", 11).value = 1
            .ByName("LBL1", 12).value = "ΠΡΟΕΔΡΟΣ": .ByName("Flag", 12).value = 1
        End With
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "CUSTOMERTYPE1" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
        With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
            .ByName("LBL1", 1).value = "ΦΥΣΙΚΟ ΠΡΟΣΩΠΟ: ΟΝΟΜΑΤΕΠΩΝΥΜΟ": .ByName("Flag", 1).value = 1
            .ByName("LBL1", 2).value = "ΦΥΣΙΚΟ ΠΡΟΣΩΠΟ: ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ (ΑΤΟΜΙΚΗ ΕΠΙΧ/ΣΗ)": .ByName("Flag", 2).value = 1
            .ByName("LBL1", 3).value = "ΝΟΜΙΚΟ ΠΡΟΣΩΠΟ: ΕΠΩΝΥΜΙΑ": .ByName("Flag", 3).value = 1
            .ByName("LBL1", 4).value = "ΝΟΜΙΚΟ ΠΡΟΣΩΠΟ: ΔΙΑΚΡΙΤΙΚΟΣ ΤΙΤΛΟΣ": .ByName("Flag", 4).value = 1
            .ByName("LBL1", 5).value = "ΜΟΝΑΔΑ ΝΟΜΙΚΟΥ ΠΡΟΣΩΠΟΥ: ΕΠΩΝΥΜΙΑ": .ByName("Flag", 5).value = 1
        End With
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "CUSTOMERTYPE2" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
        With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
            .ByName("LBL1", 1).value = "ΦΥΣΙΚΟ ΠΡΟΣΩΠΟ": .ByName("Flag", 1).value = 1
            .ByName("LBL1", 2).value = "ΝΟΜΙΚΟ ΠΡΟΣΩΠΟ": .ByName("Flag", 2).value = 1
            .ByName("LBL1", 3).value = "ΜΟΝΑΔΑ ΝΟΜΙΚΟΥ ΠΡΟΣΩΠΟΥ": .ByName("Flag", 3).value = 1
        End With
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "ADDRESSTYPE" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
        With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
            .ByName("LBL1", 1).value = "ΔΙΕΥΘΥΝΣΗ ΕΞΩΤΕΡΙΚΟΥ": .ByName("Flag", 1).value = 1
            .ByName("LBL1", 2).value = "ΤΑΧΥΔΡΟΜΙΚΗ ΘΥΡΙΔΑ": .ByName("Flag", 2).value = 1
            .ByName("LBL1", 3).value = "ΤΑΧΥΔΡΟΜΙΚΗ ΔΙΕΥΘΥΝΣΗ": .ByName("Flag", 3).value = 1
            .ByName("LBL1", 4).value = "ΤΑΧΥΔΡΟΜΙΚΟ ΓΡΑΦΕΙΟ": .ByName("Flag", 4).value = 1
        End With
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "ADDRESSTYPE2" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            Select Case SetCode
            Case 1, 2 'ΔΙΕΥΘΥΝΣΗ ΚΑΤΟΙΚΙΑΣ 'ΔΙΕΥΘΥΝΣΗ ΕΡΓΑΣΙΑΣ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΔΙΕΥΘΥΝΣΗ ΕΞΩΤΕΡΙΚΟΥ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΤΑΧΥΔΡΟΜΙΚΗ ΔΙΕΥΘΥΝΣΗ": .ByName("Flag", 2).value = 1
                End With
            Case 3 'ΔΙΕΥΘΥΝΣΗ ΕΠΙΚΟΙΝΩΝΙΑΣ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΔΙΕΥΘΥΝΣΗ ΕΞΩΤΕΡΙΚΟΥ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΤΑΧΥΔΡΟΜΙΚΗ ΘΥΡΙΔΑ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΤΑΧΥΔΡΟΜΙΚΗ ΔΙΕΥΘΥΝΣΗ": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΤΑΧΥΔΡΟΜΙΚΟ ΓΡΑΦΕΙΟ": .ByName("Flag", 4).value = 1
                End With
            End Select
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "ADDRESSFORMAT" Then
        If SetCode >= 1 And SetCode <= 4 Then
            GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            Select Case SetCode
            Case 1 'ΔΙΕΥΘΥΝΣΗ ΕΞΩΤΕΡΙΚΟΥ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΓΡΑΜΜΗ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΓΡΑΜΜΗ 2": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΓΡΑΜΜΗ 3": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΓΡΑΜΜΗ 4": .ByName("Flag", 4).value = 1
                    .ByName("LBL1", 5).value = "ΧΩΡΑ": .ByName("Flag", 5).value = 1
                End With
            Case 2 'ΤΑΧΥΔΡΟΜΙΚΗ ΘΥΡΙΔΑ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΤΑΧ. ΘΥΡΙΔΑ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΤΑΧ. ΚΩΔΙΚΑΣ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΤΑΧ. ΠΕΡΙΟΧΗ": .ByName("Flag", 3).value = 1
                End With
            Case 3 'ΤΑΧΥΔΡΟΜΙΚΗ ΔΙΕΥΘΥΝΣΗ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΙΔΟΣ ΔΡΟΜΟΥ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΟΝΟΜΑΣΙΑ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΠΕΡΙΟΧΗ": .ByName("Flag", 4).value = 1
                    .ByName("LBL1", 5).value = "ΤΑΧ. ΚΩΔΙΚΑΣ": .ByName("Flag", 5).value = 1
                    .ByName("LBL1", 6).value = "ΤΑΧ. ΠΕΡΙΟΧΗ": .ByName("Flag", 6).value = 1
                End With
            Case 4 'ΤΑΧΥΔΡΟΜΙΚΟ ΓΡΑΦΕΙΟ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "POSTE RESTANTE": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΤΑΧ. ΚΩΔΙΚΑΣ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΤΑΧ. ΠΕΡΙΟΧΗ": .ByName("Flag", 3).value = 1
                End With
            End Select
            Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
        End If
    ElseIf UCase(SetName) = "ADDRESSFORMAT2" Then
'        If SetCode >= 1 And SetCode <= 4 Or (SetCode = 11 Or SetCode = 12) Then
            GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            Select Case SetCode
            Case "ΔΝΣΕΞ" 'ΔΙΕΥΘΥΝΣΗ ΕΞΩΤΕΡΙΚΟΥ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΓΡΑΜΜΗ 1": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΓΡΑΜΜΗ 2": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΓΡΑΜΜΗ 3": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΓΡΑΜΜΗ 4": .ByName("Flag", 4).value = 1
                End With
            Case "ΤΧΘΥΡ" 'ΤΑΧΥΔΡΟΜΙΚΗ ΘΥΡΙΔΑ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΤΑΧ. ΘΥΡΙΔΑ": .ByName("Flag", 1).value = 1
                End With
            Case "ΤΧΔΝΣ" 'ΤΑΧΥΔΡΟΜΙΚΗ ΔΙΕΥΘΥΝΣΗ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΕΙΔΟΣ ΔΡΟΜΟΥ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΟΝΟΜΑΣΙΑ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 3).value = 1
                    .ByName("LBL1", 4).value = "ΠΕΡΙΟΧΗ": .ByName("Flag", 4).value = 1
                End With
            Case "ΤΧΓΡΦ" 'ΤΑΧΥΔΡΟΜΙΚΟ ΓΡΑΦΕΙΟ
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "POSTE RESTANTE": .ByName("Flag", 1).value = 1
                End With
            End Select
            Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
'        End If
    ElseIf UCase(SetName) = "COMMUNICATIONNUMBERTYPE" Then
        GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
        With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
            .ByName("LBL1", 1).value = "SWIFT": .ByName("Flag", 1).value = 1
            .ByName("LBL1", 2).value = "ΤΗΛΕΦΩΝΟ": .ByName("Flag", 2).value = 1
            .ByName("LBL1", 3).value = "TELEX": .ByName("Flag", 3).value = 1
            .ByName("LBL1", 4).value = "FAX": .ByName("Flag", 4).value = 1
            .ByName("LBL1", 5).value = "INTERNET": .ByName("Flag", 5).value = 1
            .ByName("LBL1", 6).value = "E-MAIL": .ByName("Flag", 6).value = 1
            .ByName("LBL1", 7).value = "ΤΗΛΕΦΩΝΟ ΕΞΩΤΕΡΙΚΟΥ": .ByName("Flag", 7).value = 1
        End With
        Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
    ElseIf UCase(SetName) = "COMMUNICATIONNUMBERFORMAT" Then
        If IsNumeric(SetCode) Then
            If SetCode >= 1 And SetCode <= 7 Then
                GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
                Select Case SetCode
                Case 1, 3, 4, 5, 6, 7
                    With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                        .ByName("LBL1", 1).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 1).value = 1
                    End With
                Case 2
                    With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                        .ByName("LBL1", 1).value = "ΚΩΔ. ΠΕΡΙΟΧΗΣ": .ByName("Flag", 1).value = 1
                        .ByName("LBL1", 2).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 2).value = 1
                        .ByName("LBL1", 3).value = "ΕΣΩΤΕΡΙΚΟ": .ByName("Flag", 3).value = 1
                    End With
                End Select
                Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
            End If
        Else
            GenWorkForm.AppBuffers.ByName("NameFormat").ClearData
            If SetCode = "ΤΗΛΕΦ" Then
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΚΩΔ. ΠΕΡΙΟΧΗΣ": .ByName("Flag", 1).value = 1
                    .ByName("LBL1", 2).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 2).value = 1
                    .ByName("LBL1", 3).value = "ΕΣΩΤΕΡΙΚΟ": .ByName("Flag", 3).value = 1
                End With
            Else
                With GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine")
                    .ByName("LBL1", 1).value = "ΑΡΙΘΜΟΣ": .ByName("Flag", 1).value = 1
                End With
            End If
            Set GetCRASetValue = GenWorkForm.AppBuffers.ByName("NameFormat").ByName("NameFormatLine"): Exit Function
        End If
    End If
End Function



Public Function ChkHPSComResult(inRslt, inErrors) As Integer
    ChkHPSComResult = ChkHPSComResult_(inRslt, inErrors)
    
End Function

Public Sub ChkHPSWarnings(inErrors, inStructName)
    Dim i As Integer, k As Integer
    k = 0
    For i = 1 To inErrors.ByName(inStructName).times
        If inErrors.ByName(inStructName).ByName("N_ERR", i).value > 0 Then k = k + 1: Exit For
    Next i
    If k = 0 Then Exit Sub
       
    Load HPSErrForm: Set HPSErrForm.ErrBuffer = inErrors: HPSErrForm.StructName = inStructName: HPSErrForm.Show vbModal, Me
End Sub
Public Sub cmdForceLogoff()
    cmdForceLogoff_
End Sub

Public Function IRISCom(Trn As String, rule As String, InputView, OutputView, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
    'IRISCom = IRISCom_(Me, Trn, rule, InputView, OutputView, AuthUser, Appltran, ErrorView, ErrorCount)
    
    DisableTRNCounterUpdate = True
    UpdateTrnNum_
    
    Dim atrnname As String
    atrnname = InputView.name
    If Len(atrnname) > 2 And Right(atrnname, 2) = "_I" Then atrnname = Left(atrnname, Len(atrnname) - 2)
    WriteJournal "ΔΙΑΔΙΚΑΣΙΑ: " & atrnname
        
    Dim aresult As cSNAResult
    sbWriteStatusMessage "ΔΙΑΒΙΒΑΣΗ ΣΤΟΙΧΕΙΩΝ..."
    Set aresult = iriscomnew_(Me, Left(Trn & "    ", 4), rule, InputView, OutputView, AuthUser, _
        Appltran, ErrorView, ErrorCount, False)
    IRISCom = aresult.ErrCode
    aresult.UpdateForm Me
    If (aresult.ErrCode <> 0 And aresult.ErrCode <> COM_OK) Then
        LogMsgbox "Λάθος: " & CStr(aresult.ErrCode) & " " & aresult.ErrMessage, vbCritical, "Πρόβλημα Επικοινωνίας...."
    End If
    
End Function

Public Function IRISComWithLog(Trn As String, rule As String, InputView, OutputView, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
    Dim aFlag As Boolean
    aFlag = LogIrisCom
    LogIrisCom = True
    IRISComWithLog = IRISCom_(Me, Trn, rule, InputView, OutputView, AuthUser, Appltran, ErrorView, ErrorCount)
    LogIrisCom = aFlag
End Function

Public Function GetFTFilaName(inTbl, inAppl, invalue) As String
    GetFTFilaName = GetFTFilaName_(CStr(inTbl), CStr(inAppl), CStr(invalue))
End Function

Public Function GetFTFilaDescription(inTbl, inAppl, invalue) As String
    GetFTFilaDescription = GetFTFilaDescription_(CStr(inTbl), CStr(inAppl), CStr(invalue))
End Function

Public Function ChkIRISOutput(aBuffer, Optional looseChk) As Boolean
    If IsMissing(looseChk) Then looseChk = False
    ChkIRISOutput = ChkIRISOutput_(aBuffer, looseChk)
End Function

Public Function ChkIRISOutputSkip7(aBuffer, Optional looseChk) As Boolean
    If IsMissing(looseChk) Then looseChk = False
    ChkIRISOutputSkip7 = ChkIRISOutputSkip7_(aBuffer, looseChk)
End Function

Public Function ChkCRA2Output(aBuffer, Optional looseChk) As Boolean
    If IsMissing(looseChk) Then looseChk = False
    ChkCRA2Output = ChkCRA2Output_(aBuffer, looseChk)
End Function

Public Sub PrepareCRA2IDView(aBuffer)
    With aBuffer
         .v2Value("I_ENTP") = 1
         .v2Value("C_ACOD_FI") = "001"
         .v2Value("C_ACOD_OU") = cBRANCH
         .v2Value("C_USR_ID") = UCase(cUserName)
        If cIRISUserName <> "" Then
            .v2Value("C_USR_ID") = UCase(cIRISUserName)
        End If
         
         .v2Value("C_WKST_ID") = Right(MachineName, 4)
        If cIRISComputerName <> "" Then
         .v2Value("C_WKST_ID") = UCase(Right(String(4, " ") & cIRISComputerName, 4))
        End If
    End With
End Sub

Public Function AppendMenu(pos, inName, InText) As Boolean
    Dim foundflag As Boolean, i As Integer
    Select Case pos
    Case 1:
        foundflag = False
        For i = MnuSub1.LBound To MnuSub1.UBound
            If MnuSub1(i).Tag = inName Then MnuSub1(i).Visible = True: MnuSub1(i).Caption = InText: foundflag = True: Exit For
        Next i
        If Not foundflag Then
        MnuSub1(MnuSub1.LBound).Visible = True: MNUItem(0).Visible = True: Load MnuSub1(MnuSub1.UBound + 1): MnuSub1(MnuSub2.LBound).Visible = False: MnuSub1(MnuSub1.UBound).Caption = InText: MnuSub1(MnuSub1.UBound).Tag = inName:
        End If
    Case 2:
        foundflag = False
        For i = MnuSub2.LBound To MnuSub2.UBound
            If MnuSub2(i).Tag = inName Then MnuSub2(i).Visible = True: MnuSub2(i).Caption = InText: foundflag = True: Exit For
        Next i
        If Not foundflag Then
        MnuSub2(MnuSub2.LBound).Visible = True: MNUItem2.Visible = True: Load MnuSub2(MnuSub2.UBound + 1): MnuSub2(MnuSub2.LBound).Visible = False: MnuSub2(MnuSub2.UBound).Caption = InText: MnuSub2(MnuSub2.UBound).Tag = inName:
        End If
    Case 3:
        foundflag = False
        For i = MnuSub3.LBound To MnuSub3.UBound
            If MnuSub3(i).Tag = inName Then MnuSub3(i).Visible = True: MnuSub3(i).Caption = InText: foundflag = True: Exit For
        Next i
        If Not foundflag Then
        MnuSub3(MnuSub2.LBound).Visible = True: MNUItem3.Visible = True: Load MnuSub3(MnuSub3.UBound + 1): MnuSub3(MnuSub3.LBound).Visible = False: MnuSub3(MnuSub3.UBound).Caption = InText: MnuSub3(MnuSub3.UBound).Tag = inName:
        End If
    Case 4:  MnuSub4(MnuSub2.LBound).Visible = True: MNUItem4.Visible = True: Load MnuSub4(MnuSub4.UBound + 1): MnuSub4(MnuSub4.LBound).Visible = False: MnuSub4(MnuSub4.UBound).Caption = InText: MnuSub4(MnuSub4.UBound).Tag = inName:
    Case 5:  MnuSub5(MnuSub2.LBound).Visible = True: MNUItem5.Visible = True: Load MnuSub5(MnuSub5.UBound + 1): MnuSub5(MnuSub5.LBound).Visible = False: MnuSub5(MnuSub5.UBound).Caption = InText: MnuSub5(MnuSub5.UBound).Tag = inName:
    Case 6:  MnuSub6(MnuSub2.LBound).Visible = True: MNUItem6.Visible = True: Load MnuSub6(MnuSub6.UBound + 1): MnuSub6(MnuSub6.LBound).Visible = False: MnuSub6(MnuSub6.UBound).Caption = InText: MnuSub6(MnuSub6.UBound).Tag = inName:
    Case 7:  MnuSub7(MnuSub2.LBound).Visible = True: MNUItem7.Visible = True: Load MnuSub7(MnuSub7.UBound + 1): MnuSub7(MnuSub7.LBound).Visible = False: MnuSub7(MnuSub7.UBound).Caption = InText: MnuSub7(MnuSub7.UBound).Tag = inName:
    Case 8:  MnuSub8(MnuSub2.LBound).Visible = True: MNUItem8.Visible = True: Load MnuSub8(MnuSub8.UBound + 1): MnuSub8(MnuSub8.LBound).Visible = False: MnuSub8(MnuSub8.UBound).Caption = InText: MnuSub8(MnuSub8.UBound).Tag = inName:
    Case 9:  MnuSub9(MnuSub2.LBound).Visible = True: MNUItem9.Visible = True: Load MnuSub9(MnuSub9.UBound + 1): MnuSub9(MnuSub9.LBound).Visible = False: MnuSub9(MnuSub9.UBound).Caption = InText: MnuSub9(MnuSub9.UBound).Tag = inName:
    End Select

'   Dim mnuItemInfo As MENUITEMINFO, hMenu As Long, hSubMenu As Long
'
'   hMenu = GetMenu(Me.hwnd)   ' Retrieve menu handle.
'   hSubMenu = GetSubMenu(hMenu, 0)
'   InsertMenu hSubMenu, 0, (MF_BYPOSITION Or MF_STRING Or MF_ENABLED), CStr(inName), CStr(inText)
'   hSubMenu = GetSubMenu(hSubMenu, 0)
'   DrawMenuBar (Me.hwnd)   ' Repaint top level Menu.
'   'SetWindowLong hSubMenu, GWL_WNDPROC, AddressOf amenu_Click
''   lpPrevWndProc = SetWindowLong(hSubMenu, GWL_WNDPROC, AddressOf WindowProc)
'
End Function

Public Sub Hook()
   lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
   temp = SetWindowLong(hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Function GetXMLWebLink() As cXMLWebLink
    Set GetXMLWebLink = New cXMLWebLink
End Function

Public Function GetSoapClient() As CSoapClient
    Set GetSoapClient = New CSoapClient
End Function

Public Function CalcGenBankChequeCD(inCheque) As Long
    CalcGenBankChequeCD = CalcGenBankChequeCD_(CStr(inCheque))
End Function

Public Function CalcETEChequeCD(inNum) As Long
    CalcETEChequeCD = CalcETEChequeCD_(CLng(inNum))
End Function

Public Function ChkCRAOutput(aBuffer, Optional ErrorView) As Boolean
    If IsMissing(ErrorView) Then ErrorView = ""
    ChkCRAOutput = ChkCRAOutput_(aBuffer, ErrorView)
End Function

Public Function WebLink(linkName As String) As String
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
        WebLink = WebLinks(UCase("EDUC" & linkName))
    ElseIf Left(Right(WorkEnvironment_, 8), 4) = "PROD" Then
        WebLink = WebLinks(UCase("PROD" & linkName))
    Else
        WebLink = ""
    End If
    
End Function

Public Function XML() As String
    XML = XMLFormView.XML
End Function

Public Function XMLFormView() As MSXML2.DOMDocument30
Dim xmlTrn As MSXML2.DOMDocument30
Dim aattr As IXMLDOMAttribute
Dim elm As IXMLDOMElement
Dim newElm As IXMLDOMElement
Dim newAttr As IXMLDOMAttribute
Dim i As Integer


Set xmlTrn = New MSXML2.DOMDocument30


'xmlTrn.appendChild CreateXMLNode(xmlTrn, TrnFrmNamespace, "TRN")
xmlTrn.appendChild CreateXMLNode(xmlTrn, "", "TRN")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "FIELDS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "LABELS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "BUTTONS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "CHECKS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "CHARTS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "COMBOS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "LISTS")
'xmlTrn.documentElement.appendChild CreateXMLNode(xmlTrn, "none", "GRIDS")

Set elm = xmlTrn.documentElement
If TrnVariables.count > 0 Then
    For i = 1 To TrnVariables.count
        Set newElm = xmlTrn.createElement("variable")
        Set newAttr = xmlTrn.createAttribute("name")
        newAttr.value = TrnVariables.item(i).name
        newElm.setAttributeNode newAttr
        Set newAttr = xmlTrn.createAttribute("value")
        newAttr.value = TrnVariables(i).value
        newElm.setAttributeNode newAttr
        elm.appendChild newElm
    Next i
End If
If fields.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//FIELDS")
    For i = 1 To fields.count
        Set newElm = fields.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next

End If
If Labels.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//LABELS")
    For i = 1 To Labels.count
        Set newElm = Labels.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Buttons.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//BUTTONS")
    For i = 1 To Buttons.count
        Set newElm = Buttons.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Checks.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//CHECKS")
    For i = 1 To Checks.count
        Set newElm = Checks.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Charts.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//CHARTS")
    For i = 1 To Charts.count
        Set newElm = Charts.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Combos.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//COMBOS")
    For i = 1 To Combos.count
        Set newElm = Combos.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Lists.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//LISTS")
    For i = 1 To Lists.count
        Set newElm = Lists.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
If Spreads.count > 0 Then
    'Set Elm = xmlTrn.selectSingleNode("//GRIDS")
    For i = 1 To Spreads.count
        Set newElm = Spreads.item(i).TranslateToProperties(processphase)
        elm.appendChild newElm
    Next
End If
xmlTrn.documentElement.normalize

Dim astr As String
astr = xmlTrn.XML
astr = Replace(astr, "xmlns=""""", "")
xmlTrn.LoadXML astr
'xmlTrn.save ("c:\tt.xml")

Set XMLFormView = xmlTrn
'If Fields.count > 0 Then
'    Fields.Item(1).Text = "eeeee"
'    SetXMLValue Fields.Item(1).Name
'
'End If
'
'If Buttons.count > 0 Then
'    SetXMLValue Buttons.Item(1).Name
'End If
'If Checks.count > 0 Then
'    SetXMLValue Checks.Item(1).Name
'End If
'If Combos.count > 0 Then
'    SetXMLValue Combos.Item(2).Name
'End If
'If Lists.count > 0 Then
'    SetXMLValue Lists.Item(1).Name
'End If
'If Spreads.count > 0 Then
'    SetXMLValue Spreads.Item(1).Name
'End If

End Function

Public Sub LoadXML(invalue As String)
    Dim adoc As New MSXML2.DOMDocument30
    Dim elm As IXMLDOMElement
    adoc.LoadXML invalue
    
    With adoc.documentElement
        For Each elm In .SelectNodes(".//TEXTBOX | .//COMBOBOX | .//BUTTON | .//GRID ")
            ApplyXMLControlUpdate elm, "UPDATE"
        Next elm
        For Each elm In .SelectNodes(".//variable")
            TrnVariable(elm.getAttribute("name")) = elm.getAttribute("value")
        Next elm
    End With
    RefreshView
    
End Sub

Public Sub ApplyXMLFormUpdate(inDoc, inUpdateName)
    Dim elm As IXMLDOMElement
    If inDoc.selectSingleNode(".//FORMUPDATE[@TITLE=""" & inUpdateName & """]") Is Nothing Then
        MsgBox "Δεν βρέθηκε ο μετασχηματισμός οθόνης " & inUpdateName, vbCritical, "ΛΑΘΟΣ..."
    Else
        With inDoc.selectSingleNode(".//FORMUPDATE[@TITLE=""" & inUpdateName & """]")
            For Each elm In .SelectNodes(".//TEXTBOX | .//COMBOBOX | .//BUTTON | .//GRID ")
                ApplyXMLControlUpdate elm, CStr(inUpdateName)
            Next elm
            For Each elm In .SelectNodes(".//variable")
                TrnVariable(elm.getAttribute("name")) = elm.getAttribute("value")
            Next elm
        End With
        RefreshView
    End If
End Sub

Public Function TransformToXML(sourcedoc, translationdoc, translationname) As MSXML2.DOMDocument30
    Set TransformToXML = New MSXML2.DOMDocument30
    If translationdoc.selectSingleNode(".//TRANSFORMATION[@TITLE=""" & translationname & """]/*") Is Nothing Then
        MsgBox "Δεν βρέθηκε ο μετασχηματισμός " & translationname, vbCritical, "ΛΑΘΟΣ...."
        Set TransformToXML = Nothing
        Exit Function
    End If
    Dim mergedDoc As New MSXML2.DOMDocument30
    mergedDoc.LoadXML "<ROOT>" & xmlEnvironment.XML & sourcedoc.XML & "</ROOT>"
    TransformToXML.LoadXML mergedDoc.transformNode(translationdoc.selectSingleNode(".//TRANSFORMATION[@TITLE=""" & translationname & """]/*"))
End Function

Public Function TransformToXMLFromFile(sourcedoc, filename, translationname) As MSXML2.DOMDocument30
    Dim transformdoc As New MSXML2.DOMDocument30
    transformdoc.Load filename
    Set TransformToXMLFromFile = TransformToXML(sourcedoc, transformdoc, translationname)
End Function

Public Sub ApplyXMLControlUpdate(elm As IXMLDOMElement, inUpdateName As String)
    Dim Control
    On Error GoTo continue_control
    Set Control = NamedControls.item(elm.getAttribute("FULLNAME"))
    On Error GoTo 0
    Control.SetXMLValue elm, processphase
continue_control:
    On Error GoTo 0
    If IsEmpty(Control) Then
        MsgBox "Δεν βρέθηκε το στοιχείο " & elm.getAttribute("FULLNAME") & " κατα τον μετασχηματισμό " & inUpdateName, vbCritical, "ΛΑΘΟΣ...."
    End If
    Set Control = Nothing
End Sub

Public Sub ApplyXMLFormUpdateFromFile(inFileName, inUpdateName)
Dim adoc As MSXML2.DOMDocument30
Set adoc = New MSXML2.DOMDocument30
    adoc.Load inFileName
    ApplyXMLFormUpdate adoc, inUpdateName
End Sub

Public Function ApplyXMLValidationFailure(ResponseDoc As MSXML2.DOMDocument30, statement As IXMLDOMElement)
    Dim elm As IXMLDOMElement
    Set elm = ResponseDoc.createElement("ERROR")
    
    elm.appendChild ResponseDoc.createElement("LINE")
    elm.appendChild ResponseDoc.createElement("EXCEPTIONTYPE")
    elm.firstChild.Text = statement.getAttribute("MESSAGE")
    ResponseDoc.documentElement.appendChild elm
End Function

Public Function ApplyXMLValidation(sourcedoc, validationdoc, validationname) As MSXML2.DOMDocument30
    Set ApplyXMLValidation = New MSXML2.DOMDocument30
    ApplyXMLValidation.appendChild ApplyXMLValidation.createElement("VALIDATIONRESULT")
    Dim elm As IXMLDOMElement
    For Each elm In validationdoc.SelectNodes(".//VALIDATION[@TITLE=""" & validationname & """]/STATEMENT")
        Dim selecttype As String: selecttype = "ROWS"
        If elm.getAttribute("SELECTTYPE") = "NOROWS" Then selecttype = "NOROWS"
        Dim aexpression As IXMLDOMAttribute
        Dim expressionvalue As String
        expressionvalue = elm.getAttribute("EXPRESSION")
        If selecttype = "ROWS" Then
            If sourcedoc.SelectNodes(expressionvalue).length > 0 Then
                ApplyXMLValidationFailure ApplyXMLValidation, elm
            End If
            
        Else
            If sourcedoc.SelectNodes(expressionvalue).length = 0 Then
                ApplyXMLValidationFailure ApplyXMLValidation, elm
            End If
        End If
    Next elm
End Function

Public Function ApplyXMLValidationFromFile(sourcedoc, filename, validationname) As MSXML2.DOMDocument30
    Dim validationdoc As New MSXML2.DOMDocument30
    validationdoc.Load filename
    Set ApplyXMLValidationFromFile = ApplyXMLValidation(sourcedoc, validationdoc, validationname)
End Function

Public Function CheckXMLMessages(messagedoc) As Boolean
    CheckXMLMessages = True
    If messagedoc.SelectNodes(".//ERROR").length > 0 Then
        Load XMLMessageForm
        Set XMLMessageForm.MessageDocument = messagedoc
        XMLMessageForm.Show vbModal, Me
        CheckXMLMessages = False
    ElseIf messagedoc.SelectNodes(".//WARNING").length > 0 Then
        Load XMLMessageForm
        Set XMLMessageForm.MessageDocument = messagedoc
        XMLMessageForm.Show vbModal, Me
        CheckXMLMessages = True
    End If
End Function


Public Function LoadTemplatesFromFile(inFileName As String)
    Set xmlDocumentManager = New cXMLDocumentManager
    'Set xmlDocumentManager.OwnerForm = Me
    
    Dim templatedoc As New MSXML2.DOMDocument30
    templatedoc.Load inFileName
    
    'xmlDocumentManager.LoadTemplates templatedoc.documentElement
End Function

Public Function execJob(jobName As String)
    xmlDocumentManager.XmlObjectList.item(jobName).XML
End Function


Public Function LinkToL2TRN(Document) As String
    Dim LinkDocument As New MSXML2.DOMDocument30
    Document = Replace(Document, "&", "&amp;")
    LinkDocument.LoadXML Document
    
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    Dim TrnCode As String, inDoc As MSXML2.DOMDocument30, outDoc As MSXML2.DOMDocument30
    Dim trnHandler As L2TrnHandler
    
    For Each elm In LinkDocument.documentElement.childNodes
        If UCase(elm.baseName) = UCase("trn") Then
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then TrnCode = attr.Text
            If TrnCode <> "" Then
                Set trnHandler = New L2TrnHandler
            Else
                
            End If
        ElseIf UCase(elm.baseName) = UCase("formupdate") Then
            If trnHandler Is Nothing Then
            Else
                Set inDoc = New MSXML2.DOMDocument30
                inDoc.LoadXML elm.XML
                trnHandler.addFormUpdate inDoc, inDoc.documentElement.getAttribute("name")
            End If
        End If
    Next elm
    
    If Not (trnHandler Is Nothing) Then
        trnHandler.ExecuteForm TrnCode
        
        If trnHandler.Result Is Nothing Then
            LinkToL2TRN = ""
        Else
            LinkToL2TRN = trnHandler.Result.XML
        End If
        trnHandler.CleanUp
    End If
End Function

Public Function GetWebMethodLink(aWeblink, aname As String, anamespace As String) As cXMLWebMethod
    Set GetWebMethodLink = aWeblink.DefineDocumentMethod(aname, anamespace)
    Set GetWebMethodLink.content = Nothing
End Function

Public Function GetWebMessageLink(VirtualDirectory As String) As cXMLWebLink
    Set GetWebMessageLink = New cXMLWebLink
    GetWebMessageLink.VirtualDirectory = WebLink(VirtualDirectory)
    Set GetWebMessageLink.content = Nothing
End Function

Public Function GetWebMessage(namespace As String, messagename As String) As msgmember
    If amsgmemberconstructor Is Nothing Then
        Set amsgmemberconstructor = New msgmemberwsconstructor
    End If
    Set GetWebMessage = amsgmemberconstructor.buildmessage(namespace, messagename)
    
End Function

Public Function GetWebMessageWrapper(namespace As String, Optional messagename As String) As msgwrapper
    If amsgwrapperconstructor Is Nothing Then
        Set amsgwrapperconstructor = New msgwrapperwsconstructor
    End If
    If IsMissing(messagename) Then
        Set GetWebMessageWrapper = amsgwrapperconstructor.buildwrapper(namespace)
    Else
        Set GetWebMessageWrapper = amsgwrapperconstructor.buildwrapper(namespace, messagename)
    End If
End Function

Public Function DirectCommunicate(output, encodegreek)
'    Dim resultList As Collection
'    DirectCommunicate = newCommunicate(CStr(output), CBool(encodegreek), resultList)
'    Set resultList = Nothing

    DirectCommunicate = 0
End Function

Public Function getHtmlReportObject() As cHtmlReportObject
    Set getHtmlReportObject = New cHtmlReportObject
End Function

Public Function getHostMethod(source, paramnames(), paramvalues()) As cScriptHostMethodBuilder
    Set getHostMethod = New cScriptHostMethodBuilder
    getHostMethod.buildmethod CStr(source), paramnames, paramvalues
    
End Function

Public Function onlineCom(module As String, InputView, InputViewName As String, OutputViewName As String, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
'    onlineCom = onlineCom_(Me, Module, InputView, InputViewName, OutputViewName, AuthUser, Appltran, ErrorView, ErrorCount)
    onlineCom = 0
End Function

Public Function onlineComWithLog(module As String, InputView, InputViewName As String, OutputViewName As String, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
    Dim aFlag As Boolean
    aFlag = LogIrisCom
    LogIrisCom = True
'    onlineComWithLog = onlineCom_(Me, Module, InputView, InputViewName, OutputViewName, AuthUser, Appltran, ErrorView, ErrorCount)
    onlineComWithLog = 0
    LogIrisCom = aFlag
End Function


Public Sub BuildComArea(filename As String, StructureName As String, Optional StructureAlias As String, Optional hidden As Boolean)
    
'    If Not cIRISConnected Then
'        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (BuildIRISStruct)", vbCritical: Exit Sub
'    End If
On Error GoTo GenError
    
    If IsMissing(StructureAlias) Then StructureAlias = StructureName
    If IsMissing(hidden) Then hidden = False
    
    If Trim(StructureAlias) = "" Then StructureAlias = StructureName
    Dim structurecode As String
    structurecode = ""
    
    On Error Resume Next
    Close #1
    On Error GoTo 0
    Dim s As String
    Open ReadDir & ComAreaDir & filename & ".txt" For Input As #1
    Do While Not Eof(1)
        Line Input #1, s
        structurecode = structurecode & s
        
        'Dim tokens() As String
        'tokens = Split(s, vbTab)
    Loop
    Close #1
    
    structurecode = Replace(structurecode, vbTab, " ")
    Dim res As Buffer
    Set res = TrnBuffers.DefineComArea(structurecode, StructureName, hidden)
    If res Is Nothing Then
    
    Else
        ValidationControl.ExecuteStatement "set " & Replace(StructureName, "@", "_") & "= Trnbuffers.ByName(""" & StructureName & """)"
    End If
    Exit Sub
GenError:
    MsgBox "Πρόβλημα στη δήλωση της δομής: " & StructureAlias & " Err: " & CStr(Err.number) & " - " & Err.description, vbCritical, "ΛΑΘΟΣ"
End Sub


Public Function BuildSWIFTmessage(messagename)
    Dim amessage As cSWIFTmessage
    Set amessage = New cSWIFTmessage
    amessage.prepare CStr(messagename)
    Set BuildSWIFTmessage = amessage

End Function

Public Sub ExitForm()
    If CurrAction = taGet_Input Or CurrAction = taStay_In_Form Or CurrAction = taExit_Form Then
        DoEvents
        NextAction = taEscape_Form
        OpAfterHandling = fopExitForm
        'ProcessLoop
    End If
End Sub

Public Function GetKAAMessage(aname) As cKAAMessage
    Set GetKAAMessage = New cKAAMessage
    GetKAAMessage.prepare CStr(aname)
End Function

Public Sub PrintSwiftMessage(Lst As String, amessage)
    Dim astr() As String
    astr = amessage.PrintSwiftMessage
    Dim i As Integer
    For i = 0 To UBound(astr)
        Me.Controls(Lst).AddItem astr(i)
    Next
End Sub

Public Function AddAppCRecordset(inName, inCmd, inDBName, inVirtualDirectory, Optional inCursorType, Optional inLockType) As cADORecordset
    Set AddAppCRecordset = AddAppCRecordset_(inName, inCmd, inDBName, inVirtualDirectory, inCursorType, inLockType)
End Function

 
Public Function AppCRecordsetByName(inName) As cADORecordset
    Set AppCRecordsetByName = AppCRecordsetByName_(inName)
End Function

Public Function AppCRSEntryByName(inName) As RecordsetEntry
    Set AppCRSEntryByName = AppCRSEntryByName_(inName)
End Function

Public Function ChkDocumentNo(countrycode As String, docno As String) As Boolean
    ChkDocumentNo = ChkDocumentNo_(countrycode, docno)
End Function
Public Function NetworkHomeDir() As String
    On Error GoTo creationError
    Dim fso
    Dim folder
    Dim foldername As String
    foldername = WorkDir & MachineName
    Set fso = CreateObject("Scripting.FileSystemObject")
    foldername = "\" & Replace(foldername, "\\", "\")
    If fso.FolderExists(foldername) Then
        Set folder = fso.GetFolder(foldername)
        NetworkHomeDir = foldername
        Exit Function
    Else
        Set folder = fso.CreateFolder(foldername)
        NetworkHomeDir = foldername
        Exit Function
    End If
    Exit Function
creationError:
     MsgBox "Η δημιουργία του  " & foldername & " απέτυχε.", vbInformation
End Function

Public Function TotalsVersion() As Long
    TotalsVersion = cTotalsVersion
End Function
