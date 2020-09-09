VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function GetBankName(BankCode) As String
    GetBankName = GetBankName_(CInt(BankCode))
End Function

Public Function ISOTOCURR(ByVal inUnit) As String
     ISOTOCURR = ISOTOCURR_(CStr(inUnit))
End Function

Public Function CURRTOISO(ByVal inUnit) As String
     CURRTOISO = CURRTOISO_(CStr(inUnit))
End Function

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
Public Function IRISUSERName() As String
    IRISUSERName = cIRISUserName
End Function

Public Function ChkChiefTeller() As Boolean
    ChkChiefTeller = isChiefTeller
End Function

Public Function ChkManager() As Boolean
    ChkManager = isManager
End Function

Public Function StrPad(PString, PLength, Optional PChar, Optional PLftRgt) As String
' η πιό αξιόλογη ρουτίνα της εφαρμογής. Αν δεν ξέρεις τι κάνει πάτα Alt-F4
    StrPad = StrPad_(CStr(PString), CInt(PLength), PChar, PLftRgt)
End Function

Public Function GetPassbookAmount(inAmount As Double) As String
' μορφοποίηση ποσού για εκτύπωση σε βιβλιάριο (συμπληρωμένο με *)
    GetPassbookAmount = GetPassbookAmount_(CDbl(inAmount))
End Function


Public Function ChkBankAcount(inBank, inbranch, inAcc) As Boolean
'έλεγχος check digit λογαριασμού τράπεζας
    ChkBankAcount = ChkBankAccount_(CStr(inBank), CStr(inbranch), CStr(inAcc))
End Function

Public Function ChkBankCheque(inBank, inbranch, inAcc, inCheque) As Boolean
'έλεγχος check digit επιταγής τράπεζας
    ChkBankCheque = ChkBankCheque_(CStr(inBank), CStr(inbranch), CStr(inAcc), CStr(inCheque))
End Function

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
Public Function GetPostDate_FComArea() As String
' ημερομηνία λειτουργίας σε DD.MM.YYYY μορφή
    GetPostDate_FComArea = format(cPOSTDATE, "DD.MM.YYYY")
End Function

Public Function GetTerminalID() As String
' ταυτότητα τερματικού
    GetTerminalID = cTERMINALID
End Function

Public Function GetBranchCode() As String
' κωδικός καταστήματος
    GetBranchCode = StrPad_(CStr(IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH)), 3, "0", "L")
End Function

Public Function GetBranchName() As String
' ονομασία καταστήματος
    GetBranchName = cBRANCHName
    
End Function
Public Function IRISComputerName() As String
    IRISComputerName = cIRISComputerName
    
End Function
Public Function ComputerName() As String
    ComputerName = MachineName
End Function
Public Function ForcedCRAUser() As String
    ForcedCRAUser = cIRISUserName
End Function
Public Function UserName() As String
    UserName = cUserName
End Function
Public Function CRABranch() As String
    CRABranch = cBRANCH
End Function
Public Function GetAmountText2002(inAmount) As String
' ποσό ολογράφως
    GetAmountText2002 = Amount_Str2002(CDbl(inAmount))
End Function

Public Function GetAmountText(inAmount) As String
' ποσό ολογράφως
    GetAmountText = Amount_str(CDbl(inAmount), True)
End Function

Public Function GetChequeAmountText(inAmount) As String
' ποσό ολογράφως ειδικά για επιταγές συναλλαγή 6000
    GetChequeAmountText = Cheque_Amount_str_(CDbl(inAmount), False)
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

Public Sub ClearTotals()
' μηδενισμός αθροιστών
End Sub

Public Sub ClearAllTotal(TotalName)
' πλήρης μηδενισμός αθροιστή νομίσματος (σε όλα τα νομίσματα)
End Sub

Public Sub ClearCurTotal(inTotal, inCurrency)
' μηδενισμός αθροιστή νομίσματος σε συγκεκριμένο νόμισμα
End Sub

Public Function GetADOConnection() As ADODB.Connection

End Function

Public Function GetADORecordset(inCmd, Optional inCursorType, Optional inLockType) As ADODB.Recordset

End Function

Public Function ExecADOCommand(aCommandStr) As Integer

End Function

Public Function AddAppRecordset(inName, inCmd, Optional inConnection, Optional inCursorType, Optional inLockType) As ADODB.Recordset
    Set AddAppRecordset = Nothing
End Function

Public Function AddFTFilaRecordset(inName, inFilter) As ADODB.Recordset
    Set AddFTFilaRecordset = AddFTFilaRecordset_(inName, inFilter)
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

Public Function IRISCom(Trn As String, rule As String, InputView, OutputView, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
    IRISCom = IRISCom_(GenWorkForm, Trn, rule, InputView, OutputView, AuthUser, Appltran, ErrorView, ErrorCount)
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

Public Function SQLServerName() As String

End Function

Public Sub BuildIRISAppStruct(StructureName, Optional StructureAlias, Optional hidden As Boolean)
    If IsMissing(hidden) Then hidden = False
    DatabaseMdl.BuildIRISAppStruct StructureName, StructureAlias, Not hidden
End Sub

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

Public Sub FreeAppStruct(StructureName)
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (FreeAppStruct)", vbCritical: Exit Sub
    End If
    
    GenWorkForm.AppBuffers.FreeBuffer CStr(StructureName)
End Sub

Public Sub cmdForceLogoff()
    cmdForceLogoff_
End Sub

Public Property Get AppBuffers()
    Set AppBuffers = GenWorkForm.AppBuffers
End Property


'    ValidationControl.AddObject "AppRecordset", GenWorkForm.AppRS, True
'    ValidationControl.AddObject "AppStoredProcedure", GenWorkForm.AppSP, True

Public Property Get ActiveL2Trn() As L2TrnHandler
    Set ActiveL2Trn = ActiveL2TrnHandler
End Property
'-------------------------------
'Public Function ParseComArea(area As cXmlComArea) As String
Public Function ParseComArea(area) As String
    Dim res As String
    If Not (area Is Nothing) Then
        If Not (area.content Is Nothing) Then
           'ParseComArea = area.LoadXml("<P49SP><IDATA><ISIN>EP0303200501</ISIN></IDATA></P49SP>")
           'res = area.ParseBufferCall(area.content)
           'MsgBox "1"
            UpdateTrnNum_
            res = area.ParseCall(Nothing)
            'MsgBox "2"
           If res = "" Then
           'MsgBox "3"
                ParseComArea = "<ERROR></ERROR>"
           Else
             'MsgBox "3"
                ParseComArea = res
                 'MsgBox res
           End If
           'ParseComArea = area.ParseBufferCall(area.content)
        End If
    End If
End Function

Public Function DeclareComArea(name As String, comareaid As String, method As String, TrnId As String, filename As String, InputName As String, OutputName As String) As cXmlComArea
    Set DeclareComArea = DeclareComArea_(name, comareaid, method, TrnId, filename, InputName, OutputName)
End Function

