VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private linkedForm As TRNFrm

Public Property Get TrnVariable(inName As String)
    TrnVariable = linkedForm.TrnVariable(inName)
End Property

Public Function TrnVariableDouble(inName As String) As Double
    TrnVariableDouble = linkedForm.TrnVariableDouble(inName)
End Function

Public Function TrnVariableInteger(inName As String) As Long
    TrnVariableInteger = linkedForm.TrnVariableInteger(inName)
End Function

Public Property Let TrnVariable(inName As String, invalue)
    linkedForm.TrnVariable(inName) = invalue
End Property

Public Sub UnlockPrinter()
    linkedForm.UnlockPrinter
End Sub

Public Sub sbShowCommStatus(ByVal bActive As Boolean)
    linkedForm.sbShowCommStatus (bActive)
End Sub

Public Function GetChiefTellerKey() As Boolean
    GetChiefTellerKey = linkedForm.GetChiefTellerKey
End Function

Public Function GetManagerKey() As Boolean
    GetManagerKey = linkedForm.GetManagerKey
End Function

Public Sub DisableSendFromJournal()
   linkedForm.DisableSendFromJournal
End Sub

Public Sub EnableSendFromJournal()
   linkedForm.EnableSendFromJournal
End Sub

Public Sub DisableReceiveFromJournal()
   linkedForm.DisableReceiveFromJournal
End Sub

Public Sub EnableReceiveFromJournal()
   linkedForm.EnableReceiveFromJournal
End Sub

Public Sub SetTRNCode(inPhase As Integer, inTrnCode As String)
    linkedForm.SetTRNCode inPhase, inTrnCode
End Sub

Public Sub ClearTotalEntries()
'��������� �� ����� �������� ��� ���� ���������
    linkedForm.ClearTotalEntries
End Sub

Public Sub AddDBTotalEntry(inTotalName, inAmount)
'��������� ������� ��� ����� ���������
    linkedForm.AddDBTotalEntry inTotalName, inAmount
End Sub

Public Sub AddCRTotalEntry(inTotalName, inAmount)
'��������� ������� ��� ����� ���������
    linkedForm.AddCRTotalEntry inTotalName, inAmount
End Sub

Public Sub AddCurDBTotalEntry(inTotalName, inCurrency, inAmount)
'��������� ������� ��� ����� ���������
    linkedForm.AddCurDBTotalEntry inTotalName, inCurrency, inAmount
End Sub

Public Sub AddCurCrTotalEntry(inTotalName, inCurrency, inAmount)
'��������� ������� ��� ����� ���������
    linkedForm.AddCurCrTotalEntry inTotalName, inCurrency, inAmount
End Sub

Public Sub StoreTotalEntries()
'���������� ���� ��������� ��� �� ����� ���������
    linkedForm.StoreTotalEntries
End Sub

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
    linkedForm.sbWriteStatusMessage sMessage
End Sub

Public Function fnReadStatusMessage() As String
    fnReadStatusMessage = linkedForm.fnReadStatusMessage
End Function

Public Function PrepareSendOut() As Boolean
    PrepareSendOut = linkedForm.PrepareSendOut
End Function

Public Function SendOut() As Boolean
    SendOut = linkedForm.SendOut
End Function

Public Sub SetOutBuffer(outString)
    linkedForm.SetOutBuffer outString
End Sub

Public Sub ReadIn()
'1.������� �� AfterInScript
'2.��������� ��� ����� �� ����������� ��� buffer ��� �������� ��� �� ��
'3.��� ���������� ��� ���������� �� �����
'���������� TRUE �� � �������� ����������� ����� ��������
    linkedForm.ReadIn
End Sub

Public Sub ReadBuffer()
'���� �������������� ��� �� ReadIn
'����������� ���� ��� ����������� �� ��� ���������� ��� ����� ���� ���������������
    linkedForm.ReadBuffer
End Sub

Public Sub xClearDoc()
' ���������� ������������ ������������
    linkedForm.xClearDoc
End Sub

Public Sub xSetDocLine(inLineNo, inLineData)
' ������� ����� �� ������ ������������
    linkedForm.xSetDocLine CInt(inLineNo), CStr(inLineData)
End Sub

Public Sub xSetInDocLine(inLineNo, inLineData, inX, inW, inAlign)
' ������� ����� �� ������� ������������
' Align: "L" � "R"
    linkedForm.xSetInDocLine CInt(inLineNo), CStr(inLineData), CInt(inX), CInt(inW), CStr(inAlign)
End Sub

Public Sub xPrintDoc(Optional inPrompt)
' �������� ��������� ������ ������������
    linkedForm.xPrintDoc
End Sub

Public Sub SetInBuffer(inString As String)
' �������� ��� buffer ����� ��������� ��� ��
    linkedForm.SetInBuffer inString
End Sub

Public Function GetTrnNum() As String
' ������� ��������� ����������
    GetTrnNum = linkedForm.GetTrnNum
End Function

Public Sub WriteJournal(amessage)
' ������� ��������� ��� ������ �����������
    linkedForm.WriteJournal (amessage)
End Sub

Public Sub WriteJournalFinal()
' ������� �������� ������ ���������� ��� ������ �����������
    linkedForm.WriteJournalFinal
End Sub

Public Sub SetKey(aKey)
' ������ ��� ������������ ��� �� ��������� ��������
    linkedForm.SetKey aKey
End Sub

Public Function WriteErrorMessage(inMessage) As Boolean
'������� ��� ���������� ��� ��� ������ ��������� ������ ������
    linkedForm.WriteErrorMessage inMessage
End Function

Public Sub SEND()
'Public Sub SEND(Sender As Control)
    linkedForm.SEND
End Sub

Public Function WriteTotals() As Boolean
    WriteTotals = linkedForm.WriteTotals
End Function

'Public Function ControlLostFocus(sender) As Boolean
'    ControlLostFocus = linkedForm.ControlLostFocus(sender)
'End Function

Public Function GetFormatedFld(inFldName As String) As String
    GetFormatedFld = linkedForm.GetFormatedFld(inFldName)
End Function

Public Sub SetDefaultFocus(inPhase As Integer)
'�������� �� ����� active ����� ��� ������
    linkedForm.SetDefaultFocus inPhase
End Sub

Public Sub RefreshView()
    linkedForm.RefreshView
End Sub

Public Function UpdateTrnNum()
    linkedForm.UpdateTrnNum
End Function

Public Sub UpdateScreen()
    linkedForm.UpdateScreen
End Sub

Public Sub PrintDoc(inPhase As Integer)
    linkedForm.PrintDoc inPhase
End Sub

Public Sub ProcessLoop()
    linkedForm.ProcessLoop
End Sub

Public Sub WriteJournalBeforeSend()
    linkedForm.WriteJournalBeforeSend
End Sub

Public Function GetIRISAuth() As String
    GetIRISAuth = linkedForm.GetIRISAuth
End Function


Public Function GetInBuffer() As String
    GetInBuffer = linkedForm.GetInBuffer
End Function

'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------
Public Sub WriteRegistry(RootKey, Key, Variable, value)
    linkedForm.WriteRegistry RootKey, Key, Variable, value
End Sub

Public Function StructDefine(aStructName As String, aStructDesc As String) As Integer
    StructDefine = linkedForm.StructDefine(aStructName, aStructDesc)
End Function

Public Sub DimTRNVariable(VariableName, Optional value)
    If IsMissing(value) Then Set value = Nothing
    linkedForm.DimTRNVariable VariableName, value
End Sub

Public Sub BuildIRISStruct(StructureName, Optional StructureAlias, Optional hidden As Boolean)
    If IsMissing(StructureAlias) Then StructureAlias = StructureName
    If IsMissing(hidden) Then hidden = False
    linkedForm.BuildIRISStruct StructureName, StructureAlias, hidden
End Sub

Public Sub HandleEvent(SenderName, code, State)
    Dim aflag As Boolean
    aflag = False
    linkedForm.HandleEvent SenderName, code, State, aflag
End Sub

Public Sub LinkToTrnV3(inTrnCD, PArray)
    linkedForm.LinkToTrnV3 inTrnCD, PArray
End Sub

Public Property Get Param(aname As String)
    Param = linkedForm.Param(aname)
End Property

Public Property Let Param(aname As String, value)
    linkedForm.Param(aname) = value
End Property

Public Function GetCRASetValue(SetName As String, Optional SetCode)
    GetCRASetValue = linkedForm.GetCRASetValue(SetName, SetCode)
End Function

Public Function IRISCom(Trn As String, Rule As String, InputView, OutputView, Optional AuthUser, Optional Appltran, Optional ErrorView, Optional ErrorCount) As Integer
    IRISCom = linkedForm.IRISCom(Trn, Rule, InputView, OutputView, AuthUser, Appltran, ErrorView, ErrorCount)
End Function

Public Function AppendMenu(pos, inName, InText) As Boolean
    AppendMenu = linkedForm.AppendMenu(pos, inName, InText)
End Function


