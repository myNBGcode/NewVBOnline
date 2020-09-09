VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TotalEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Type TotalEntry
    TotalName As String
    TotalPos As Integer
    Currency As Integer
    Amount As Double
End Type

Private TotalEntriesList() As TotalEntry
Private ListSize As Integer, EntriesNum As Integer

Private Sub Class_Initialize()
    EntriesNum = 0: ListSize = 10: ReDim TotalEntriesList(10)
End Sub

Public Sub ClearEntries_()
'καθαρίζει τη λίστα εγγραφών για τους αθροιστές
    EntriesNum = 0
End Sub

Public Sub AddDBEntry_(inTotalName, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    If ListSize = EntriesNum Then ReDim Preserve TotalEntriesList(ListSize + 10): ListSize = ListSize + 10
    EntriesNum = EntriesNum + 1
    With TotalEntriesList(EntriesNum)
        .TotalName = inTotalName: .TotalPos = 1: .Currency = 0: .Amount = inAmount
    End With
End Sub

Public Sub AddCREntry_(inTotalName, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    If ListSize = EntriesNum Then ReDim Preserve TotalEntriesList(ListSize + 10): ListSize = ListSize + 10
    EntriesNum = EntriesNum + 1
    With TotalEntriesList(EntriesNum)
        .TotalName = inTotalName: .TotalPos = 2: .Currency = 0: .Amount = inAmount
    End With
End Sub

Public Sub AddCurDBEntry_(inTotalName, inCurrency, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    If ListSize = EntriesNum Then ReDim Preserve TotalEntriesList(ListSize + 10): ListSize = ListSize + 10
    EntriesNum = EntriesNum + 1
    With TotalEntriesList(EntriesNum)
        .TotalName = inTotalName: .TotalPos = 1: .Currency = inCurrency: .Amount = inAmount
    End With
End Sub

Public Sub AddCurCrEntry_(inTotalName, inCurrency, inAmount)
'προσθέτει εγγραφή στη λίστα αθροιστών
    If ListSize = EntriesNum Then ReDim Preserve TotalEntriesList(ListSize + 10): ListSize = ListSize + 10
    EntriesNum = EntriesNum + 1
    With TotalEntriesList(EntriesNum)
        .TotalName = inTotalName: .TotalPos = 2: .Currency = inCurrency: .Amount = inAmount
    End With
End Sub

Public Sub StoreEntries_()
'ενημερώνει τους αθροιστές από τη λίστα αθροιστών
Dim i As Integer, Line As Long, aname As String, acurr As Integer, aPos As Integer, aAmount As Double
On Error GoTo ErrMessage
Line = 1
    For i = 1 To EntriesNum
Line = 2
        With TotalEntriesList(i)
Line = 3
            aname = .TotalName
Line = 4
            acurr = .Currency
Line = 5
            aAmount = .Amount
Line = 6
            aPos = .TotalPos
Line = 7
            If acurr = 0 Then
Line = 8
                If aPos = 1 Then AddDBTotal_ aname, aAmount Else AddCRTotal_ aname, aAmount
Line = 9
            Else
Line = 10
                If aPos = 1 Then AddCurDBTotal_ aname, acurr, aAmount Else AddCurCRTotal_ aname, acurr, aAmount
Line = 11
            End If
Line = 12
        End With
Line = 13
    Next i
Line = 14
    Exit Sub
ErrMessage:
    LogMsgbox "Λαθος κατα την ενημέρωση αθροιστων Γραμμη: " & CStr(Line) & " - " & Err.Number & " - " & Err.description
End Sub

