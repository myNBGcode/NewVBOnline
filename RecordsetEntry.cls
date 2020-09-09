VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordsetEntry"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public rs As ADODB.Recordset
Public Name As String

Public Function Fields()
    Fields = rs.Fields
End Function

Public Function NVLDouble(inFld, Optional retValue As Double) As Double
    If IsMissing(retValue) Then retValue = 0
    On Error GoTo ErrorPos
    NVLDouble = NVLDouble_(rs.Fields(inFld).Value, retValue)
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLDouble = retValue
End Function

Public Function NVLInteger(inFld, Optional retValue As Integer) As Long
    If IsMissing(retValue) Then retValue = 0
    On Error GoTo ErrorPos
    NVLInteger = NVLInteger_(rs.Fields(inFld).Value, retValue)
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLInteger = retValue
End Function

Public Function NVLString(inFld, Optional retValue As String) As String
    If IsMissing(retValue) Then retValue = ""
    On Error GoTo ErrorPos
    NVLString = NVLString_(rs.Fields(inFld).Value, retValue)
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLString = retValue
End Function

Public Function NVLBoolean(inFld, Optional retValue As Boolean) As Boolean
    If IsMissing(retValue) Then retValue = False
    On Error GoTo ErrorPos
    NVLBoolean = NVLBoolean_(rs.Fields(inFld).Value, retValue)
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLBoolean = retValue
End Function

Public Function NVLDate(inFld, Optional retValue As Date) As Date
    If IsMissing(retValue) Then retValue = DateSerial(1900, 1, 1)
    On Error GoTo ErrorPos
    NVLDate = NVLDate_(rs.Fields(inFld).Value, retValue)
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLDate = retValue
End Function

Public Function NVLDateF8(inFld) As String
    Dim adate As Date
    On Error GoTo ErrorPos
    adate = NVLDate_(rs.Fields(inFld).Value, DateSerial(1900, 1, 1))
    If adate <= DateSerial(1900, 1, 1) Then
        NVLDateF8 = ""
    Else
        NVLDateF8 = Right("00" & CStr(Day(adate)), 2) & "/" & Right("00" & CStr(Month(adate)), 2) & "/" & Right("0000" & CStr(Year(adate)), 4)
    End If
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLDateF8 = String(10, " ")
End Function

Public Function NVLDateU8(inFld) As String
    Dim adate As Date
    On Error GoTo ErrorPos
    adate = NVLDate_(rs.Fields(inFld).Value, DateSerial(1900, 1, 1))
    If adate <= DateSerial(1900, 1, 1) Then
        NVLDateU8 = ""
    Else
        NVLDateU8 = Right("00" & CStr(Day(adate)), 2) & Right("00" & CStr(Month(adate)), 2) & Right("0000" & CStr(Year(adate)), 4)
    End If
    Exit Function
ErrorPos:
    MsgBox "Το Πεδίο " & inFld & " δεν υπάρχει..."
    NVLDateU8 = String(8, " ")
End Function



