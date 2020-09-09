VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComAreaField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public owner As CComAreaField

Public Level As Integer
Public name As String
Public occurs As Integer
Public occurs1flag As Boolean
Public datatype As String
Public length As Integer
Public pic As String
Public init
Public endofstruct As Boolean
Public DataLength As Long

Public fields As Collection

Public Sub ResolveDataLength()
    If fields Is Nothing Then
        If datatype = "CHAR" Then
            DataLength = length
        ElseIf datatype = "FIXED DEC" Then
            DataLength = (length + 1) Mod 2 + (length + 1) \ 2
        ElseIf datatype = "FIXED DECIMAL" Then
            DataLength = (length + 1) Mod 2 + (length + 1) \ 2
        ElseIf datatype = "DEC FIXED" Then
            DataLength = (length + 1) Mod 2 + (length + 1) \ 2
        ElseIf datatype = "DECIMAL FIXED" Then
            DataLength = (length + 1) Mod 2 + (length + 1) \ 2
        ElseIf datatype = "BIN FIXED" Then
            If length = 31 Then
                DataLength = 4
            ElseIf length = 15 Then
                DataLength = 2:
            Else
                Err.Raise 1300, "cComAreaTokenizer", "Λάθος μέγεθος " & length & " για το τύπο FIXED BIN"
            End If
        ElseIf datatype = "BINARY FIXED" Then
            If length = 31 Then
                DataLength = 4
            ElseIf length = 15 Then
                DataLength = 2:
            Else
                Err.Raise 1300, "cComAreaTokenizer", "Λάθος μέγεθος " & length & " για το τύπο FIXED BIN"
            End If
        ElseIf datatype = "FIXED BIN" Then
            If length = 31 Then
                DataLength = 4
            ElseIf length = 15 Then
                DataLength = 2:
            Else
                Err.Raise 1300, "cComAreaTokenizer", "Λάθος μέγεθος " & length & " για το τύπο FIXED BIN"
            End If
        ElseIf datatype = "FIXED BINARY" Then
            If length = 31 Then
                DataLength = 4
            ElseIf length = 15 Then
                DataLength = 2:
            Else
                Err.Raise 1300, "cComAreaTokenizer", "Λάθος μέγεθος " & length & " για το τύπο FIXED BIN"
            End If
        End If
    Else
        Dim afield As CComAreaField
        DataLength = 0
        For Each afield In fields
            afield.ResolveDataLength
            DataLength = DataLength + (afield.DataLength * afield.occurs)
        Next afield
    End If
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub

Public Sub CleanUp()
    If fields Is Nothing Then
    Else
        While fields.count > 0
            fields(1).CleanUp
            fields.Remove 1
        Wend
    End If
    Set fields = Nothing
    Set owner = Nothing
End Sub
