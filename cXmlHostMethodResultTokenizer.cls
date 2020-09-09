VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlHostMethodResultTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public resultList As Collection

Private Result As String


Public Sub Prepare(inResultList As Collection)
    Set resultList = inResultList
    
    Dim i As Integer
    Result = ""
    For i = 1 To resultList.count
        Result = Result & resultList(i)
    Next i
End Sub

Public Function GetToken(Length As Integer) As String
    If Len(Result) > Length Then
        GetToken = Left(Result, Length)
        Result = Right(Result, Len(Result) - Length)
    Else
        GetToken = Result
        Result = ""
    End If
End Function

Public Function TestToken(Length As Integer) As String
    If Len(Result) > Length Then
        TestToken = Left(Result, Length)
    Else
        TestToken = Result
    End If
End Function

Public Function GetAll() As String
    GetAll = Result
    Result = ""
End Function

Public Function TestAll() As String
    TestAll = Result
End Function

