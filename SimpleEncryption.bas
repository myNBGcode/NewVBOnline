Attribute VB_Name = "SimpleEncryption"
Option Explicit

Public Function SimpleEncrypt(strtoencrypt As String, key As String) As String

    Dim s2 As Integer, i As Integer
    Dim tmp As String
    Dim c As String * 1
    
    s2 = 0
    For i = 1 To Len(key)
        c = Mid(key, i, 1)
        s2 = s2 + Asc(c)
    Next i
    
    tmp = ""
    For i = 1 To Len(strtoencrypt)
        c = Mid(strtoencrypt, i, 1)
        tmp = tmp & Chr((Asc(c) Xor (s2 * i)) Mod 256)
    Next i
    
    SimpleEncrypt = tmp
End Function

Public Function SimpleDecrypt(strtodecrypt As String, key As String) As String

    Dim s2 As Integer, i As Integer
    Dim tmp As String
    Dim c As String * 1
    
    s2 = 0
    For i = 1 To Len(key)
        c = Mid(key, i, 1)
        s2 = s2 + Asc(c)
    Next i
    
    tmp = ""
    For i = 1 To Len(strtodecrypt)
        c = Mid(strtodecrypt, i, 1)
        tmp = tmp & Chr((Asc(c) Xor (s2 * i)) Mod 256)
    Next i
    
    SimpleDecrypt = tmp
End Function
