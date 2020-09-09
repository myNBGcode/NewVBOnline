Attribute VB_Name = "BufferMdl"
Option Explicit
Public Const ptChar = 1
Public Const ptVarchar = 2
Public Const ptInt = 3
Public Const ptSmall = 4
Public Const ptDecimal = 5
Public Const ptDate = 6
Public Const ptTime = 7
Public Const ptTimeStamp = 8
Public Const ptPicture = 9
Public Const ptStruct = 99


'Public BufferList() As Buffer, BufferNum As Integer

Public Function AsciiToEbcdic_(inputStr As String) As String
Dim InputAscii As String, OutputStr As String

InputAscii = inputStr & Chr$(0)
OutputStr = GKTranslate(InputAscii, EBCDIC_CP_STRING, ASCII_CP_STRING)
AsciiToEbcdic_ = Left(OutputStr, Len(inputStr))
End Function

Public Function EbcdicToAscii_(inputStr As String) As String
Dim InputAscii As String, OutputStr As String
inputStr = Replace(inputStr, Chr(0), Chr(64))
InputAscii = inputStr & Chr$(0)
OutputStr = GKTranslate(InputAscii, ASCII_CP_STRING, EBCDIC_CP_STRING)
EbcdicToAscii_ = Left(OutputStr, Len(inputStr))
End Function

Public Function DecimalToHPS_(invalue As Double, Digits As Long) As String

Dim astr As String, OutStr As String, i As Integer, k As Integer

astr = format(Abs(invalue), Left("00000000000000000000", Digits))
If Len(astr) Mod 2 = 0 Then astr = "0" & astr
For i = 1 To Len(astr) - 1 Step 2
    k = CInt(Mid(astr, i, 1)) * 16 + CInt(Mid(astr, i + 1, 1))
    OutStr = OutStr & Chr(k)
Next i
OutStr = OutStr & Chr(CInt(Right(astr, 1)) * 16 + IIf(invalue >= 0, 12, 13))

DecimalToHPS_ = OutStr
End Function

Public Function HPSToDecimal_(invalue As String) As Double

Dim astr As String, bstr As String, OutStr As String, i As Integer, pos As Integer
On Error GoTo ErrorReport
pos = 1
HPSToDecimal_ = 0
pos = 2
astr = invalue
pos = 3
OutStr = ""
pos = 4
For i = 1 To Len(astr) - 1 Step 1
pos = 5
    OutStr = OutStr & Trim(CStr(CLng(Asc(Mid(astr, i, 1)) \ 16))) & Trim(CStr(CLng(Asc(Mid(astr, i, 1)) Mod 16)))
Next i
pos = 6
OutStr = OutStr & Trim(CStr(CLng(Asc(Right(astr, 1)) \ 16)))
pos = 7
OutStr = IIf(Asc(Right(astr, 1)) Mod 16 = 13, "-", "") & OutStr
pos = 8
HPSToDecimal_ = CDbl(OutStr)
Exit Function
ErrorReport:
    astr = ""
    If invalue <> "" Then
        For i = 1 To Len(invalue) Step 1
            astr = astr & Asc(Mid(invalue, i, 1)) & ","
        Next i
    End If
    bstr = ""
    If OutStr <> "" Then
        For i = 1 To Len(OutStr) Step 1
            bstr = bstr & Asc(Mid(OutStr, i, 1)) & ","
        Next i
    End If
    LogMsgbox "Λαθος Κατα τη Μετατροπή του Asc:" & astr & " σε FD Θεση " & CStr(pos) & " Output:|" & OutStr & "|:" & bstr
End Function

Public Function IntToHps_(ByVal InputInt As Long) As String
    Dim i As Integer, StrOUT As String, Step As Long, Resd As Long
    StrOUT = ""
    Step = InputInt
    For i = 1 To 4
        Resd = Step Mod 256
        Step = Step \ 256
        StrOUT = Chr$(Resd) & StrOUT
    Next i
    IntToHps_ = StrOUT
End Function

Public Function HpsToInt_(ByVal InputInt As String) As Long
    Dim i As Integer, ValOut As Long
    ValOut = 0
    For i = 1 To 4
        ValOut = ValOut * 256 + Asc(Mid(InputInt, i, 1))
    Next i
    HpsToInt_ = ValOut
End Function

Public Function VarCharToHps_(ByVal InputVarChar As String, ByVal InputLen As Integer) As String
     Dim StrL As String
     Dim StrOUT As String
     
     StrL = SmallToHps_(Len(Trim(InputVarChar)))
     
     StrL = SmallToHps_(InputLen)
     VarCharToHps_ = StrL & AsciiToEbcdic_(Left(InputVarChar & String(InputLen, " "), InputLen))
End Function

Public Function SmallToHps_(ByVal InputSmall As Long) As String
    Dim i As Integer, StrOUT As String, Step As Long, Resd As Long
    StrOUT = ""
    Step = IIf(InputSmall < 0, 65536 + InputSmall, InputSmall)
    For i = 1 To 2
        Resd = Step Mod 256
        Step = Step \ 256
        StrOUT = Chr$(Resd) & StrOUT
    Next i
    SmallToHps_ = StrOUT
End Function

Public Function HpsToSmall_(ByVal InputInt As String) As Long
    Dim i As Integer, ValOut As Long
    ValOut = 0
    For i = 1 To 2
        ValOut = ValOut * 256 + Asc(Mid(InputInt, i, 1))
    Next i
    HpsToSmall_ = IIf(ValOut > 32768, -(65536 - ValOut), ValOut)
End Function

