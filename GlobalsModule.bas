Attribute VB_Name = "GlobalsModule"
Option Explicit

Public ReadDir As String
Public WorkEnvironment_ As String ' IRISEDUC για εκπαιδευτικό - IRISPROD για παραγωγή

Public Function gFormat_(FormatString, inParams)
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
    Dim CopyString As String, aSizeString As String, allSize As Integer, decimalsize As Integer, paramNo As Integer, NewPart As String, ResultString As String, astr As String
    Dim apctpos As Integer, bpctpos As Integer, apointpos As Integer, asignflag As Boolean, aLeadZero As Boolean, aType As String
    paramNo = 0: CopyString = FormatString: ResultString = ""
    
    Do
    
        apctpos = InStr(1, CopyString, "%", vbTextCompare)
        If (apctpos > 0 And apctpos < Len(CopyString)) Then bpctpos = InStr(apctpos + 1, CopyString, "%", vbTextCompare) Else bpctpos = 0
        If apctpos * bpctpos > 0 And bpctpos > apctpos + 3 Then
            asignflag = (Mid(CopyString, apctpos + 1, 1) = "-")
            If asignflag And bpctpos <= apctpos + 4 Then
                apctpos = 0: bpctpos = 0
            Else
                aType = UCase(Mid(CopyString, bpctpos - 2, 2))
                aSizeString = Mid(CopyString, apctpos + IIf(asignflag, 2, 1), bpctpos - apctpos - 1 - 2 - IIf(asignflag, 1, 0))
                apointpos = InStr(1, aSizeString, ".")
                If apointpos > 0 And (apointpos = 1 Or apointpos = Len(aSizeString)) Then
                    apctpos = 0: bpctpos = 0
                Else
                    If aType <> "FS" Then
                        aLeadZero = (Left(aSizeString, 1) = "0")
                        If apointpos = 0 Then allSize = CInt(aSizeString) Else allSize = CInt(Left(aSizeString, apointpos - 1))
                        If apointpos = 0 Then decimalsize = 0 Else decimalsize = CInt(Right(aSizeString, Len(aSizeString) - apointpos))
                    End If
                    Select Case aType
                        Case "FD": NewPart = Right(String(allSize, IIf(aLeadZero, "0", " ")) & FormatNumber(inParams(paramNo), decimalsize), allSize)
                        Case "ST":
                                    If IsObject(inParams(paramNo)) Then
                                        If inParams(paramNo) Is Nothing Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    Else
                                        If IsEmpty(inParams(paramNo)) Then NewPart = String(allSize, " ") Else NewPart = Left(inParams(paramNo) & String(allSize, " "), allSize)
                                    End If
                        Case "FS": NewPart = Left(Format(inParams(paramNo), aSizeString) & String(Len(aSizeString), " "), Len(aSizeString))
                    End Select
                    
                    If apctpos > 1 Then astr = Left(CopyString, apctpos - 1) Else astr = ""
                    ResultString = ResultString & astr & NewPart
                    
                    If bpctpos < Len(CopyString) Then CopyString = Right(CopyString, Len(CopyString) - bpctpos) Else CopyString = ""
                    paramNo = paramNo + 1
                End If
            End If
        ElseIf apctpos * bpctpos > 0 And bpctpos = apctpos + 1 Then
            If apctpos > 1 Then astr = Left(CopyString, apctpos - 1) Else astr = ""
            ResultString = ResultString & astr & "%"
            If bpctpos < Len(CopyString) Then CopyString = Right(CopyString, Len(CopyString) - bpctpos) Else CopyString = ""
        Else
            apctpos = 0: bpctpos = 0
        End If
    
    Loop While (Len(CopyString) > 0 And apctpos * bpctpos > 0)
    
    gFormat_ = ResultString
End Function

