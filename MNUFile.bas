Attribute VB_Name = "MNUFile"
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long


Public LastTRNCode As String
Public LastTRNNum As Integer

'Public LastChief As String
'Public LastManager As String
'
Public Log_db As ADODB.Connection
'Public Log_cmd As ADODB.Command

Public Function Cheque_Amount_str_(ptxtAmount As String, _
                            Optional pBolLang As Variant) As String
' H Function Amount_str μαζί με την strnum δίνουν το ολογράφως
' του ποσού.
' Παράμετροι: το ποσό υποχρεωτικό σαν  string και
'                           προαιρετικό μία True False για την γλώσσα του νομίσματος
'                           True  η δραχμή και όλα τα θηλυκά νομίσματα
'                           False όλα τα ουδέτερα νομίσματα
' Το μήκος του ποσού θεωρείται 17 χαρακτήρες με δύο δεκαδικά

Dim MstrAmount As String
Dim MArrAmount(6) As String
Dim mintLen As Integer
Dim MstrAnalAmount As String
Dim inti As Integer

 If IsMissing(pBolLang) Then
   pBolLang = True
 End If

MstrAmount = Unformat_num(ptxtAmount)

If Val(MstrAmount) = 0 Then
    Cheque_Amount_str_ = "ΜΗΔΕΝ"
    Exit Function
End If
mintLen = Len(MstrAmount)
MstrAmount = Space(17 - Len(MstrAmount)) + MstrAmount

' χωρισμός ανά τριάδες και 2 δεκαδικά
For inti = 1 To 6
    If inti < 6 Then
        MArrAmount(inti) = Mid(MstrAmount, (inti - 1) * 3 + 1, 3)
    Else
        MArrAmount(inti) = "0" + Mid(MstrAmount, 16, 2)
    End If
Next
MstrAnalAmount = " "
For inti = 1 To 6
    'If MArrAmount(inti) <> Space(3) And MArrAmount(inti) <> "000" Then
    If Val(MArrAmount(inti)) <> 0 Then
        Select Case inti
            Case 1
                    MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 1, pBolLang) + "ΤΡΙΣ "
            Case 2
                   MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 2, pBolLang) + "ΔΙΣ "
            Case 3
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 3, pBolLang) + "ΕΚΑΤΟΜ "
            Case 4
                    If Val(MArrAmount(inti)) = 1 Then
                        If pBolLang Then
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙΕΣ "
                        Else
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙA "
                        End If
                    Else
                        MstrAnalAmount = MstrAnalAmount + _
                                strnum(MArrAmount(inti), 4, pBolLang) + "ΧΙΛΙΑΔΕΣ "
                     End If
            Case 5
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 5, pBolLang)
            Case 6
                     MstrAnalAmount = MstrAnalAmount + "ΚΑΙ " + CStr(CLng("0" & MArrAmount(inti))) & " / 100"
        End Select
    End If
Next
Cheque_Amount_str_ = LTrim(MstrAnalAmount)
End Function

Public Function Amount_str(ptxtAmount As String, _
                            Optional pBolLang As Variant) As String
' H Function Amount_str μαζί με την strnum δίνουν το ολογράφως
' του ποσού.
' Παράμετροι: το ποσό υποχρεωτικό σαν  string και
'                           προαιρετικό μία True False για την γλώσσα του νομίσματος
'                           True  η δραχμή και όλα τα θηλυκά νομίσματα
'                           False όλα τα ουδέτερα νομίσματα
' Το μήκος του ποσού θεωρείται 17 χαρακτήρες με δύο δεκαδικά

Dim MstrAmount As String
Dim MArrAmount(6) As String
Dim mintLen As Integer
Dim MstrAnalAmount As String
Dim inti As Integer

 If IsMissing(pBolLang) Then
   pBolLang = True
 End If

Dim aval As Double

MstrAmount = Unformat_num(ptxtAmount)

If Val(MstrAmount) = 0 Then
    Amount_str = "ΜΗΔΕΝ"
    Exit Function
End If
mintLen = Len(MstrAmount)
MstrAmount = Space(17 - Len(MstrAmount)) + MstrAmount

' χωρισμός ανά τριάδες και 2 δεκαδικά
For inti = 1 To 6
    If inti < 6 Then
        MArrAmount(inti) = Mid(MstrAmount, (inti - 1) * 3 + 1, 3)
    Else
        MArrAmount(inti) = "0" + Mid(MstrAmount, 16, 2)
    End If
Next
MstrAnalAmount = " "
For inti = 1 To 6
    'If MArrAmount(inti) <> Space(3) And MArrAmount(inti) <> "000" Then
    If Val(MArrAmount(inti)) <> 0 Then
        Select Case inti
            Case 1
                    MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 1, pBolLang) + "ΤΡΙΣ "
            Case 2
                   MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 2, pBolLang) + "ΔΙΣ "
            Case 3
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 3, pBolLang) + "ΕΚΑΤΟΜ "
            Case 4
                    If Val(MArrAmount(inti)) = 1 Then
                        If pBolLang Then
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙΕΣ "
                        Else
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙA "
                        End If
                    Else
                        MstrAnalAmount = MstrAnalAmount + _
                                strnum(MArrAmount(inti), 4, pBolLang) + "ΧΙΛΙΑΔΕΣ "
                     End If
            Case 5
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 5, pBolLang)
            Case 6
                     MstrAnalAmount = MstrAnalAmount + "ΚΑΙ " + _
                            strnum(MArrAmount(inti), 6, pBolLang) + " ΕΚΑΤΟΣΤΑ"
        End Select
    End If
Next
Amount_str = LTrim(MstrAnalAmount)
End Function


Public Function Amount_Str2002(ptxtAmount As String, _
                            Optional pBolLang As Variant, Optional pCurFlag As Variant) As String
' H Function Amount_str μαζί με την strnum δίνουν το ολογράφως
' του ποσού.
' Παράμετροι: το ποσό υποχρεωτικό σαν  string και
'                           προαιρετικό μία True False για την γλώσσα του νομίσματος
'                           True  η δραχμή και όλα τα θηλυκά νομίσματα
'                           False όλα τα ουδέτερα νομίσματα
' Το μήκος του ποσού θεωρείται 17 χαρακτήρες με δύο δεκαδικά

Dim MstrAmount As String
Dim MArrAmount(6) As String
Dim mintLen As Integer
Dim MstrAnalAmount As String
Dim inti As Integer


 If IsMissing(pBolLang) Then pBolLang = (cVersion < 20020101)
 If IsMissing(pCurFlag) Then pCurFlag = True
Dim aval As Double

MstrAmount = Unformat_num(ptxtAmount)

'If Val(MstrAmount) = 0 Then
'    Amount_Str2002 = "ΜΗΔΕΝ" & IIf(cVersion >= 20020101, "ΕΥΡΩ", " ΔΡΧ.")
'    Exit Function
'End If

If Val(MstrAmount) = 0 Then
    If pCurFlag Then
        Amount_Str2002 = "ΜΗΔΕΝ" & IIf(cVersion >= 20020101, " ΕΥΡΩ", " ΔΡΧ.")
    Else
        Amount_Str2002 = "ΜΗΔΕΝ"
    End If
    Exit Function
End If

mintLen = Len(MstrAmount)
MstrAmount = Space(17 - Len(MstrAmount)) + MstrAmount

' χωρισμός ανά τριάδες και 2 δεκαδικά
For inti = 1 To 6
    If inti < 6 Then
        MArrAmount(inti) = Mid(MstrAmount, (inti - 1) * 3 + 1, 3)
    Else
        MArrAmount(inti) = "0" + Mid(MstrAmount, 16, 2)
    End If
Next
MstrAnalAmount = " "
For inti = 1 To 6
    'If MArrAmount(inti) <> Space(3) And MArrAmount(inti) <> "000" Then
    If Val(MArrAmount(inti)) <> 0 Then
        Select Case inti
            Case 1
                    MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 1, pBolLang) + "ΤΡΙΣ "
            Case 2
                   MstrAnalAmount = MstrAnalAmount + _
                        strnum(MArrAmount(inti), 2, pBolLang) + "ΔΙΣ "
            Case 3
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 3, pBolLang) + "ΕΚΑΤΟΜ "
            Case 4
                    If Val(MArrAmount(inti)) = 1 Then
                        If pBolLang Then
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙΕΣ "
                        Else
                            MstrAnalAmount = MstrAnalAmount + "ΧΙΛΙA "
                        End If
                    Else
                        MstrAnalAmount = MstrAnalAmount + _
                                strnum(MArrAmount(inti), 4, pBolLang) + "ΧΙΛΙΑΔΕΣ "
                     End If
            Case 5
                    MstrAnalAmount = MstrAnalAmount + _
                            strnum(MArrAmount(inti), 5, pBolLang)
            Case 6
                     MstrAnalAmount = MstrAnalAmount + "ΚΑΙ " + _
                            strnum(MArrAmount(inti), 6, pBolLang) + " ΕΚΑΤΟΣΤΑ"
        End Select
    End If
Next
If pCurFlag Then
   Amount_Str2002 = LTrim(MstrAnalAmount) & IIf(cVersion >= 20020101, " ΕΥΡΩ", " ΔΡΧ.")
Else
   Amount_Str2002 = LTrim(MstrAnalAmount)
End If
End Function

Function strnum(ptxtAmount As String, intpos As Integer, _
                                pBolLang As Variant) As String
Dim strAmount As String
Dim inti As Integer
Dim intj As Integer
Dim strEkat As Variant
Dim strDeka As Variant
Dim strMona As Variant

strEkat = Array(" ", "ΕΚΑΤΟ", "ΔΙΑΚΟΣΙ", "ΤΡΙΑΚΟΣΙ", "ΤΕΤΡΑΚΟΣΙ", "ΠΕΝΤΑΚΟΣΙ", "ΕΞΑΚΟΣΙ", "ΕΠΤΑΚΟΣΙ", "ΟΚΤΑΚΟΣΙ", "ΕΝΝΙΑΚΟΣΙ")
strDeka = Array(" ", "ΔΕΚΑ ", "ΕΙΚΟΣΙ ", "ΤΡΙΑΝΤΑ ", "ΣΑΡΑΝΤΑ ", "ΠΕΝΗΝΤΑ ", "ΕΞΗΝΤΑ ", "ΕΒΔΟΜΗΝΤΑ ", "ΟΓΔΟΝΤΑ ", "ΕΝΕΝΗΝΤΑ ")
strMona = Array(" ", "ΕΝΑ ", "ΔΥΟ ", "ΤΡΙΑ ", "ΤΕΣΣΕΡΑ ", "ΠΕΝΤΕ ", "ΕΞΙ ", "ΕΠΤΑ ", "ΟΚΤΩ ", "ΕΝΝΕΑ ")

strAmount = " "
For inti = 1 To 3
    intj = Val(Mid(ptxtAmount, inti, 1))
    If intpos = 6 And inti = 2 And intj = 0 Then
        strAmount = " ΜΗΔΕΝ "
    ElseIf intj <> 0 Then
        Select Case inti
            Case 1
               strAmount = strEkat(intj)
               If intj > 1 Then
                   If (pBolLang = False And intpos = 5) Or intpos < 4 Then
                       strAmount = strAmount + "A "
                    Else
                       strAmount = strAmount + "ΕΣ "
                    End If
               Else
                    If Mid(ptxtAmount, 2, 2) = "00" Then
                        strAmount = strAmount + " "
                    Else
                        strAmount = strAmount + "Ν "
                    End If
               End If
            Case 2
                If intj = 1 And (Mid(ptxtAmount, 3, 1) = "1" Or Mid(ptxtAmount, 3, 1) = "2") Then
                    If Mid(ptxtAmount, 3, 1) = "1" Then
                        strAmount = strAmount + " ΕΝΤΕΚΑ "
                    Else
                        strAmount = strAmount + " ΔΩΔΕΚΑ "
                    End If
                    strnum = strAmount
                    Exit Function
                End If
                strAmount = strAmount + strDeka(intj)
            Case 3
                    If intpos = 4 Or intpos = 5 Then
                        Select Case intj
                            Case 1
                                  If intpos = 5 And pBolLang = False Then
                                      strAmount = strAmount + strMona(intj) '+ " "
                                  Else
                                      strAmount = strAmount + "ΜΙΑ "
                                  End If
                            Case 3
                                If intpos = 5 And pBolLang = False Then
                                      strAmount = strAmount + strMona(intj) '+ " "
                                Else
                                     strAmount = strAmount + "ΤΡΕΙΣ "
                                End If
                            Case 4
                                If intpos = 5 And pBolLang = False Then
                                    strAmount = strAmount + strMona(intj) '+ " "
                                Else
                                     strAmount = strAmount + "TΕΣΣΕΡΕΙΣ "
                                End If
                            Case Else
                                strAmount = strAmount + strMona(intj) '+ " "
                        End Select
                    Else
                        strAmount = strAmount + strMona(intj) '+ " "
                    End If
         End Select
    End If
Next
strnum = strAmount
End Function

Public Function EmbedValues(owner As Form, aLine As String, aPageNo As Integer) As String
Dim aPos As Integer, oldLine As String
    
    oldLine = ""
    While oldLine <> aLine
    
        oldLine = aLine
        aPos = InStr(aLine, "%dl") 'date long format
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & format(Date, "Long Date") & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%ds") 'date short format
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & format(Date, "dd/mm/yyyy") & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%tl") 'time long format
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & format(Time, "Long Time") & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%ts") 'time short format
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & format(Time, "HH:MM:SS") & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%tm") 'terminal
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & cTERMINALID & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%br") 'BRANCH
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & cBRANCHName & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%pg") 'page
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & aPageNo & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
        aPos = InStr(aLine, "%tn") 'transaction number
        If aPos > 0 Then
            aLine = Left(aLine, aPos - 1) & CStr(cTRNNum) & _
                    Right(aLine, Len(aLine) - aPos - 2)
        End If
    
Dim alinepart As String, aFldName As String
Dim bpos As Integer
        
        aPos = InStr(aLine, "%f") 'fieldname
        If aPos > 0 Then
            alinepart = Trim(Right(aLine, Len(aLine) - aPos - 1))
            bpos = InStr(alinepart, " ")
            If bpos > 0 Then
                aFldName = Trim(Left(alinepart, bpos))
                aLine = Left(aLine, aPos - 1) & _
                    owner.GetFormatedFld(aFldName) & _
                        Right(alinepart, Len(alinepart) - bpos + 1)
            Else
                If alinepart <> "" Then
                    aFldName = alinepart
                    aLine = Left(aLine, aPos - 1) & _
                        owner.GetFormatedFld(aFldName)
                Else
                    aLine = Left(aLine, aPos - 1) & _
                        Right(aLine, Len(aLine) - aPos - 1)
                End If
            End If
        End If
    Wend
    EmbedValues = aLine
End Function

Public Function ExtractFirstLineFromString(inString As String) As String
Dim aPos As Integer, aLine As String
    
    aPos = InStr(1, inString, vbLf, vbBinaryCompare)
    If aPos > 0 Then
        If aPos > 2 Then
            aLine = Left(inString, aPos - 2)
        Else: aLine = "": End If
        If aPos < Len(inString) - 1 Then
            inString = Right(inString, Len(inString) - aPos)
        Else: inString = "": End If
    Else
        aLine = inString
        inString = ""
    End If
    ExtractFirstLineFromString = aLine
End Function

Public Function eJournalClearString(astr As String) As String
'Dim I As Integer, bstr As String
'    I = Len(astr)
'    bstr = ""
'    While I > 0
'        If Mid(astr, I, 1) = "'" Then
'            bstr = "΄" & bstr
'        ElseIf Mid(astr, I, 1) = """" Then
'            bstr = " " & bstr
'        ElseIf Asc(Mid(astr, I, 1)) = 127 Then
'            bstr = "΄" & bstr
'        ElseIf Asc(Mid(astr, I, 1)) <> 0 Then
'            bstr = Mid(astr, I, 1) & bstr
'        End If
'        I = I - 1
'    Wend
'    ClearString = bstr

Dim src As String, dst As String, aLen As Integer
dst = "΄΄΄" & Chr$(0)
src = "'""" & Chr$(127) & Chr$(0)
aLen = Len(astr)
eJournalClearString = Left(GKTranslate(astr & Chr$(0), src, dst), aLen)

End Function

Public Sub CenterFormOnScreen(pFormToCenter As Form)
  pFormToCenter.Move GenWorkForm.Left + (GenWorkForm.width - pFormToCenter.width) / 2, GenWorkForm.Top + (GenWorkForm.height - pFormToCenter.height) / 2
End Sub

Public Function Unformat_num(PStrTxtnum As String, _
                             Optional PIntDecPos As Variant, _
                             Optional PStrDecChr As Variant) _
                             As String

' Η Function Unformat_num δέχεται ένα αριθμητικό
' formated string (π.χ. 123.456.789,00) και
' επιστρέφει ένα unformated αριθμητικό string
' ( π.χ. 123456789.00)
'
' Παράμετροι :
' PStrTxtnum το formated αριθμητικό string
' PIntDecPos προαιρετικά ένας ακεραιος που δηλώνει
' δεκαδικά ψηφία, default 0
' PStrDecChr προαιρετικά ο χαρακτήρας που θα
' χρησιμοποιηθεί για υποδιαστολή, default "."
'
' π.χ.
'
' txtDate.Text = "15/12/1995"
' Unformat_num(txtDATE.Text)--> 15121995
'
' txtAmount.Text = "123.456,00"
' Unformat_num(txtAmount.Text)--> 12345600
' Unformat_num(txtAmount.Text,2)--> 123456.00
' Unformat_num(txtAmount.Text,2,",")--> 123456,00

    Dim minti As Integer
    Dim mintLen As Integer
    Dim MStrS As String

    MStrS = ""
    mintLen = Len(PStrTxtnum)
    If IsMissing(PIntDecPos) Then
       PIntDecPos = 0
    End If

    If IsMissing(PStrDecChr) Then
       PStrDecChr = "."
    End If

    For minti = 1 To mintLen
        If Mid(PStrTxtnum, minti, 1) Like "#" Then
           MStrS = MStrS + Mid(PStrTxtnum, minti, 1)
        End If
    Next
    If PIntDecPos > 0 Then
       Unformat_num = Mid(MStrS, 1, (Len(MStrS) - PIntDecPos)) + _
                      PStrDecChr + Right(MStrS, PIntDecPos)
    Else
       Unformat_num = MStrS
    End If
End Function

'Public Function HasKeyChanged(nowchiefuname As String, nowmanageruname As String) As Boolean
'    HasKeyChanged = False
'    If LastChief <> nowchiefuname Then HasKeyChanged = True: Exit Function
'    If LastManager <> nowmanageruname Then HasKeyChanged = True: Exit Function
'End Function
'
'Public Function SetLastKey(nowchiefuname As String, nowmanageruname As String) As String
'    If nowchiefuname = "" Then LastChief = "": SetLastKey = "": Exit Function
'    If LastChief <> nowchiefuname Then LastChief = nowchiefuname: SetLastKey = "Εγκριση: CT:" & LastChief: Exit Function
'    If nowmanageruname = "" Then LastManager = "": SetLastKey = "": Exit Function
'    If LastManager <> nowmanageruname Then LastManager = nowmanageruname: SetLastKey = "Εγκριση: M :" & LastManager: Exit Function
'End Function
