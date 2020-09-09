Attribute VB_Name = "ChkDigits"
Public Function CalcCd1(Acc As String, Digits As Integer) As Integer
'Πρώτο Check Digit λογαριασμού
    Dim total, i, Rm As Integer
  
    total = 0
    For i = Digits To 1 Step -1
        total = total + Val(Mid(Acc, i, 1)) * (Digits + 2 - i)
    Next
    CalcCd1 = 11 - (total Mod 11)
    CalcCd1 = (1 - (CalcCd1 \ 10)) * CalcCd1
End Function

Public Function CalcCd2(aAcc10 As String) As Integer
'Δεύτερο Check Digit λογαριασμού

    Dim i, total, Cd2Cor As Integer
    
    total = 0
    For i = 2 To 10 Step 2
        total = total + (2 * Val(Mid(aAcc10, i, 1)) Mod 10) + _
                        (2 * Val(Mid(aAcc10, i, 1)) \ 10)
    Next
    
    For i = 1 To 9 Step 2
        total = total + Val(Mid(aAcc10, i, 1))
    Next
    
    CalcCd2 = total Mod 10
    If CalcCd2 <> 0 Then CalcCd2 = 10 - CalcCd2
    
End Function

Public Function CalcChequeCd(Cheque As Integer) As Integer
  Dim Rm As Integer
  Rm = (Cheque \ 10) Mod 11
  If Rm = 10 Then Rm = 0
  CalcChequeCd = Rm
End Function

Public Function CalcSAccCd(SAcc As String) As Integer
'  If Len(SAcc) <= 3 Then _
'    SAcc = "073" & StrPad(SAcc, 3, "0", "L")
    
  CalcSAccCd = CalcCd1(SAcc, Len(SAcc))
End Function

Public Function ChkDocument(inNum As String) As Boolean
'Συναλλαγή 1020 Πεδίο Έγγραφο / Επιταγή
    Dim copyNum As String, part1 As String
    copyNum = StrPad_(inNum, 11, "0", "L")
    ChkDocument = _
        ((CalcCd1(Left(copyNum, 7), 7) = CInt(Mid(copyNum, 8, 1))) _
        And (CLng(Left(copyNum, 7)) > 0))
End Function

Public Function ChkCard(inNum As String) As Boolean
'Συναλλαγή 4010 Πεδίο Αριθμός Κάρτας
    Dim copyNum As String, Num1 As Integer, Sum1 As Integer, i As Integer
    copyNum = StrPad_(inNum, 16, "0", "L")
    
    If CDbl(copyNum) = 0 Then
        ChkCard = True
    ElseIf CDbl(Left(copyNum, 8)) = 0 Then
        ChkCard = (CalcCd1(Mid(copyNum, 9, 7), 7) = CInt(Right(copyNum, 1)))
    ElseIf CDbl(Left(copyNum, 3)) = 0 Then
        ChkCard = True
    Else
        Sum1 = 0
        For i = 1 To 15 Step 2
            Num1 = CInt(Mid(copyNum, i, 1) * 2)
            Sum1 = Sum1 + Num1 \ 10 + Num1 Mod 10
        Next
        For i = 2 To 14 Step 2
            Sum1 = Sum1 + CInt(Mid(copyNum, i, 1) * 2)
        Next i
        ChkCard = (((10 - (Sum1 Mod 10)) Mod 10) = Right(copyNum, 1))
    End If
End Function

Public Function ChkTaxID(inNum As String) As Boolean
'Συναλλαγή 4150 Πεδίο ΑΦΜ - Συναλλαγή 4350
    Dim copyNum As String, Sum1 As Long, i As Long
    copyNum = StrPad_(inNum, 10, "0", "L")
    For i = 1 To 9 Step 1
        Sum1 = Sum1 + CInt(Mid(copyNum, i, 1)) * (2 ^ (10 - i))
    Next
    ChkTaxID = ((Sum1 Mod 11) = CInt(Right(copyNum, 1))) Or ((Sum1 Mod 11) = 10) 'να ελεγχθει η περιπτωση sum1 mod 11 = 10
End Function

Public Function ChkRet(inNum As String) As Boolean
'Συναλλαγή 4170 Πεδίο ΑΜ Συνταξιούχου ' Δεν γινεται έλεγχος για τους ΟΕΚ
    Dim copyNum As String
    copyNum = StrPad_(inNum, 12, "0", "L")
    
    ChkRet = (((CDbl(Left(copyNum, 11)) Mod 11) Mod 10) = CInt(Right(copyNum, 1))) _
        Or (Left(copyNum, 2) = "64" And Right(copyNum, 1) = "0")
End Function

Public Function ChkMet(inNum As String) As Boolean
'Συναλλαγή 4193 Πεδίο ΑΜ Μετοχου
    Dim copyNum As String
    copyNum = StrPad_(inNum, 9, "0", "L")
    ChkMet = _
        ((CalcCd1(Left(copyNum, 6), 6) = CInt(Mid(copyNum, 7, 1))) _
        And (CLng(Left(copyNum, 6)) > 0))
End Function

Public Function ChkSt(inNum As String) As Boolean
'Συναλλαγή 4350 Πεδίο Αριθμός Ειδοποιητηρίου
    Dim copyNum As String, Weights As String, Sum1 As Long, i As Long
    copyNum = StrPad_(inNum, 12, "0", "L"): Weights = "65432765432"
    Sum1 = 0
    For i = 1 To 11
        Sum1 = CInt(Mid(copyNum, i, 1)) * CInt(Mid(Weights, i, 1))
    Next
    ChkSt = (((11 - (Sum1 Mod 11)) Mod 10) = Right(copyNum, 1))
'αν το copynum(10) = 1 να συμπληρωθεί ΑΦΜ
'αν το copynum(10) = 0 να μην συμπληρωθούν πεδία 32 και 35
End Function
    
Public Function ChkMobile(inNum As String) As Boolean
'Συναλλαγή 4800 Πεδίο Αριθμός Κινητου
    Dim copyNum As String
    copyNum = StrPad_(inNum, 10, "0", "L")
    ChkMobile = (Left(copyNum, 2) = "09")
End Function

Public Function Chk_Xrhmat(pfrmCurrent As Form, pIndex As Integer) As Boolean
'***************************************************

' Η Function Chk_Xrhmat ελέγχει το check digit
' του αριθμού χρηματ. πελάτη, αν είναι σωστό επιστρέφει
' True αλλιώς False
'
' Παράμετροι :
' PStrXrhmat ο δεκαψήφιος αριθμός χρηματ. πελάτη
' PStrDigit   το Check Digit
'
' π.χ.
'
' If Not Chk_Xrhmat(MStrXrhmat, MStrDigit) Then
'    Beep
'    MsgBox "Λανθασμένος Αριθμός χρηματ. πελάτη !!"
' End If
    
    Dim minti As Integer
    Dim MIntJ As Integer
    Dim MIntCd As Integer
    Dim strXrhmat As String
    Dim strDigit As String
    
    Chk_Xrhmat = True
    If ValidationInProgress Then Exit Function
    
    minti = MIntJ = MIntCd = 0
    
    strXrhmat = pfrmCurrent.txtinput(pIndex).Text
    strDigit = Mid(strXrhmat, 10, 1)

    For minti = 10 To 2 Step -1
        MIntJ = MIntJ + 1
        MIntCd = MIntCd + (Val(Mid(strXrhmat, MIntJ, 1)) * minti)
    Next
    MIntCd = MIntCd Mod 11
    If MIntCd = 1 Or MIntCd = 0 Then
       MIntCd = 0
    Else
       MIntCd = 11 - MIntCd
    End If
    
    If MIntCd <> Val(strDigit) Then
        Chk_Xrhmat = False
'        Call FocusWrongInputField(pfrmCurrent, pIndex, "Λανθασμένο Check Digit!!!")
    End If

End Function

Function DoChkBankAccount(inChk As String, inmod As Integer, _
                        inminus As Integer, infactors As Variant, inmod10 As Integer, inmod11 As Integer, _
                        inlowValid As Single, inhiValid As Single, in2digs As Boolean) As Boolean
Dim inCd, corCd As Integer
Dim ntotal As Single

ntotal = 0
inCd = Val(Mid(inChk, 19, 1))

If in2digs = False Then
   For i = 1 To 18
      ntotal = ntotal + Val(Mid(inChk, i, 1)) * infactors(i - 1)
   Next i
Else
   For i = 1 To 18
      ntotal = ntotal + (Val(Mid(inChk, i, 1)) * infactors(i - 1)) \ 10 + _
                        (Val(Mid(inChk, i, 1)) * infactors(i - 1)) Mod 10
   Next i
End If

corCd = ntotal Mod inmod
If inminus = 100 Then
   corCd = (((ntotal \ 10) + 1) * 10) - ntotal
Else
   corCd = Abs(inminus - corCd)
End If

If corCd = 10 Then
   corCd = inmod10
ElseIf corCd = 11 Then
   corCd = inmod11
End If

If inCd = corCd Then
   DoChkBankAccount = True
Else
   DoChkBankAccount = False
End If

End Function

Public Function ChkBankAcount_(inBank As String, inbranch As String, inAcc As String) As Boolean
Dim aflag As Boolean
Dim aBank, aBranch, aAcc As String
Dim atotal As Variant

aBank = StrPad_(inBank, 3, "0")
aBranch = StrPad_(inbranch, 3, "0")
aAcc = StrPad_(inAcc, 13, "0")

Select Case Val(aBank)
Case 12, 13, 16, 17  ' ΕΜΠΟΡΙΚΗ, ΙΟΝΙΚΗ, ΑΤΤΙΚΗΣ, ΠΕΙΡΑΙΩΣ
   If Mid(aAcc, 1, 1) = 9 Then
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
   Else
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
      If Val(aAcc) = 57000007 And Val(aBank) = 13 Then aflag = True
   End If

Case 14  ' ΠΙΣΤΕΩΣ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 16384, 8192, 4096, 0, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)

Case 15  ' ΓΕΝΙΚΗ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 6, 5, 4, 0, 0, 0, 0, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   If Val(aAcc) = 574 Then aflag = True

Case 18  ' ΑΘΗΝΩΝ
   aflag = DoChkBankAccount(aBank + aBranch + "0" + Mid(aAcc, 1, 12), 11, 11, _
                Array(0, 0, 0, 31, 29, 26, 0, 0, 0, 23, 22, 19, 17, 13, 7, 6, 3, 2), 0, 0, 0, 9999999999999#, False)
   If aflag = True Then
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 6, 5, 4, 0, 0, 0, 0, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   End If

Case 19  ' ΚΡΗΤΗΣ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)

Case 20  ' ΕΡΓΑΣΙΑΣ
   If Mid(aAcc, 11, 2) = "99" Then
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 1, 100, _
                 Array(1, 2, 1, 2, 1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 2), 0, 0, 0, 9999999999999#, True)
   Else
      If Mid(aAcc, 7, 2) = "00" And (Mid(aAcc, 9, 2) = "00" Or Mid(aAcc, 9, 2) = "04") Then
         aflag = DoChkBankAccount(aBank + aBranch + "0" + Mid(aAcc, 1, 12), 1, 100, _
                Array(0, 0, 0, 1, 2, 1, 0, 0, 2, 1, 2, 1, 2, 0, 0, 0, 0, 0), 0, 0, 0, 9999999999999#, True)
         If aflag = True Then
            aflag = DoChkBankAccount(aBank + aBranch + aAcc, 1, 100, _
                 Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2, 0), 0, 0, 0, 9999999999999#, True)
         End If
      Else
        aflag = False
      End If
   End If

Case 21  ' ΜΑΚ-ΘΡΑΚ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)

Case 22  ' ΚΕΝΤΡΙΚΗΣ ΕΛΛΑΔΑΣ
   If Val(aAcc) = 15700099 And Val(aBranch) = 7 Then
      aflag = True
   Else
      If Val(Mid(aAcc, 1, 3)) = 1 Then
         aflag = DoChkBankAccount(aBank + aBranch + "0" + Mid(aAcc, 1, 12), 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
      Else
         aflag = False
      End If
   End If

Case 24  ' XIOSBANK
   If Val(aAcc) = 7457000159# And Val(aBranch) = 1 Then
      aflag = True
   Else
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 25  ' ΔΩΡΙΚΗ
   If Val(aAcc) = 41110000 Then
      aflag = True
   Else
      atotal = (Val(aBranch) + 4073800000#) - (Int((Val(aBranch) + 4073800000#) / 97)) * 97
      atotal = atotal * 10000000000000# + Int(Val(aAcc) / 100) * 100
      atotal = 97 - (atotal - 97 * Int(atotal / 97))
      If Val(Mid(aAcc, 12, 2)) = atotal Then
         aflag = True
      Else
         aflag = False
      End If
   End If

Case 26, 27  ' INTERBANK, ΕΥΡΩΕΠΕΝΔΥΤΙΚΗ
   If Val(aBank) = 26 And Val(aBranch) = 1 And Val(aAcc) = 41110000 Then
      aflag = True
   Else
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 28  ' ΕΓΝΑΤΙΑ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 3, 5, 6, 1, 4, 0, 0, 0, 8, 9, 2, 7, 4, 9, 5, 2, 3), 0, 0, 0, 9999999999999#, False)

Case 31  ' ΕΥΡΩ-ΛΑΪΚΗ
   If Val(aAcc) = 918696001 Then
      aflag = True
   Else
      aflag = DoChkBankAccount(aBank + aBranch + "000" + Mid(aAcc, 1, 10), 1, 100, _
                Array(0, 0, 0, 1, 2, 1, 0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2), 0, 0, 0, 9999999999999#, True)
   End If

Case 43  ' ATE
   atotal = 0
   atotal = Val(Mid(aBranch, 1, 1)) * 7 + Val(Mid(aBranch, 2, 1)) * 11 + Val(Mid(aBranch, 3, 1)) * 13 + _
            Val(Mid(aAcc, 4, 1)) * 17 + Val(Mid(aAcc, 5, 1)) * 19 + Val(Mid(aAcc, 6, 1)) * 23 + _
            Val(Mid(aAcc, 7, 1)) * 29 + Val(Mid(aAcc, 8, 1)) * 31 + Val(Mid(aAcc, 9, 1)) * 37 + _
            Val(Mid(aAcc, 10, 1)) * 41 + Val(Mid(aAcc, 11, 1)) * 43
   If Val(Mid(aAcc, 12, 2)) = atotal Mod 97 Then
      aflag = True
   Else
      aflag = False
   End If

Case 47  ' ASPIS
   aflag = DoChkBankAccount(aBank + aBranch + "0" + Mid(aAcc, 1, 12), 11, 0, _
                Array(0, 0, 0, 0, 0, 0, 0, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)

Case 60  ' ABN AMRO
   If Val(aAcc) = 900440007 Then
      aflag = True
   Else
      atotal = 0
      For i = 5 To 13
          atotal = atotal + Val(Mid(aAcc, i, 1)) * (14 - i)
      Next i
      If atotal Mod 11 = 0 Then
         aflag = True
      Else
         aflag = False
      End If
   End If

Case 62  ' ANZ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)

Case 63  ' NWB
   aflag = DoChkBankAccount("1" + Mid(aBank, 1, 2) + aBranch + aAcc, 11, 0, _
                Array(115, 0, 0, 16, 16, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4), 0, 0, 0, 9999999999999#, False)

Case 65, 71, 97  ' BARCLAYS, MIDLAND
   If Val(aBank) = 65 And Val(aAcc) = 9237919000010# Then
      aflag = True
   Else
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 67  ' SOCIETE
   Dim atotal2 As Variant
   
   atotal = Val(Mid(aAcc, 2, 1)) + Val(Mid(aAcc, 4, 1)) + Val(Mid(aAcc, 6, 1)) + _
            Val(Mid(aAcc, 8, 1)) + Val(Mid(aAcc, 10, 1)) + Val(Mid(aAcc, 12, 1))
   Do While atotal > 9
      atotal = Int(atotal / 10) + (atotal Mod 10)
   Loop
   atotal2 = 2 * (Val(Mid(aAcc, 3, 1)) + Val(Mid(aAcc, 5, 1)) + Val(Mid(aAcc, 7, 1)) + _
            Val(Mid(aAcc, 9, 1)) + Val(Mid(aAcc, 11, 1)))
   Do While atotal2 > 9
      atotal2 = Int(atotal2 / 10) + (atotal2 Mod 10)
   Loop
   If Val(Mid(aAcc, 13, 1)) = ((Int((atotal + atotal2) / 10) + 1) * 10 - (atotal + atotal2)) Mod 10 Then
      aflag = True
   Else
      aflag = False
   End If
   
Case 68  ' CCF
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
                Array(0, 0, 0, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)

Case 70  ' BNP
   atotal = (Int(Val(aBranch + Mid(aAcc, 2, 7) + "00") / 97) * 97)
   atotal = Val(aBranch + Mid(aAcc, 2, 7) + "00") - atotal
   If Val(Mid(aAcc, 9, 2)) = 97 - atotal Then
      aflag = True
   Else
      aflag = False
   End If
   
Case 72  ' BAYER
   If Val(Mid(aAcc, 8, 4)) <> 2811 Or Val(Mid(aAcc, 12, 2)) < 1 Or Val(Mid(aAcc, 12, 2)) > 6 Then
      aflag = False
   Else
      aflag = DoChkBankAccount(aBank + aBranch + "000000" + Mid(aAcc, 1, 7), 11, 11, _
             Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   End If
   
Case 73  ' ΚΥΠΡΟΥ
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
             Array(7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   
Case 80  ' AMERICAN EXPRESS
   atotal = 0
   atotal = Val(Mid(aAcc, 5, 1)) * 2 + Val(Mid(aAcc, 6, 1)) * 3 + Val(Mid(aAcc, 7, 1)) * 4 + _
            Val(Mid(aAcc, 8, 1)) * 5 + Val(Mid(aAcc, 9, 1)) * 6 + Val(Mid(aAcc, 10, 1)) * 7 + _
            Val(Mid(aAcc, 11, 1)) * 8 + Val(Mid(aAcc, 12, 1)) * 9
   
   If Val(Mid(aAcc, 13, 1)) = (atotal Mod 11) Then
      aflag = True
   ElseIf ((atotal Mod 11) = 10) Then
      aflag = False
   Else
      If 10 + Val(Mid(aAcc, 13, 1)) - Val(Mid(aAcc, 12, 1)) = 11 Then
        aflag = True
      Else
        aflag = False
      End If
   End If

Case 81  ' BOF
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 10, 10, _
             Array(0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2), 0, 0, 0, 9999999999999#, False)

Case 84  ' CITYBANK
   If Val(Mid(aAcc, 1, 7)) = 93601 Then
      aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
              Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   Else
      aflag = DoChkBankAccount(aBank + aBranch + "0" + Mid(aAcc, 1, 12), 11, 11, _
              Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 3, 2, 7, 6, 5, 4, 3, 2), 7, 8, 0, 9999999999999#, False)
   End If
   
Case 87  ' PAGKRHTIA
   aflag = DoChkBankAccount(aBank + aBranch + aAcc, 11, 11, _
              Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   
Case Else
End Select

ChkBankAcount_ = aflag
End Function

Function DoChkBankCheque(inChk As String, inmod As Integer, _
                        Optional inminus As Integer = 999, Optional infactors As Variant, Optional inmod10 As Integer, Optional inmod11 As Integer) As Boolean
Dim inCd, corCd As Integer
Dim ntotal As Single

ntotal = 0
inCd = Val(Mid(inChk, 10, 1))

If inmod = 7 Then
   corCd = Val(Mid(inChk, 1, 9)) Mod 7
ElseIf inmod = 11 Then
   If IsMissing(infactors) = True Then
      corCd = Val(Mid(inChk, 1, 9)) Mod 11
   Else
      For i = 1 To 9
         ntotal = ntotal + Val(Mid(inChk, i, 1)) * infactors(i - 1)
      Next i
      corCd = ntotal Mod 11
   End If
   If inminus <> 999 Then
      corCd = inminus - corCd
   End If
   If corCd = 10 Then
      corCd = inmod10
   ElseIf corCd = 11 Then
      corCd = nmod11
   End If
End If

If inCd = corCd Then
   DoChkBankCheque = True
Else
   DoChkBankCheque = False
End If

End Function

Public Function ChkBankCheque_(inBank As String, inbranch As String, _
    inAcc As String, inCheque As String) As Boolean

Dim aflag As Boolean
Dim aBank As String, aAcc As String, aBranch As String, aCheque As String
Dim atotal As Variant

aBank = StrPad_(inBank, 3, "0")
aBranch = StrPad_(inbranch, 3, "0")
aAcc = StrPad_(inAcc, 13, "0")
aCheque = StrPad_(inCheque, 10, "0")

Select Case Val(aBank)
   Case 12
      aflag = DoChkBankCheque(aCheque, 7)
   Case 13
      If Val(aAcc) = 57000007 Then
         aflag = DoChkBankCheque(aCheque, 7)
      Else
         aflag = DoChkBankCheque(aCheque, 11, 11, , 7, 0)
      End If
   Case 14
      aflag = DoChkBankCheque(aCheque, 11, 11, , 0, 0)
   Case 15
      If Val(Mid(aCheque, 1, 2)) = 0 Then
         aflag = DoChkBankCheque(aCheque, 11, 11, , 0, 0)
      ElseIf Val(Mid(aCheque, 1, 1)) = 0 Then
         aflag = DoChkBankCheque("00" + Mid(aCheque, 3, 8), 11, 11, , 0, 0)
      Else
         aflag = False
      End If
   Case 16
      aflag = DoChkBankCheque(aCheque, 7)
   Case 17
      aflag = DoChkBankCheque(aCheque, 7)
   Case 18
      If Val(aCheque) <= 89999990 And Val(aCheque) >= 30000000 Then
         aflag = DoChkBankCheque(aCheque, 11, , , 0, 0)
      Else
         aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 2, 7, 6, 5, 4, 3, 2))
      End If
   Case 19
      If Val(Mid(aCheque, 1, 2)) <> 0 Then
         aflag = False
      Else
         aflag = DoChkBankCheque(aCheque, 7)
      End If
   Case 20
      If Val(Mid(aCheque, 1, 2)) <> 0 Then
         aflag = False
      Else
         If Val(Mid(aAcc, 11, 2)) <> 99 Then
            aflag = True
         Else
            For i = 3 To 9
               atotal = atotal + Int((Val(Mid(aCheque, i, 1)) * ((i Mod 2) + 1)) / 10) + _
                                     (Val(Mid(aCheque, i, 1)) * ((i Mod 2) + 1) Mod 10)
            Next i
         End If
      End If
   Case 21
      If Val(Mid(aCheque, 1, 2)) <> 0 Then
         aflag = False
      Else
         aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 64, 32, 16, 8, 4, 2, 1))
      End If
   Case 22
      If Val(aBranch) = 7 And Val(aaccount) = 15700099 And Val(Mid(aaccount, 1, 4)) = 0 Then
         aflag = True
      Else
         If Val(Mid(aCheque, 1, 2)) <> 0 Then
            aflag = False
         Else
            aflag = DoChkBankCheque(aCheque, 11, 11, , 0, 0)
         End If
      End If
   Case 24
      aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 64, 32, 16, 8, 4, 2, 1))
   Case 25
      aflag = DoChkBankCheque(aCheque, 11, 11, , 0, 0)
   Case 26
      If Val(aaccount) <> 41110000 Then
         aflag = True
      Else
         If Val(Mid(aCheque, 1, 2)) <> 0 Then
            aflag = False
         Else
            aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 64, 32, 16, 8, 4, 2, 1))
         End If
      End If
   Case 27
      If Val(aaccount) = 41110006 Or Val(aaccount) = 41110000 Then
         aflag = True
      Else
         If Val(Mid(aCheque, 1, 3)) <> 0 Then
            aflag = False
         Else
            aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 0, 32, 16, 8, 4, 2, 1))
         End If
      End If
   Case 43
      If Val(aCheque) < 1000000 Or Val(aCheque) > 10000000 Then
         aflag = False
      Else
         aflag = DoChkBankCheque(aCheque, 11, , , 0, 0)
      End If
   Case 60
      If Val(Mid(aChequem, 1, 4)) <> 0 Then
         aflag = False
      Else
         aflag = True
      End If
   Case 65
      If Val(aaccount) = 9237919000010# Then
         aflag = True
      Else
         aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 64, 32, 16, 8, 4, 2, 1))
      End If
   Case 71
      If Val(Mid(aCheque, 1, 2)) <> 0 Then
         aflag = False
      Else
         aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 64, 32, 16, 8, 4, 2, 1))
      End If
   Case 73
      aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 0, 2, 7, 6, 5, 4, 3, 2))
   Case 87
      aflag = DoChkBankCheque(aCheque, 11, , , 0, 0)
   Case 97
      aflag = DoChkBankCheque(aCheque, 11, 11, Array(0, 128, 64, 32, 16, 8, 4, 2, 1))
   
   Case Else
      aflag = True
      
End Select
ChkBankCheque_ = aflag
End Function



