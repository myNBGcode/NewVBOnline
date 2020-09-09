Attribute VB_Name = "ChkDigits"
Option Explicit

Public Function CalcCd1_(Acc As String, Digits As Integer) As Integer
'Πρώτο Check Digit λογαριασμού
    Dim total, i, Rm As Integer
  
    total = 0
    For i = Digits To 1 Step -1
        total = total + Val(Mid(Acc, i, 1)) * (Digits + 2 - i)
    Next
    CalcCd1_ = 11 - (total Mod 11)
    CalcCd1_ = (1 - (CalcCd1_ \ 10)) * CalcCd1_
End Function

Public Function CalcCd1_4330_(Acc As String, Digits As Integer) As Integer
'Πρώτο Check Digit για συναλλαγή 4330 (mod 11 για εφορία)
    Dim total As Integer, i As Integer, Rm As Integer, k As Integer
  
    total = 0: k = 1
    For i = Len(Acc) - 1 To 1 Step -1
        k = k + 1
        If k > 7 Then k = 2
        total = total + Val(Mid(Acc, i, 1)) * k
    Next
    CalcCd1_4330_ = 11 - (total Mod 11)
    CalcCd1_4330_ = CalcCd1_4330_ Mod 10
End Function

Public Function CalcCdR(Acc As String, Digits As Integer) As Integer
'Check Digit αριθμού εγγραφής
    Dim total, i, Rm As Integer
  
    total = 0
    For i = Digits To 1 Step -1
        total = total + Val(Mid(Acc, i, 1)) * (Digits + 2 - i)
    Next
    CalcCdR = 11 - (total Mod 11)
    CalcCdR = (CalcCdR \ 10)
End Function

Public Function CalcCd2_(aAcc10 As String) As Integer
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
    
    CalcCd2_ = total Mod 10
    If CalcCd2_ <> 0 Then CalcCd2_ = 10 - CalcCd2_
    
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
    
  CalcSAccCd = CalcCd1_(SAcc, Len(SAcc))
End Function

Public Function ChkDocument(inNum As String) As Boolean
'Συναλλαγή 1020 Πεδίο Έγγραφο / Επιταγή
    Dim copyNum As String, part1 As String
    copyNum = StrPad_(inNum, 11, "0", "L")
    ChkDocument = _
        ((CalcCd1_(Left(copyNum, 7), 7) = CInt(Mid(copyNum, 8, 1))) _
        And (CLng(Left(copyNum, 7)) > 0))
End Function

Public Function ChkCard2(Acc As String, Digits As Integer) As Integer
'Check Digit Κάρτας ότα ο αριθμός έχει μέχρι 8 ψηφία
'τότε είναι αριθμός δανείου
    Dim total, i, Rm As Integer
  
    total = 0
    For i = Digits To 1 Step -1
        total = total + Val(Mid(Acc, i, 1)) * (Digits + 2 - i)
    Next
    ChkCard2 = (11 - (total Mod 11)) Mod 10
End Function

Public Function ChkCard(inNum As String) As Boolean
'Συναλλαγή 4010 Πεδίο Αριθμός Κάρτας
    Dim copyNum As String, Num1 As Integer, Sum1 As Integer, i As Integer
    copyNum = StrPad_(inNum, 16, "0", "L")
    
    If CDbl(copyNum) = 0 Then
        ChkCard = True
    ElseIf CDbl(Left(copyNum, 3)) = 0 And CDbl(Left(copyNum, 4)) > 0 Then '13 μη μηδενικά ψηφία
        ChkCard = True
    ElseIf CDbl(Left(copyNum, 8)) = 0 Then '8 τουλάχιστο μηδενικό
        ChkCard = (ChkCard2(Mid(copyNum, 9, 7), 7) = CInt(Right(copyNum, 1)))
    ElseIf CDbl(Left(copyNum, 1)) > 0 Then '16 μη μηδενικά ψηφία
        Sum1 = 0
        For i = 1 To 15 Step 2
            Num1 = CInt(Mid(copyNum, i, 1) * 2)
            Sum1 = Sum1 + Num1 \ 10 + Num1 Mod 10
        Next
        For i = 2 To 14 Step 2
            Sum1 = Sum1 + CInt(Mid(copyNum, i, 1))
        Next i
        ChkCard = (((10 - (Sum1 Mod 10)) Mod 10) = Right(copyNum, 1))
    Else
        ChkCard = False
    End If
End Function

Public Function ChkTaxID_(inNum As String) As Boolean
'Συναλλαγή 4150 Πεδίο ΑΦΜ - Συναλλαγή 4350
    Dim copyNum As String, Sum1 As Long, i As Long
    On Error GoTo ErrorPos
    copyNum = StrPad_(inNum, 9, "0", "L")
    For i = 1 To 8 Step 1
        Sum1 = Sum1 + CInt(Mid(copyNum, i, 1)) * (2 ^ (9 - i))
    Next
    ChkTaxID_ = (((Sum1 Mod 11) Mod 10) = CInt(Right(copyNum, 1))): Exit Function
ErrorPos:
    ChkTaxID_ = False: Exit Function
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
        ((CalcCd1_(Left(copyNum, 6), 6) = CInt(Mid(copyNum, 7, 1))) _
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
                        
                        

Dim inCd As Integer, corCd As Integer, i As Integer
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

Public Function ChkBankAccount_(inBank As String, inbranch As String, inAcc As String, Optional inChequeType As Integer) As Boolean
Dim aFlag As Boolean, i As Integer
Dim abank, abranch, aacc As String
Dim atotal As Variant
Dim CD, acc1, mod11 As Double

abank = StrPad_(inBank, 3, "0")
abranch = StrPad_(inbranch, 3, "0")
aacc = StrPad_(inAcc, 13, "0")

Select Case Val(abank)
Case 12
    aFlag = False
    If Len(inAcc) = 13 Then
        If Left(inAcc, 2) = "92" Then aFlag = True
    End If
    
    If Not aFlag Then
         aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                 Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
         If Not aFlag Then
        
        
             If Mid(aacc, 1, 1) = 9 Then
                aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                          Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
             Else
                aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                          Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
                If Val(aacc) = 57000007 And Val(abank) = 13 Then aFlag = True
             End If
         End If
    End If
Case 16
    aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
            Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
    If Not aFlag Then
        If Mid(aacc, 1, 1) = 9 Then
           aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                     Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
        Else
           aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                     Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
           If Val(aacc) = 57000007 And Val(abank) = 13 Then aFlag = True
        End If
    End If
    
    If Not aFlag Then aFlag = (aacc = "0000005700000" And (abranch = "065" Or abranch = "069"))
Case 13, 17  ' ΕΜΠΟΡΙΚΗ, ΙΟΝΙΚΗ, ΑΤΤΙΚΗΣ, ΠΕΙΡΑΙΩΣ
    aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
            Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
    If Not aFlag Then
   
   
        If Mid(aacc, 1, 1) = 9 Then
           aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                     Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
        Else
           aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                     Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
           If Val(aacc) = 57000007 And Val(abank) = 13 Then aFlag = True
        End If
    End If

Case 14  ' ΠΙΣΤΕΩΣ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 16384, 8192, 4096, 0, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)

Case 15  ' ΓΕΝΙΚΗ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   If Not aFlag Then
        'παλιά μορφή λογαριασμού
        aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                    Array(0, 0, 0, 6, 5, 4, 0, 0, 0, 0, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   End If
   If Val(aacc) = 574 Then aFlag = True

Case 18  ' ΑΘΗΝΩΝ
   aFlag = DoChkBankAccount(abank + abranch + "0" + Mid(aacc, 1, 12), 11, 11, _
                Array(0, 0, 0, 31, 29, 26, 0, 0, 0, 23, 22, 19, 17, 13, 7, 6, 3, 2), 0, 0, 0, 9999999999999#, False)
   If aFlag = True Then
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 6, 5, 4, 0, 0, 0, 0, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   End If

Case 19  ' ΚΡΗΤΗΣ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)

Case 20  ' ΕΡΓΑΣΙΑΣ
   If Mid(aacc, 11, 2) = "99" Then
      If Mid(aacc, 11, 3) = "991" And inbranch = 137 Then
         aFlag = True
      Else
        aFlag = DoChkBankAccount(abank + abranch + aacc, 1, 100, _
                   Array(1, 2, 1, 2, 1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 2), 0, 0, 0, 9999999999999#, True)
      End If
   Else
      If Mid(aacc, 7, 2) = "00" And (Mid(aacc, 9, 2) = "00" Or Mid(aacc, 9, 2) = "04") Then
         aFlag = DoChkBankAccount(abank + abranch + "0" + Mid(aacc, 1, 12), 1, 100, _
                Array(0, 0, 0, 1, 2, 1, 0, 0, 2, 1, 2, 1, 2, 0, 0, 0, 0, 0), 0, 0, 0, 9999999999999#, True)
         If aFlag = True Then
            aFlag = DoChkBankAccount(abank + abranch + aacc, 1, 100, _
                 Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2, 0), 0, 0, 0, 9999999999999#, True)
         End If
      Else
        aFlag = False
      End If
   End If

Case 21  ' ΜΑΚ-ΘΡΑΚ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)

Case 22  ' ΚΕΝΤΡΙΚΗΣ ΕΛΛΑΔΑΣ
   If Val(aacc) = 15700099 And Val(abranch) = 7 Then
      aFlag = True
   Else
      If Val(Mid(aacc, 1, 3)) = 1 Then
         aFlag = DoChkBankAccount(abank + abranch + "0" + Mid(aacc, 1, 12), 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
      Else
         aFlag = False
      End If
   End If

Case 24  ' XIOSBANK
   If Val(aacc) = 7457000159# And Val(abranch) = 1 Then
      aFlag = True
   Else
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 25  ' ΔΩΡΙΚΗ
   If Val(aacc) = 41110000 Or Val(aacc) = 41120005 Then
      aFlag = True
   Else
      atotal = (Val(abranch) + 4073800000#) - (Int((Val(abranch) + 4073800000#) / 97)) * 97
      atotal = atotal * 10000000000000# + Int(Val(aacc) / 100) * 100
      atotal = 97 - (atotal - 97 * Int(atotal / 97))
      If Val(Mid(aacc, 12, 2)) = atotal Then
         aFlag = True
      Else
         aFlag = False
      End If
   End If

Case 26, 27  ' INTERBANK, ΕΥΡΩΕΠΕΝΔΥΤΙΚΗ
   If Val(abank) = 26 And Val(abranch) = 1 And Val(aacc) = 41110000 Then
      aFlag = True
   Else
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 28  ' ΕΓΝΑΤΙΑ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 3, 5, 6, 1, 4, 0, 0, 0, 8, 9, 2, 7, 4, 9, 5, 2, 3), 0, 0, 0, 9999999999999#, False)

Case 31  ' ΕΥΡΩ-ΛΑΪΚΗ
   If Val(aacc) = 918696001 Then
      aFlag = True
   Else
      aFlag = DoChkBankAccount(abank + abranch + "000" + Mid(aacc, 1, 10), 1, 100, _
                Array(0, 0, 0, 1, 2, 1, 0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2), 0, 0, 0, 9999999999999#, True)
   End If

Case 32 'elliniki
    aFlag = DoChkBankAccount(Right(String(19, "0") & Trim(aacc), 19), 10, 10, Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 0, 2, 1, 2, 1, 2), 0, 0, 0, 9999999999999#, True)

Case 36  ' ΣΥΝ/ΚΗ ΤΡΑΠΕΖΑ ΔΥΤ.ΜΑΚΕΔΟΝΙΑΣ

    Dim newacc36 As String
    newacc36 = Right(aacc, 11)
    newacc36 = Left(newacc36, 1) + abranch + Mid(newacc36, 2)

    CD = CDbl(Mid(newacc36, 14, 1))
    acc1 = CDbl(Mid(newacc36, 1, 13))
    CD = Int(CD + 0.5)
    acc1 = Int(acc1 + 0.5)
    mod11 = acc1 - (11# * Fix(acc1 / 11))
    
    If mod11 = 10 Then mod11 = 0
    
    If CD = mod11 Then
        aFlag = True
    Else
        aFlag = False
    End If

Case 38 'nova
   Dim aaccString As String
   aaccString = Right("0000000000" & aacc, 10)
   Dim vArray() As Variant
   Dim aTotal38 As Long
   atotal = 0
   vArray = Array(2, 1, 2, 1, 2, 1, 2, 1, 2, 1)
   For i = 0 To 9
       aTotal38 = aTotal38 + ((Mid(aaccString, i + 1, 1) * vArray(i)) Mod 10)
   Next i
   If aTotal38 Mod 10 = 0 Then
    aFlag = True
   Else
    aFlag = False
   End If

Case 43  ' ATE
   atotal = 0
   atotal = Val(Mid(abranch, 1, 1)) * 7 + Val(Mid(abranch, 2, 1)) * 11 + Val(Mid(abranch, 3, 1)) * 13 + _
            Val(Mid(aacc, 4, 1)) * 17 + Val(Mid(aacc, 5, 1)) * 19 + Val(Mid(aacc, 6, 1)) * 23 + _
            Val(Mid(aacc, 7, 1)) * 29 + Val(Mid(aacc, 8, 1)) * 31 + Val(Mid(aacc, 9, 1)) * 37 + _
            Val(Mid(aacc, 10, 1)) * 41 + Val(Mid(aacc, 11, 1)) * 43
   If Val(Mid(aacc, 12, 2)) = atotal Mod 97 Then
      aFlag = True
   Else
      aFlag = False
   End If

Case 47  ' ASPIS
   If Not IsMissing(inChequeType) And inChequeType = 1 Then 'Ιδιωτική
        If Left(aacc, 1) = "0" Then
           aFlag = DoChkBankAccount(abank & abranch & aacc, 11, 0, _
                Array(0, 0, 0, 0, 0, 0, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)
        Else
           aFlag = DoChkBankAccount("000001" & aacc, 11, 0, _
                Array(0, 0, 0, 0, 0, 16384, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)
        End If
   ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
        If Right(aacc, 1) <> "0" Or Left(aacc, 2) <> "00" Then
            aFlag = False
        Else
            aFlag = DoChkBankAccount(abank + abranch + "0" + Mid(aacc, 1, 12), 11, 0, _
                          Array(0, 0, 0, 0, 0, 0, 0, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)
        End If
   Else
        If Right(aacc, 1) <> "0" Then
             aFlag = False
        Else
             aFlag = DoChkBankAccount(abank + abranch + "0" + Mid(aacc, 1, 12), 11, 0, _
                          Array(0, 0, 0, 0, 0, 0, 0, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)
        End If
   End If

Case 49  ' ΠΑΝΕΛΛΗΝΙΟΣ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 1), 7, 8, 0, 9999999999999#, False)

Case 54  ' PROBANK
    
    If Right("000" & Trim(abank), 3) = "054" And _
       Right("000" & Trim(abranch), 3) = "800" And _
       Right("0000000000" & Trim(aacc), 10) = "0157000000" Then
       
       'Τραπεζικές επιταγές
        aFlag = True
    Else
        aFlag = DoChkBankAccount(abank & abranch & aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)
    End If
Case 55  ' FIRST BUSINESS BANK

Case 56  ' Aegean Baltic Bank A.T.E
   If Val(aacc) = 0 Then
    aFlag = False
   Else
    
    aFlag = DoChkBankAccount(abank & abranch & aacc, 11, 11, _
                 Array(0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 7, 8, 0, 9999999999999#, False)
     If aFlag Then
         If Not IsMissing(inChequeType) And inChequeType = 1 Then
         ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then
             If (abranch = "100" And Val(aacc) = 900000157024#) Or (abranch = "102" And Val(aacc) = 900000157218#) Then
             Else
                 aFlag = False
             End If
         End If
     End If
     
   End If
Case 60  ' ABN AMRO
   If Val(aacc) = 900440007 Then
      aFlag = True
   Else
      atotal = 0
      For i = 5 To 13
          atotal = atotal + Val(Mid(aacc, i, 1)) * (14 - i)
      Next i
      If atotal Mod 11 = 0 Then
         aFlag = True
      Else
         aFlag = False
      End If
   End If

Case 62  ' ANZ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)

Case 63  ' NWB
    aFlag = True
'   aflag = DoChkBankAccount("1" + Mid(aBank, 1, 2) + aBranch + aAcc, 11, 0, _
'                Array(115, 0, 0, 16, 16, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4), 0, 0, 0, 9999999999999#, False)

Case 64  'THE ROYAL BANK OF SCOTLAND PLC

    If abranch = "095" Then
        Dim ntotal64 As Single
        Dim infactors64 As Variant
        infactors64 = Array(0, 0, 0, 0, 9, 8, 7, 6, 5, 4, 3, 2, 1)

        ntotal64 = 0
        For i = 1 To Len(aacc)
            ntotal64 = ntotal64 + Val(Mid(aacc, i, 1)) * infactors64(i - 1)
        Next i
        If ntotal64 Mod 11 = 0 Then
            aFlag = True
        Else
            aFlag = False
        End If
    Else
        aFlag = True
    End If
             
    If aFlag And Not IsMissing(inChequeType) And inChequeType = 9 Then   'τραπεζικη
        If Not (aacc = "0000900440007" And abranch = "095") Then
            aFlag = False
        End If
    End If
             
Case 65, 71, 41, 74  ' BARCLAYS, MIDLAND, ETBA, Συν.Λαμίας
   If Val(abank) = 65 And Val(aacc) = 9237919000010# Then
      aFlag = True
   Else
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If

Case 67  ' SOCIETE
   Dim atotal2 As Variant
   
   atotal = Val(Mid(aacc, 2, 1)) + Val(Mid(aacc, 4, 1)) + Val(Mid(aacc, 6, 1)) + _
            Val(Mid(aacc, 8, 1)) + Val(Mid(aacc, 10, 1)) + Val(Mid(aacc, 12, 1))
   Do While atotal > 9
      atotal = Int(atotal / 10) + (atotal Mod 10)
   Loop
   atotal2 = 2 * (Val(Mid(aacc, 3, 1)) + Val(Mid(aacc, 5, 1)) + Val(Mid(aacc, 7, 1)) + _
            Val(Mid(aacc, 9, 1)) + Val(Mid(aacc, 11, 1)))
   Do While atotal2 > 9
      atotal2 = Int(atotal2 / 10) + (atotal2 Mod 10)
   Loop
   If Val(Mid(aacc, 13, 1)) = ((Int((atotal + atotal2) / 10) + 1) * 10 - (atotal + atotal2)) Mod 10 Then
      aFlag = True
   Else
      aFlag = False
   End If
   
Case 68  ' CCF
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2), 0, 0, 0, 9999999999999#, False)

Case 69  ' ΣΥΝΕΤΑΙΡΙΣΤΙΚΗ ΤΡΑΠΕΖΑ ΧΑΝΙΩΝ
    Dim cd69, acc69, mod11_69 As Double
    cd69 = CDbl(Mid(aacc, 13, 1))
    acc69 = CDbl(Mid(aacc, 1, 12))
    cd69 = Int(cd69 + 0.5)
    acc69 = Int(acc69 + 0.5)
    mod11_69 = acc69 - (11# * Fix(acc69 / 11))
    
    If mod11_69 = 10 Then mod11_69 = 0
    
    If cd69 = mod11_69 Then
        aFlag = True
    Else
        aFlag = False
    End If

Case 70  ' BNP
   atotal = (Int(Val(abranch + Mid(aacc, 2, 7) + "00") / 97) * 97)
   atotal = Val(abranch + Mid(aacc, 2, 7) + "00") - atotal
   If Val(Mid(aacc, 9, 2)) = 97 - atotal Then
      aFlag = True
   Else
      aFlag = False
   End If
   
Case 72  ' BAYER
   If Val(Mid(aacc, 8, 4)) <> 2811 Or Val(Mid(aacc, 12, 2)) < 1 Or Val(Mid(aacc, 12, 2)) > 6 Then
      aFlag = False
   Else
      aFlag = DoChkBankAccount(abank + abranch + "000000" + Mid(aacc, 1, 7), 11, 11, _
             Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   End If
   
Case 73  ' ΚΥΠΡΟΥ
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
             Array(7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   
Case 75, 94  ' ΣΥΝ/ΚΗ ΤΡΑΠΕΖΑ ΙΩΑΝΝΙΝΩΝ & ΣΥΝ/ΚΗ ΤΡΑΠΕΖΑ ΠΙΕΡΙΑΣ
    Dim newacc As String
    newacc = Right(aacc, 11)
    newacc = Left(newacc, 1) + abranch + Mid(newacc, 2)

    CD = CDbl(Mid(newacc, 14, 1))
    acc1 = CDbl(Mid(newacc, 1, 13))
    CD = Int(CD + 0.5)
    acc1 = Int(acc1 + 0.5)
    mod11 = acc1 - (11# * Fix(acc1 / 11))
    
    If mod11 = 10 Then mod11 = 0
    
    If CD = mod11 Then
        aFlag = True
    Else
        aFlag = False
    End If

Case 77 ' Αχαϊκής Συνεταιριστικής
    
    Dim newacc77 As String
    newacc77 = Right(aacc, 11)
    newacc77 = Left(newacc77, 1) + abranch + Mid(newacc77, 2)

    CD = CDbl(Mid(newacc77, 14, 1))
    acc1 = CDbl(Mid(newacc77, 1, 13))
    CD = Int(CD + 0.5)
    acc1 = Int(acc1 + 0.5)
    mod11 = acc1 - (11# * Fix(acc1 / 11))
    
    If mod11 = 10 Then mod11 = 0
    
    If CD = mod11 Then
        aFlag = True
    Else
        aFlag = False
    End If
    
    

    
    
    
Case 78  ' ING BANK
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
             Array(7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
Case 79  ' Συνεταιριστικη Δωδεκανησου
'   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
'             Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 9, 8, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
             Array(0, 0, 0, 0, 0, 0, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
   
Case 80  ' AMERICAN EXPRESS
   atotal = 0
   atotal = Val(Mid(aacc, 5, 1)) * 2 + Val(Mid(aacc, 6, 1)) * 3 + Val(Mid(aacc, 7, 1)) * 4 + _
            Val(Mid(aacc, 8, 1)) * 5 + Val(Mid(aacc, 9, 1)) * 6 + Val(Mid(aacc, 10, 1)) * 7 + _
            Val(Mid(aacc, 11, 1)) * 8 + Val(Mid(aacc, 12, 1)) * 9
   
   If Val(Mid(aacc, 13, 1)) = (atotal Mod 11) Then
      aFlag = True
   ElseIf ((atotal Mod 11) = 10) Then
      aFlag = False
   Else
      If 10 + Val(Mid(aacc, 13, 1)) - Val(Mid(aacc, 12, 1)) = 11 Then
        aFlag = True
      Else
        aFlag = False
      End If
   End If

Case 81  ' BOF
   aFlag = DoChkBankAccount(abank + abranch + aacc, 10, 10, _
             Array(0, 0, 0, 0, 0, 0, 0, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2), 0, 0, 0, 9999999999999#, False)

Case 84  ' CITYBANK
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
              Array(131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)

Case 87  ' Παγκρήτια
   'aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
              'Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
           Array(0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)
   If Not aFlag Then
        'παλιά μορφή λογαριασμού
        aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
              Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 256, 128, 64, 32, 16, 8, 4, 2, 1), 0, 0, 0, 9999999999999#, False)
   End If


Case 88 ' συνεταιριστικη εβρου
   If Not IsMissing(inChequeType) And inChequeType = 1 Then 'Ιδιωτική
        aFlag = True
   ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
        If aacc = "0000000700001" Then
            aFlag = True
        Else
            aFlag = False
        End If
   Else
        aFlag = True
   End If

Case 89
    Dim check As String
    check = Mid(aacc, 3, 1) & abranch & Mid(aacc, 4, 9)

    CD = CDbl(Right(aacc, 1))
    acc1 = CDbl(check)
    CD = Int(CD + 0.5)
    acc1 = Int(acc1 + 0.5)
    mod11 = acc1 - (11# * Fix(acc1 / 11))
    
    If mod11 = 10 Then mod11 = 0
    aFlag = (CD = mod11)

Case 91 'ΣΥΝΕΤΑΙΡΙΣΤΙΚΗ ΤΡΑΠΕΖΑ ΘΕΣΣΑΛΙΑΣ ΣΥΝ.ΠΕ.
    aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(2 ^ 17, 2 ^ 16, 2 ^ 15, 2 ^ 14, 2 ^ 13, 2 ^ 12, 2 ^ 11, 2 ^ 10, 2 ^ 9, 2 ^ 8, 2 ^ 7, 2 ^ 6, 2 ^ 5, 2 ^ 4, 2 ^ 3, 2 ^ 2, 2 ^ 1, 2 ^ 0), 0, 0, 0, 9999999999999#, False)

Case 92 'ΣΥΝΕΤΑΙΡΙΣΤΙΚΗ ΤΡΑΠΕΖΑ ΠΕΛΟΠΟΝΝΗΣΟΥ
    atotal = Mid(aacc, 3, 1) & abranch & Mid(aacc, 4, 9)
    atotal = Right("0000000000000" & atotal, 13)
        
    atotal2 = (Left(atotal, 6) Mod 11) * (10000000 Mod 11) + (Right(atotal, 7) Mod 11)
    atotal2 = atotal2 Mod 11

    If atotal2 = 10 Then atotal2 = 0
    aFlag = (atotal2 = CInt(Right(aacc, 1)))
    
    
Case 95 'Συνεταιριστικη Τράπεζα Δραμας
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2), 7, 8, 0, 9999999999999#, False)
Case 96 'Ταχυδρομικο Ταμιευτηριο
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(0, 0, 0, 0, 0, 0, 0, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2), 0, 1, 0, 9999999999999#, False)
Case 97  ' ΤαΠαρΔαν
      aFlag = DoChkBankAccount(abank + abranch + aacc, 11, 11, _
                Array(3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2), 0, 0, 0, 9999999999999#, False)

Case 98 'ΣΥΝΕΤΑΙΡΙΣΤΙΚΗ ΤΡΑΠΕΖΑ ΛΕΣΒΟΥ ΛΗΜΝΟΥ
   Dim aTotal98 As Long
   Dim aaccString98 As String
   Dim v98Array As Variant
   Dim cor98CD As Integer
   aaccString98 = Right("000000000000" & aacc, 12)
   aTotal98 = 0: v98Array = Array(6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
   For i = 0 To 10
       aTotal98 = aTotal98 + ((Mid(aaccString98, i + 1, 1) * v98Array(i)))
   Next i
   cor98CD = aTotal98 Mod 11
   cor98CD = 11 - cor98CD
   If cor98CD > 9 Then cor98CD = cor98CD - 3
   If cor98CD = Mid(aaccString98, 12, 1) Then
      aFlag = True
   Else
      aFlag = False
   End If
   
Case 99
    Dim acc99 As String
    
    acc99 = Right(aacc, 11)
    acc99 = Left(acc99, 1) + abranch + Mid(acc99, 2)
    If Len(acc99) <> 14 Then
        aFlag = False
    Else
        CD = CDbl(Right(acc99, 1))
        acc1 = CDbl(Mid(acc99, 1, 13))
        CD = Int(CD + 0.5)
        acc1 = Int(acc1 + 0.5)
        mod11 = acc1 - (11# * Fix(acc1 / 11))

        If mod11 = 10 Then mod11 = 0

        If CD = mod11 Then
            aFlag = True
        Else
            aFlag = False
        End If
    End If

Case 107 'GREEK BRANCH OF CLOSED JOINT STOCK COMPANY COMMERCIAL BANK
    Dim acc107 As String
    Dim infactors107 As Variant
    Dim ntotal107 As Double
    acc107 = Mid(aacc, 2, 11)
    
    infactors107 = Array(6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    For i = 1 To 11
        ntotal107 = ntotal107 + Val(Mid(acc107, i, 1)) * infactors107(i - 1)
    Next i
    mod11 = ntotal107 Mod 11

    CD = 11 - mod11
    If CD > 9 Then
        CD = CD - 3
    End If
    If CD = CDbl(Right(aacc, 1)) Then
        aFlag = True
    Else
        aFlag = False
    End If

    If Not IsMissing(inChequeType) And inChequeType = 9 And aFlag Then 'Τραπεζική
        If aacc = "0000000000067" Then
            aFlag = True
        Else
            aFlag = False
        End If
    End If

Case 109 'T.C. ZIRAAT BANKASI A.S.

    Dim acc109 As String
    Dim infactors As Variant
    Dim ntotal As Double
    acc109 = Mid(aacc, 2, 11)

    infactors = Array(6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    For i = 1 To 11
        ntotal = ntotal + Val(Mid(acc109, i, 1)) * infactors(i - 1)
    Next i
    mod11 = ntotal Mod 11

    CD = 11 - mod11
    If CD > 9 Then
        CD = CD - 3
    End If
    If CD = CDbl(Right(aacc, 1)) Then
        aFlag = True
    Else
        aFlag = False
    End If

    If Not IsMissing(inChequeType) And inChequeType = 9 And aFlag Then 'Τραπεζική
        If aacc = "0000321430015" Then
            aFlag = True
        Else
            aFlag = False
        End If
    End If

Case Else
   aFlag = True
End Select

ChkBankAccount_ = aFlag
End Function

Function DoChkBankCheque(inChk As String, inmod As Integer, _
                        Optional inminus As Integer = 999, Optional infactors As Variant, _
                        Optional inmod10 As Integer, Optional inmod11 As Integer) As Boolean
Dim inCd As Integer, corCd As Integer, i As Integer
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
      corCd = inmod11
   End If
End If

If inCd = corCd Then
    DoChkBankCheque = True
Else
    DoChkBankCheque = False
'    If cVersion >= 20010101 Then DoChkBankCheque = ChkGenBankCheque_(inChk)
End If

End Function

Public Function ChkBankCheque_(inBank As String, inbranch As String, _
    inAcc As String, inCheque As String, Optional inChequeType As Integer) As Boolean

Dim aFlag As Boolean, i As Integer
Dim abank As String, aacc As String, abranch As String, acheque As String
Dim atotal As Variant
Dim chequeno As Long

abank = StrPad_(inBank, 3, "0")
abranch = StrPad_(inbranch, 3, "0")
aacc = StrPad_(inAcc, 13, "0")
acheque = StrPad_(inCheque, 10, "0")
chequeno = CLng(Left(acheque, Len(acheque) - 1))

Select Case Val(abank)
   Case 12, 14, 15, 16, 17, 20
      aFlag = DoChkBankCheque(acheque, 7)
   Case 25
      aFlag = DoChkBankCheque(acheque, 11, 11, , 0, 0)
   Case 26
      aFlag = ChkGenBankCheque_(acheque)
   Case 34 'Επενδυτική
    aFlag = ChkGenBankCheque_(acheque)
   Case 28, 31, 32, 37
      aFlag = ChkGenBankCheque_(acheque)
   Case 38 'nova
      aFlag = DoChkBankCheque(acheque, 11, 11, Array(0, 128, 64, 32, 16, 8, 4, 2, 1))
   Case 41, 43
      aFlag = ChkGenBankCheque_(acheque)
   Case 47
      If Not IsMissing(inChequeType) And inChequeType = 1 Then 'Ιδιωτική
        aFlag = ChkGenBankCheque_(acheque)
      ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
        aFlag = (Len(CStr(Val(acheque))) = 8)
      Else
        aFlag = True
      End If
   Case 49
      aFlag = ChkGenBankCheque_(acheque)
   Case 54, 55, 60
      aFlag = ChkGenBankCheque_(acheque)
   Case 67, 69, 70, 71, 72, 73, 74, 75, 78, 79, 88, 99, 77, 109, 92, 36, 56, 94
      aFlag = ChkGenBankCheque_(acheque)
   Case 64
      aFlag = ChkGenBankCheque_(acheque)
      If aFlag Then
        If Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
          If Left(CStr(chequeno), 1) <> "1" Then
            aFlag = False
            ChkBankCheque_ = aFlag
            Exit Function
          End If
        End If
      End If
   Case 107
      aFlag = ChkGenBankCheque_(acheque)
      If aFlag Then
        If Not IsMissing(inChequeType) And inChequeType = 1 Then 'Ιδιωτική
          If chequeno < 1 Or chequeno > 9000000 Then
            aFlag = False
          End If
        ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
          If chequeno < 9000001 Or chequeno > 9999999 Then
            aFlag = False
          End If
        End If
      End If
   Case 80, 81
      aFlag = True
   Case 84
      aFlag = DoChkBankCheque(acheque, 11, 11, Array(256, 128, 64, 32, 16, 8, 4, 2, 1, 0), 0, 0)
   Case 87, 97
      aFlag = ChkGenBankCheque_(acheque)
   Case 89
      aFlag = DoChkBankCheque(acheque, 11, 11, Array(0, 9, 8, 7, 6, 5, 4, 3, 2), 0, 0)
   Case 96
      If Not IsMissing(inChequeType) And inChequeType = 1 Then 'Ιδιωτική
        aFlag = DoChkBankCheque(acheque, 11, 11, Array(0, 9, 8, 7, 6, 5, 4, 3, 2, 0), 0, 1)
      ElseIf Not IsMissing(inChequeType) And inChequeType = 9 Then 'Τραπεζική
        aFlag = DoChkBankCheque(acheque, 11, 11, Array(0, 9, 8, 7, 6, 5, 4, 3, 2), 0, 0)
      Else
        aFlag = True
      End If
   Case 98
      aFlag = DoChkBankCheque(acheque, 11)
   Case 95, 91
        aFlag = False
   
   Case Else
      aFlag = True
      
End Select
If Not aFlag Then aFlag = ChkGenBankCheque_(inCheque)
ChkBankCheque_ = aFlag
End Function

Public Function ChkGenBankCheque_(inCheque As String) As Boolean
' Γενικός έλεγχος CD επιταγών που θα ισχύσει για το ευρω
Dim i As Integer, k As Integer, asum As Long, amod As Integer
    asum = 0: k = 2
    For i = Len(inCheque) - 1 To 1 Step -1
        asum = asum + CInt(Mid(inCheque, i, 1) * k)
        k = k + 1
    Next i
    
    amod = (11 - (asum Mod 11)) Mod 10
    ChkGenBankCheque_ = (CStr(amod) = Right(inCheque, 1))
End Function

Public Function ChkETECheque_(inNum As Long) As Boolean
'Έλεγχος αριθμού επιταγής ΕΤΕ
Dim astr As String, aFlag As Boolean, ares As Long
    aFlag = False
    On Error GoTo ExitPos
    astr = CStr(inNum): astr = Trim(astr)
    If astr = "" Or astr = "0" Then
        aFlag = True
    Else
        aFlag = ChkGenBankCheque_(astr)
    End If

'    If astr = "" Or astr = "0" Then
'        aFlag = True
'    ElseIf ((CLng(astr) < 100000000) Or (CLng(astr) < 900000000 And CLng(astr) >= 800000000)) _
'    And (CLng(astr) < 13000012 Or CLng(astr) > 13250000) And (CLng(astr) < 899985010 Or CLng(astr) > 899995009) Then   'ΔΕΗ
'        If Len(astr) > 1 Then
'            ares = CLng(Left(astr, Len(astr) - 1)) Mod 11
'            If ares = 10 Then ares = 0
'            If CInt(Right(astr, 1)) = ares Then aFlag = True
'        Else
'            aFlag = True
'        End If
'    Else
'        aFlag = ChkGenBankCheque_(astr)
'    End If
ExitPos:
    ChkETECheque_ = aFlag
End Function

Public Function CalcGenBankChequeCD_(inCheque As String) As Long
' Υπολογισμός έλεγχος CD επιταγών που θα ισχύσει για το ευρω
Dim i As Integer, k As Integer, asum As Long, amod As Integer
    asum = 0: k = 2
    For i = Len(inCheque) To 1 Step -1
        asum = asum + CInt(Mid(inCheque, i, 1) * k)
        k = k + 1
    Next i
    
    amod = (11 - (asum Mod 11)) Mod 10
    CalcGenBankChequeCD_ = amod
End Function

Public Function CalcETEChequeCD_(inNum As Long) As Long
'Υπολογισμός Cd αριθμού επιταγής ΕΤΕ
Dim astr As String, aValue As Long, ares As Long
    aValue = 0
    On Error GoTo ExitPos
    astr = CStr(inNum): astr = Trim(astr)
    
    If astr = "" Or astr = "0" Then
    ElseIf ((CLng(astr) < 10000000) Or (CLng(astr) < 90000000 And CLng(astr) >= 80000000)) _
    And (CLng(astr) < 1300001 Or CLng(astr) > 1320000) And (CLng(astr) < 89998501 Or CLng(astr) > 89999500) Then   'ΔΕΗ
        aValue = CLng(astr) Mod 11
        If aValue = 10 Then aValue = 0
    Else
        
        aValue = CalcGenBankChequeCD_(astr)
    End If
ExitPos:
    CalcETEChequeCD_ = aValue
End Function

Public Function ChkFldType_(invalue As String, inValidationCode As Integer)
'01: ΧΩΡΙΣ VALIDATION
'02: Λογαριασμός με CD,
'03: Λογαριασμός χωρίς CD,
'04: Δάνειο με CD,
'05: Δάνειο χωρίς CD,
'06: Ημερομηνία
'07: Αριθμός Εγγραφής
'08: Ειδικός με CD
'09: Γενικός Λογαριασμός Δανείου
'10: Λογαριασμός Καταθέσεων με 1 CD
'11: Τραπεζική Επιταγή
'12: ΕΘΝΟΚΑΡΤΑ
'13: Τραπεζική Εντολή

Dim astr As String, ares As Integer, aFlag As Boolean
On Error GoTo chkFailed
    astr = invalue
    ChkFldType_ = True
    If Trim(astr) = "" Then Exit Function
    
    Select Case inValidationCode
    Case 2
        If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
        If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 11, "0", "L")
        
        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If

        ares = CalcCd2_(Left(astr, 10))
        If CInt(Mid(astr, 11, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 3
    
    Case 10
        If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
'        If Len(astr) = 8 Then astr = fnReadConst_("BranchID") & astr
        If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 10, "0", "L")
        
        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If
        
        ares = CalcCd2_(Left(astr, 10))
            
    Case 11 'Τραπεζική Επιταγή
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
        
    Case 13 'Τραπεζική Εντολή
        astr = Trim(astr)
        aFlag = False
        On Error GoTo chkFailed
        If astr = "" Or astr = "0" Then
            aFlag = True
        Else
            If CLng(astr) >= 550000000 And CLng(astr) <= 600000000 And cVersion >= 20010101 Then
                aFlag = ChkGenBankCheque_(astr)
            Else
            
                ''FBB 17/06/2013
                'aFlag = ChkGenBankCheque_(astr)
                If Len(astr) > 1 Then
                    ares = CLng(Left(astr, Len(astr) - 1)) Mod 11
                    If ares = 10 Then ares = 0
                    If CInt(Right(astr, 1)) = ares Then aFlag = True
                Else
                    aFlag = True
                End If
            End If
        End If
        
        If Not aFlag Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
        
    Case 14 'ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
'        If cVersion >= 20010101 Then
            astr = Trim(astr)
            If Not ChkETECheque_(CLng(astr)) Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
'        End If
    Case 4
        If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
        ares = CalcCd1_(Left(astr, 9), 9)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 5
    Case 6 '06: Ημερομηνία
        If Len(astr) <= 6 Then
            astr = StrPad_(astr, 6, "0", "L")
            astr = Left(astr, 2) & "/" & Mid(astr, 3, 2) & "/" & Right(astr, 2)
        ElseIf Len(astr) = 8 > 6 And Len(astr) <= 8 Then
            astr = StrPad_(astr, 8, "0", "L")
            If Mid(astr, 3, 1) <> "/" Then
                astr = Left(astr, 2) & "/" & Mid(astr, 3, 2) & "/" & Right(astr, 4)
            End If
        End If
        
        If Not IsDate(astr) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
    Case 7
        If Len(astr) < 13 Then astr = StrPad_(astr, 13, "0", "L")
        ares = CalcCd1_(Left(astr, 12), 12)
        If CInt(Right(astr, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 8
    Case 9
        astr = Trim(astr)
        If Len(astr) > 3 Then
            ares = CalcSAccCd(Left(astr, Len(astr) - 1))
            If CInt(Right(astr, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If
    Case 12
        astr = Trim(astr)
        If Not ChkCard(astr) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 21 'Λογαριασμος Καταθέσεων Γερμανία
        If Len(astr) < 8 Then 'astr = StrPad_(astr, 8, "0", "L")
            astr = StrPad_(astr, 7, "0", "L")
            ares = CalcCd1_(Mid(astr, 1, 6), 6)
            If CInt(Mid(astr, 7, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        Else
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
            ares = CalcCd2_(Left(astr, 10))
            If CInt(Mid(astr, 11, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Υποχρεωτικό πεδίο"
                GoTo chkFailed
            End If
        End If
    End Select
    
    ChkFldType_ = True
    
    Exit Function
chkFailed:
    ChkFldType_ = False

End Function


Public Function FormatFldBeforeOut_(invalue As String, ValidationCode As Integer, OutMask As String) As String

    Dim astr As String
    If ValidationCode = 2 Then
        astr = invalue
        If astr <> "" Then
            If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        FormatFldBeforeOut_ = astr
    ElseIf ValidationCode = 3 Then
        astr = invalue
        If astr <> "" Then
            If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
            If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 9, "0", "L")
            astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        FormatFldBeforeOut_ = astr
    ElseIf ValidationCode = 10 Then
        astr = invalue
        If astr <> "" Then
            If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
            If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        FormatFldBeforeOut_ = astr
    ElseIf ValidationCode = 21 Then
        astr = invalue
        If astr <> "" Then
            If Len(astr) < 8 Then astr = cBRANCH & StrPad_(astr, 7, "0", "L") & "0"
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        FormatFldBeforeOut_ = astr
    Else
        If invalue <> "" And OutMask <> "" Then
            FormatFldBeforeOut_ = format(invalue, OutMask)
        Else
            FormatFldBeforeOut_ = invalue
        End If
    End If
End Function

Public Function ChkValidIBAN_(IBANaccount As String) As Integer

Dim stepStr As String
Dim remainingStr As String
Dim stepMod As Integer
Dim Step As Integer
Dim CD As Integer
Dim i As Integer
Dim NullStr As String
Dim LeftStr As String
Dim RightStr As String
Dim ValidIBAN As Integer 'Boolean
Dim IBANChar As String
Dim IBANTableStr As String
Dim IBANNumber As Integer
Dim aPos As Integer
Dim Country As String
Dim CLength As String
Dim cpos As Integer
Dim lengthis As String
Dim CheckNum34Char As String

On Error GoTo chkFailed
 
      Country = "AT,BE,DK,FI,FR,DE,GR,IS,IE,IT,LU,NL,NO,PL,PT,ES,SE,CH,GB,AD,GI,LI,MC,SM,CY,LT,MT,HU,SK,SI,RO,LV,PL,CZ,EE,BG,SA,AE"
      CLength = "20,16,18,18,27,22,27,26,22,27,20,18,15,28,25,24,24,21,22,24,23,21,27,27,28,20,31,28,24,19,24,21,28,24,20,22,24,23"
      
      IBANTableStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      IBANaccount = Trim(IBANaccount)
      IBANaccount = UCase(IBANaccount)
      
      aPos = InStr(1, IBANaccount, " ")
      While aPos > 0
         LeftStr = RTrim(Left(IBANaccount, aPos))
         RightStr = LTrim(Right(IBANaccount, Len(IBANaccount) - aPos))
         IBANaccount = LeftStr & RightStr
         aPos = InStr(1, IBANaccount, " ")
      Wend

      Select Case UCase(Left(IBANaccount, 2))
      Case "AT", "BE", "DK", "FI", "FR", "DE", "GR", "IE", "IS", "IT", "LU", "NL", "NO", "PL", "PT", "ES", "SE", "CH", "GB", "AD", "GI", "LI", "MC", "SM", "CY", "LT", "MT", "HU", "SK", "SI", "RO", "LV", "PL", "CZ", "EE", "BG", "SA", "AE"
         ValidIBAN = 0
      Case Else
         ValidIBAN = 2
         GoTo chkFailed
      End Select

      cpos = InStr(1, Country, Left(IBANaccount, 2))
      lengthis = Mid(CLength, cpos, 2)
      If CInt(lengthis) = Len(IBANaccount) Then
          ValidIBAN = 0
      Else
          ValidIBAN = 2
          GoTo chkFailed
      End If
        
        If ValidIBAN = 0 Then
            CheckNum34Char = Mid(IBANaccount, 3, 2)
            If Not IsNumeric(CheckNum34Char) Then
             ValidIBAN = 2
             GoTo chkFailed
            End If
        End If
      
      If ValidIBAN = 0 Then


          i = 1
    
          LeftStr = Left(IBANaccount, 4)
          RightStr = Right(IBANaccount, Len(IBANaccount) - 4)
          IBANaccount = RightStr & LeftStr
    
          Do While i <= Len(IBANaccount)
             IBANChar = Mid(IBANaccount, i, 1)
             'IBANChar = UCase(IBANChar)
             If Not IsNumeric(IBANChar) Then
                If InStr(1, IBANTableStr, IBANChar) = 0 Then
                    ValidIBAN = 2: Exit Do
                End If
                IBANNumber = InStr(1, IBANTableStr, IBANChar) + 9
                LeftStr = Left(IBANaccount, i - 1)
                RightStr = Right(IBANaccount, Len(IBANaccount) - i)
                IBANaccount = LeftStr & CStr(IBANNumber) & RightStr
                i = i + 1
             End If
             i = i + 1
          Loop
    
          If ValidIBAN = 0 Then
              remainingStr = "RemainingStr"
              Step = 9
        
              While remainingStr <> ""
                 If Len(IBANaccount) < Step Then
                    Step = Len(IBANaccount)
                 End If
                 stepStr = Left(IBANaccount, Step)
                 remainingStr = Right(IBANaccount, Len(IBANaccount) - Step)
                 stepMod = CLng(stepStr) Mod 97
                 IBANaccount = CStr(stepMod) & remainingStr
              Wend
        
              If stepMod = 1 Then
                 ValidIBAN = 0
              Else
                 ValidIBAN = 1
              End If
          End If

      End If

      ChkValidIBAN_ = ValidIBAN
      Exit Function
chkFailed:
    ChkValidIBAN_ = ValidIBAN
      
End Function

Public Function CreateIBAN_(branch As String, account As String) As String
Dim astr As String
Dim stepStr As String
Dim remainingStr As String
Dim stepMod As Integer
Dim Step As Integer
Dim CD As Integer
' curCD = 0, 1, 2 (drx ,in, out)
' Eiaaneaoiio ia check digits
'      branch = Right("0000" & branch, 4)
'      account = Right("000000000000000" & account, 15)
      
'      If IsMissing(curCD) Then
'           account = "0" & account
'      Else
'        Select Case curCD
'         Case 1
'            account = "1" & account
'         Case 2
'            account = "2" & account
'         Case Else
'            account = "0" & account
'        End Select
'      End If

      
    'αν ειδος λογ/μου = προγραμμα 1 ή 2 τοτε λογ/μος = 0000066860168393
    'αν ειδος λογ/μου = προγραμμα 5 τοτε λογ/μος = 1000066860168393
    'branch = '0668
    'account = '0000066860168393 ή 1000066860168393
      
      astr = "011" & branch & account & "162700"
      remainingStr = "RemainingStr"
      Step = 9

      While remainingStr <> ""
        If Len(astr) < Step Then
            Step = Len(astr)
        End If
        stepStr = Left(astr, Step)
        remainingStr = Right(astr, Len(astr) - Step)
        stepMod = CLng(stepStr) Mod 97
        astr = CStr(stepMod) & remainingStr
      Wend

      CD = 98 - stepMod
      astr = "GR" & Right("00" & CStr(CD), 2) & "011" & branch & account
      CreateIBAN_ = astr
      ' Must return 27

End Function

Public Function ChkDocumentNo_(countrycode As String, docno As String) As Boolean
ChkDocumentNo_ = False

    countrycode = Trim(countrycode)
    docno = Trim(docno)

    Dim regExp
    Dim astr As String
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Global = True
    regExp.IgnoreCase = False

    If countrycode = "14" Then 'AUSTRIA
       regExp.Pattern = "^[0-9]{4}/[0-9]{6}$"
       If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "21" Then 'BELGIUM
        regExp.Pattern = "^[0-9]{6} [0-9]{3} [0-9]{2}$"
        If Not regExp.test(docno) Then
            Exit Function
        Else
           If Not IsDate(Mid(docno, 5, 2) & "/" & Mid(docno, 3, 2) & "/" & Mid(docno, 1, 2)) Then Exit Function
        End If
    
    ElseIf countrycode = "52" Then 'CYPRUS
        regExp.Pattern = "^[0-9]{8}[a-zA-Z]$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "53" Then 'CZECH REPUBLIC
        regExp.Pattern = "^([0-9]{3}-[0-9]{10}|CZ[0-9]{10}|[0-9]{11,12}|[0-9]{6}/[0-9]{3,4})$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "56" Then 'DENMARK
        regExp.Pattern = "^([0-9]{2}\.){2}[0-9]{2}/[0-9]{4}$"
        If Not regExp.test(docno) Then
            Exit Function
        Else
           If Not IsDate(Replace(Left(docno, 8), ".", "/")) Then Exit Function
        End If
    
    ElseIf countrycode = "61" Then 'ESTONIA
        regExp.Pattern = "^[0-9]{11}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "66" Then 'FINLAND
        astr = ""
        regExp.Pattern = "^([0-9]{2}/){2}[0-9]{2} ([0-9]|[A-Z]){4}$"
        If regExp.test(docno) Then
            astr = Left(docno, 8)
        Else
            regExp.Pattern = "^([0-9]{2}/){2}[0-9]{4}(A|\+)([0-9]|[A-Z]){4}$"
            If regExp.test(docno) Then astr = Left(docno, 10)
        End If
        If astr <> "" Then
            If Not IsDate(astr) Then Exit Function
        Else
            Exit Function
        End If
    
    ElseIf countrycode = "71" Then 'FRANCE
        regExp.Pattern = "^(1|2)/([0-9]{2}/){3}([0-9]{3}/){2}[0-9]{2}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "54" Then 'GERMANY
    
    ElseIf countrycode = "94" Then 'HUNGARY
        regExp.Pattern = "^([0-9]{10}|([0-9]{3} ){2}[0-9]{3})$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "96" Then 'IRELAND
        regExp.Pattern = "^[0-9]{7}[A-Z]$"
        If Not regExp.test(docno) Then Exit Function
        
    ElseIf countrycode = "104" Then 'ITALY
        regExp.Pattern = "^([A-Z]{3} ){2}[0-9]{2}[A-Z][0-9]{2} [A-Z][0-9]{3}[A-Z]$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "129" Then 'LATVIA
        regExp.Pattern = "^[0-9]{11}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "127" Then 'LITHUANIA
        regExp.Pattern = "^[0-9]{10,11}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "128" Then 'LUXEMBOURG
        regExp.Pattern = "^[0-9]{4} ([0-9]{2} ){2}[0-9]{3}$"
        If Not regExp.test(docno) Then
            Exit Function
        Else
           If Not IsDate(Mid(docno, 9, 2) & "/" & Mid(docno, 6, 2) & "/" & Mid(docno, 1, 4)) Then Exit Function
        End If
    
    ElseIf countrycode = "157" Then 'NETHERLANDS
        regExp.Pattern = "^[0-9]{1,9}$"
        If Not regExp.test(docno) Then Exit Function
        docno = StrPad_(docno, 9, "0", "L")
        
        Dim i, total As Integer
        total = 0
        For i = 1 To 8
            total = total + Mid(docno, i, 1) * (10 - i)
        Next
        If Right(docno, 1) <> total Mod 11 Then Exit Function
    
    ElseIf countrycode = "171" Then 'POLAND
        regExp.Pattern = "^[0-9]{10}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "175" Then 'PORTUGAL
        regExp.Pattern = "^[0-9]{9}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "193" Then 'SLOVAKIA
        regExp.Pattern = "^[0-9]{8}$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "191" Then 'SLOVENIA
        regExp.Pattern = "^([0-9]{8}|SI[0-9]{8})$"
        If Not regExp.test(docno) Then Exit Function
    
    ElseIf countrycode = "64" Then 'SPAIN
    
    ElseIf countrycode = "188" Then 'SWEDEN
        regExp.Pattern = "^[0-9]{6}-[0-9]{4}$"
        If Not regExp.test(docno) Then
            Exit Function
        Else
           If Not IsDate(Mid(docno, 5, 2) & "/" & Mid(docno, 3, 2) & "/" & Mid(docno, 1, 2)) Then Exit Function
        End If
    
    ElseIf countrycode = "73" Then 'UNITED KINGDOM
        regExp.Pattern = "^[0-9]{5} [0-9]{5}$"
        If Not regExp.test(docno) Then
            regExp.Pattern = "^[A-Z]{2}[0-9]{6}[ABCD]?$"
            If Not regExp.test(docno) Then Exit Function
            regExp.Pattern = "^[DFIOQUV]$"
            If regExp.test(Left(docno, 1)) Or regExp.test(Mid(docno, 2, 1)) Then Exit Function
            regExp.Pattern = "^(FY|GB|NK|TN)$"
            If regExp.test(Left(docno, 2)) Then Exit Function
        End If
    
    ElseIf countrycode = "158" Then 'NORWAY
        regExp.Pattern = "^NO ([0-9]{3} ){2}[0-9]{3}$"
        If Not regExp.test(docno) Then
            regExp.Pattern = "^[0-9]{11}$"
            If Not regExp.test(docno) Then
                Exit Function
            Else
                If Not IsDate(Left(docno, 2) & "/" & Mid(docno, 3, 2) & "/" & Mid(docno, 5, 2)) Then Exit Function
            End If
        End If
    
    ElseIf countrycode = "41" Then 'SWITZERLAND

    
    End If

ChkDocumentNo_ = True
End Function

Public Function ChkPensionCD(account As String) As Boolean
    Dim astr As String
    Dim inCd As Integer
    Dim total As Double
    Dim corCd As Double
    astr = Right("000000000000" & account, 12)
    inCd = CInt(Right(astr, 1))
    total = CDbl(Mid(astr, 1, 11))
    corCd = total - (Fix(total / 11) * 11)
    If corCd = 10 Then corCd = 0
    If corCd <> inCd Then
        ChkPensionCD = False
    Else
        ChkPensionCD = True
    End If
End Function
