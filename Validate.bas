Attribute VB_Name = "Validate"
Option Explicit

Dim ValidationInProgress As Boolean

'******************************************************
Public Function Check_Keypress( _
                PIntKeyAscii As Integer, _
                Optional PVarKey_Valid As Variant, _
                Optional PVarKey_Invalid As Variant) _
                As Integer
'******************************************************

' Η Function CheckKeypress δέχεται σε επίπεδο πεδίου
' τα πλήκτρα που πατήθηκαν και ελέγχει αν είναι
' αποδεκτά ή όχι.
'
' Αν είναι αποδεκτά επιστρέφει τον ascii κωδικό
' του αποδεκτού χαρακτήρα, αλλιώς επιστρέφει 0
'
' Παράμετροι :
' PIntKeyAscii ο ascii κωδικός του πλήκτρου που
' πατήθηκε
' PVarKey_valid προαιρετικά ένα string με τους
' αποδεκτούς χαρακτήρες, default όλοι
' PVarKey_Invalid προαιρετικά ένα string με τους
' μή αποδεκτούς χαρακτήρες, default κανείς
'
' Μπορεί να χρησιμοποιηθεί στα KeyPress, KeyUp,
' KeyDown event που χειρίζονται τα πλήκτρα που
' πατήθηκαν
'
' π.χ.
'
' αποδεκτοί χαρακτήρες μόνο αριθμοί και το Backspace
' Private Sub txtACCOUNT_KeyPress(KeyAscii As Integer)
'    KeyAscii = Check_Keypress(KeyAscii, _
'                            "0123456789" + Chr(8) )
'                                                 |
'          παραλείπεται ή προαιρετική παράμετρος -|
' End Sub
'
' αποδεκτοί όλοι οι χαρακτήρες εκτός από
' τους αριθμούς
' Private Sub txtACCOUNT_KeyPress(KeyAscii As Integer)
'     KeyAscii = Check_Keypress(KeyAscii, ,"0123456789")
'                                        |
' παραλείπεται ή προαιρετική παράμετρος -|
' End Sub
                
    Dim MStrKey_Char As String
    
    If IsMissing(PVarKey_Valid) Then
       PVarKey_Valid = ""
    End If

    If IsMissing(PVarKey_Invalid) Then
       PVarKey_Invalid = ""
    End If
    
    MStrKey_Char = Chr(PIntKeyAscii)
    
    If PVarKey_Valid = "" Then
        Check_Keypress = PIntKeyAscii
    Else
       If MStrKey_Char Like "[" + PVarKey_Valid + "]" Then
            Check_Keypress = PIntKeyAscii
        Else
            Check_Keypress = 0
       End If
    End If
    
    If Check_Keypress <> 0 Then
       If PVarKey_Invalid = "" Then
          Check_Keypress = PIntKeyAscii
       Else
          If MStrKey_Char Like "[" + PVarKey_Invalid + "]" Then
             Check_Keypress = 0
          Else
             Check_Keypress = PIntKeyAscii
          End If
       End If
    End If
    
    If Check_Keypress <> 0 Then
        If (Screen.ActiveForm.ActiveControl.BackColor = &HFF&) Then
        Screen.ActiveForm.ActiveControl.BackColor = &H80000005
        End If
    End If
End Function

'***************************************************
Public Function Chk_Digit1(PStrAccount As String, _
                           PStrDigit As String, PBoolDigit As Boolean) _
                           As Boolean
'***************************************************

' Η Function Chk_Digit1 ελέγχει το πρώτο check
' digit του λογαριασμού, αν είναι σωστό επιστρέφει
' True αλλιώς False
'
' Παράμετροι :
' PStrAccount ο εξαψήφιος αριθμός λογαριασμού
' PStrDigit   το πρώτο Check Digit
'pBoolDigit αν θα ελέχθεί ή θα υπολογισθεί (true, false)
'
' π.χ.
'
' If Not Chk_Digit1(MStrAccount, MStrDigit1,true) Then
'    Beep
'    MsgBox "Λανθασμένος Αριθμός Λογαριασμού !!"
' End If
    
    Dim minti As Integer
    Dim MIntJ As Integer
    Dim MIntCd As Integer
    
    minti = MIntJ = MIntCd = 0
    For minti = 7 To 2 Step -1
        MIntJ = MIntJ + 1
        MIntCd = MIntCd + (Val(Mid(PStrAccount, MIntJ, 1)) * minti)
    Next
    MIntCd = MIntCd Mod 11
    If MIntCd = 1 Or MIntCd = 0 Then
       MIntCd = 0
    Else
       MIntCd = 11 - MIntCd
    End If
    
    If Not PBoolDigit Then
        PStrDigit = Trim(Str(MIntCd))
    End If

    Chk_Digit1 = (MIntCd = Val(PStrDigit))
End Function

'***************************************************
Public Function Chk_Digit2(PStrAccount As String, _
                            ByRef PStrDigit As String, PBoolDigit As Boolean) _
                           As Boolean
'***************************************************
' Η Function Chk_Digit2 ελέγχει το δεύτερο check
' digit του λογαριασμού, αν είναι σωστό επιστρέφει
' True αλλιώς False
'
' Παράμετροι :
' PStrAccount ο δεκαψήφιος αριθμός λογαριασμού με
'             μορφή ΚΚΚΛΛΛΛΛΛD
'                   |_||____||
'   κατάστημα--------|    |  |
'   λογαριασμός-----------|  |
'  πρώτο check digit --------|
'
' PStrDigit   το δεύτερο Check Digit
'pBoolDigit αν θα ελέχθεί ή θα υπολογισθεί (true, false)
' π.χ.
'
' If Not Chk_Digit2(MStrAccount, MStrDigit2,True) Then
'    Beep
'    MsgBox "Λανθασμένος Αριθμός Λογαριασμού !!"
' End If
'***************************************************

    Dim MStrS As String
    Dim MDblCd  As Double
    Dim minti As Integer
    
    MStrS = Mid(PStrAccount, 2, 1) + Mid(PStrAccount, 4, 1) + _
            Mid(PStrAccount, 6, 1) + Mid(PStrAccount, 8, 1) + _
            Mid(PStrAccount, 10, 1)
           
    MDblCd = Val(MStrS) * 2
    MStrS = Format(MDblCd)
    MStrS = Mid("000000", 1, (6 - Len(MStrS))) + MStrS
    
    MDblCd = 0
    For minti = 1 To 6
        MDblCd = MDblCd + Val(Mid(MStrS, minti, 1))
    Next
    
    For minti = 1 To 9 Step 2
        MDblCd = MDblCd + Val(Mid(PStrAccount, minti, 1))
    Next
    
    MDblCd = Val(Right(Format(MDblCd), 1))
    If MDblCd <> 0 Then
       MDblCd = 10 - MDblCd
    End If
    
    If Not PBoolDigit Then
        PStrDigit = Trim(Str(MDblCd))
    End If
    
    Chk_Digit2 = (MDblCd = Val(PStrDigit))
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
    
'    Dim minti As Integer
'    Dim MIntJ As Integer
'    Dim MIntCd As Integer
'    Dim strXrhmat As String
'    Dim strDigit As String
'
'    Chk_Xrhmat = True
'    If ValidationInProgress Then Exit Function
'
'    minti = MIntJ = MIntCd = 0
'
'    strXrhmat = pfrmCurrent.txtinput(pIndex).Text
'    strDigit = Mid(strXrhmat, 10, 1)
'
'    For minti = 10 To 2 Step -1
'        MIntJ = MIntJ + 1
'        MIntCd = MIntCd + (Val(Mid(strXrhmat, MIntJ, 1)) * minti)
'    Next
'    MIntCd = MIntCd Mod 11
'    If MIntCd = 1 Or MIntCd = 0 Then
'       MIntCd = 0
'    Else
'       MIntCd = 11 - MIntCd
'    End If
'
'    If MIntCd <> Val(strDigit) Then
'        Chk_Xrhmat = False
'        Call FocusWrongInputField(pfrmCurrent, pIndex, "Λανθασμένο Check Digit!!!")
'    End If

End Function
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
Public Function format_num(PStrTxtnum As String, _
                             Optional PIntDecPos As Variant) _
                             As String
Dim mcurposo As Currency

If IsMissing(PIntDecPos) Then
       PIntDecPos = 2
End If
mcurposo = 10 ^ PIntDecPos
mcurposo = Val(PStrTxtnum) / mcurposo
'format_num = Format(mcurposo, "Standard")
format_num = Format(mcurposo, "#,##0.00\ ;#,##0.00\-")

End Function

Public Function TextToDate(PStrTxtdate As String) _
                           As Date
    
' Η Function TextToDate δέχεται ένα αριθμητικό
' unformated string ημερομηνίας (π.χ. 150395 ή
' 15031995) και επιστρέφει ένα date πεδίο,
' αν η ημερομηνία είναι λανθασμένη τότε επιστρέφει
' EMPTY ημερομηνία
'
' Παράμετροι :
' PStrTxtdate η unformated string ημερομηνία,
' αποδεκτοί 6 ή 8 χαρακτήρες, αν είναι 6 τότε
' η ημερομηνία αφορά τον τρέχοντα αιώνα
'
' π.χ.
'
' txtDate.Text = "15121995"
' TextToDate(txtDATE.Text)--> 15/12/1995 (date)
'
' txtDate.Text = "151295"
' TextToDate(txtDATE.Text)--> 15/12/1995 (date)
'
' txtDate.Text = "1512"
' TextToDate(txtDATE.Text)--> EMPTY (date)
'
' txtDate.Text = "321295"
' TextToDate(txtDATE.Text)--> EMPTY (date)
'
' txtDate.Text = "151795"
' TextToDate(txtDATE.Text)--> EMPTY (date)
'
' Π Ρ Ο Σ Ο Χ Η
'---------------
'
' txtDate.Text = "122795"
' TextToDate(txtDATE.Text)--> 27/12/1995 (date)
'----------------------------------------------
   
    Dim MStrDate As String
    MStrDate = ""
    
    Select Case Len(PStrTxtdate)
           Case 6
                MStrDate = Mid(PStrTxtdate, 1, 2) + "/" + _
                           Mid(PStrTxtdate, 3, 2) + "/" + _
                           Mid(Format(Year(Date)), 1, 2) + _
                           Mid(PStrTxtdate, 5, 2)
           Case 8
                MStrDate = Mid(PStrTxtdate, 1, 2) + "/" + _
                           Mid(PStrTxtdate, 3, 2) + "/" + _
                           Mid(PStrTxtdate, 5, 4)
    End Select
    
    If IsDate(MStrDate) Then
        TextToDate = CDate(MStrDate)
    Else
        TextToDate = Empty
    End If
End Function
Public Function Check_Num_Text(pstrpin, pIndex As Integer, _
                                pKeyAscii As Integer) As Boolean
    Select Case pstrpin(pIndex, 0)
        Case "14", "17", "18", "45", "46"
            Call Text_Keypress(pKeyAscii)
        Case Else
            Call Num_Keypress(pKeyAscii)
    End Select
End Function
Public Sub Num_Keypress(PIntKeyAscii As Integer)

' Η Procedure Num_Keypress πρέπει να χρησιμοποιείται
' στο KeyPress event των πεδίων που δέχονται μόνο
' αριθμητικούς χαρακτήρες
'
' Παράμετροι :
' PIntKeyAscii o ascii κώδικας του χαρακτήρα που
' πατήθηκε
'
' π.χ.
'
' Private Sub txtACCOUNTNUM_KeyPress(KeyAscii As Integer)
'     Call Num_Keypress(KeyAscii)
' End Sub
'
' Private Sub txtAMOUNT_KeyPress(KeyAscii As Integer)
'     Call Num_Keypress(KeyAscii)
' End Sub

    PIntKeyAscii = Check_Keypress(PIntKeyAscii, "0123456789" + Chr(8))
End Sub
    
Public Sub Text_Keypress(PIntKeyAscii As Integer)

' Η Procedure Text_Keypress πρέπει να χρησιμοποι-
' είται στο KeyPress event των πεδίων που δέχονται
' κείμενο
'
' Παράμετροι :
' PIntKeyAscii o ascii κώδικας του χαρακτήρα που
    ' πατήθηκε
'
' π.χ.
'
' Private Sub txtΤΕΧΤ_KeyPress(KeyAscii As Integer)
'     Call Text_Keypress(KeyAscii)
' End Sub

    PIntKeyAscii = Check_Keypress(PIntKeyAscii, , Chr(13))
End Sub

Public Function DATE_GotFocus(frmCurrent As Form, pIndex As Integer) As Boolean

    DATE_GotFocus = True
    
    If ValidationInProgress Then Exit Function
    
    frmCurrent.txtinput(pIndex).Text = Unformat_num(frmCurrent.txtinput(pIndex).Text)

End Function

Public Function StrPad(PString As String, _
                       PIntLen As Integer, _
              Optional PStrChar As Variant, _
              Optional PStrLftRgt As Variant) _
                       As String

' Η Function StrPad δέχεται ένα string πεδίο και
' επιστρέφει ένα string του μήκους που ορίζεται
' προσθέτοντας δεξιά ή αριστερά τον χαρακτήρα που
' ορίζεται όσες φορές χρειάζεται ώστε το Input πεδίο
' να γίνει στο επιθυμητό μήκος
'
' Παράμετροι :
' PString    το input String
' PIntLen    το μήκος του string που θα επιτρέψει
' PStrChar   προαιρετικά ο χαρακτήρας που θα γεμίσει
'            το υπόλοιπο μήκος default <SPACE>
' PStrLftRgt προαιρετικά αν θα προσθέσει χαρακτήρες
'            δεξιά (R) ή αριστερά (L) default δεξιά (L)
'
' π.χ.
' StrPad("12345",10)         -> "     12345"
' StrPad("12345",10, ,"R")   -> "12345     "
' StrPad("12345",10,"0")     -> "0000012345"
' StrPad("12345",10,"0","R") -> "1234500000"
' StrPad("12345",4)          -> "2345"
' StrPad("12345",4, ,"R")    -> "1234"
    
    Dim MString As String
    Dim minti As Integer
    
    If IsMissing(PStrChar) Then
       PStrChar = " "
    End If

    If IsMissing(PStrLftRgt) Then
       PStrLftRgt = "L"
    End If
  
    For minti = 1 To PIntLen
        MString = MString + PStrChar
    Next

    If PStrLftRgt Like "[Ll]" Then
       StrPad = Right(MString + PString, PIntLen)
    Else
       StrPad = Left(PString + MString, PIntLen)
    End If
    
End Function




Public Sub Key_Control(PIntKeycode As Integer)

'***********************************************
' Public Sub Key_Control(PIntKeycode As Integer)
'***********************************************
' Η procedure Key_Control ελέγχει σε επίπεδο Form
' τα πλήκτρα που πατήθηκαν και χειρίζεται το ESC,
' το ENTER, το πάνω και το κάτω βελάκι
'
' Παραμετροι :
' PIntKeycode περιέχει τον κωδικό του πλήκτρου
' που πατήθηκε
'
' πρέπει να μπαίνει στο Form_KeyDown event της form
'
' π.χ.
'
' Private Sub Form_KeyDown(Keycode As Integer, _
'                          Shift As Integer)
'    Call Key_Control(Keycode)
' End Sub
'***************************************************************
Dim I As Integer
    Select Case PIntKeycode
            Case vbKeyEscape
                 ' ESC ξεφορτώνει την form
                  Unload Screen.ActiveForm
            Case vbKeyReturn, vbKeyDown, vbKeySeparator
                ' ENTER ή κάτω βελάκι
                ' μηδενίζει το πληκτρο που πατήθηκε
                ' και στέλνει TAB & END έτσι ώστε
                ' ο cursor να βρεθεί στο τέλος του
                ' επόμενου πεδίου

                 PIntKeycode = 0
                 SendKeys "{TAB}{END}"
            Case vbKeyUp
                ' μηδενίζει το πληκτρο που πατήθηκε
                ' και στέλνει SHIFT TAB & END έτσι
                ' ώστε ο cursor να βρεθεί στο τέλος
                ' του προηγούμενου πεδίου

                 PIntKeycode = 0
                 SendKeys "+{TAB}{END}"
            
        End Select
End Sub

