Attribute VB_Name = "Library"
Option Explicit
' κωδικας που γίνεται export ή χρησημοποιείται από την TrnFrm
Public PrintMsg As String

Const Str928 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρσςτυφχψωάέήϊίόύϋώΆΈΉΊΎΏ"
Const Str437 = "€‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©ª«¬­®―ΰαβγδεζηθικλμνοπ"

Public Function ConvertTo437(astr As String) As String
Dim i As Integer, aSize As Integer, aPos As Integer
Dim bstr As String
    bstr = ""
    If Len(astr) > 0 Then
        aSize = Len(astr)
        For i = 1 To aSize Step 1
            aPos = InStr(1, Str928, Mid(astr, i, 1), 1)
            If aPos > 0 Then bstr = bstr & Mid(Str437, aPos, 1) Else bstr = bstr & Mid(astr, i, 1)
        Next i
    End If
    ConvertTo437 = bstr
End Function

Private Sub clearDoc_()
Dim i As Integer
For i = 0 To DocumentLines - 1: DocLines(i) = String(255, " "): Next i
End Sub

Public Sub PrintDocLines_(owner, Optional inMsg)
Dim DocumentExist As Boolean, i As Integer
    DocumentExist = False
    For i = 1 To LastDocLine + 1
        If Trim(DocLines(i - 1)) <> "" Then DocumentExist = True:  Exit For
    Next i
    If Not DocumentExist Then Exit Sub
    
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
       owner.SPCPanel.LockPrinter
    End If
    
    If IsMissing(inMsg) Then PrintMsg = owner.PrintPromptMessage Else PrintMsg = CStr(inMsg)
    
    If G0Data.count > 0 Then DocForm.Show vbModal _
    Else NBG_MsgBox PrintMsg, True, PrintMsg
    
    If owner Is Nothing Then
    Else
        owner.sbWriteStatusMessage "Εκτύπωση Παραστατικών...."
    End If
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then DocLines(1) = "********* Α Κ Υ Ρ Ο *********"
    
    Dim Status As Long, alldata As String
    
    On Error GoTo errorHandling
    If cPassbookPrinter = 9 Then
        Printer.Orientation = vbPRORPortrait
        Printer.ScaleMode = vbCharacters
        Printer.FontName = "Ub-Courier"
        Printer.FontSize = 10
        For i = 1 To LastDocLine + 1
            If Trim(DocLines(i - 1)) <> "" Then
                Printer.Print Left$(DocLines(i - 1), 80)
            Else
                Printer.Print " "
            End If
        Next i
        Printer.EndDoc
    ElseIf cPassbookPrinter <> 0 Then
        If cPassbookPrinter = 6 Then
            alldata = ""
            For i = 1 To LastDocLine + 1
                If Trim(DocLines(i - 1)) <> "" Then
                    alldata = alldata & "Line" + StrPad_(CStr(i), 3, "0", "L") + "=" + Left(DocLines(i - 1), 80) & Chr(13) & Chr(10)
                End If
            Next i
            
            For i = 131 To 161
                alldata = Replace(alldata, Chr(i), " ")
            Next i
            alldata = Replace(alldata, Chr(127), " ")
            alldata = Replace(alldata, Chr(129), " ")
            alldata = Replace(alldata, Chr(255), " ")
            
            alldata = Replace(alldata, "`", " ")
            alldata = Replace(alldata, "΄", " ")
            alldata = Replace(alldata, Chr(162), "Α")
            alldata = Replace(alldata, Chr(184), "Ε")
            alldata = Replace(alldata, Chr(186), "Ι")
            alldata = Replace(alldata, Chr(188), "Ο")
            alldata = Replace(alldata, Chr(190), "Υ")
            alldata = Replace(alldata, Chr(191), "Ω")
            alldata = Replace(alldata, Chr(218), "Ι")
            alldata = Replace(alldata, Chr(219), "Υ")
            alldata = Replace(alldata, Chr(185), "Η")
                
            Status = owner.SPCPanel.PrintText(alldata)
        ElseIf cPassbookPrinter = 5 Then
            alldata = ""
            For i = 1 To LastDocLine + 1
                alldata = alldata & IIf(Trim(Left(DocLines(i - 1), 80)) = "", vbCrLf, Left(DocLines(i - 1), 80) & vbCrLf)
            Next i
            alldata = "EX$WDRIVERS" & Chr(13) & Chr(10) & "PRINTERSHINE4905" & vbCrLf & alldata
            
            Status = owner.SPCPanel.PrintText(alldata)
        End If
    End If
    Exit Sub
errorHandling:
    MsgBox "Πρόβλημα στην εκτύπωση: " & Err.number & " " & Err.description
End Sub

Public Function PrepareSPCPanel() As SPCPanelX
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
        Set PrepareSPCPanel = CreateObject("SPCPanelXControl.SPCPanelX")
        PrepareSPCPanel.host = cPRINTERSERVER
        PrepareSPCPanel.Port = cPrinterPort
    End If
End Function

Public Sub L2PrintDocLines_(owner, Optional inMsg)
    
    Dim DocumentExist As Boolean, i As Integer

    DocumentExist = False
    For i = 1 To LastDocLine + 1
        If Trim(DocLines(i - 1)) <> "" Then DocumentExist = True:  Exit For
    Next i
    If Not DocumentExist Then Exit Sub

    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
        gPanel.LockPrinter
    End If
            
    
    If IsMissing(inMsg) Then PrintMsg = owner.PrintPromptMessage Else PrintMsg = CStr(inMsg)
    
    NBG_MsgBox PrintMsg, True, PrintMsg
    If owner Is Nothing Then
    Else
        owner.sbWriteStatusMessage "Εκτύπωση Παραστατικών...."
    End If
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then DocLines(1) = "********* Α Κ Υ Ρ Ο *********"
    
Dim Status As Long, alldata As String
    On Error GoTo errorHandling
    If cPassbookPrinter = 9 Then
        Printer.Orientation = vbPRORPortrait
        Printer.ScaleMode = vbCharacters
        Printer.FontName = "Ub-Courier"
        Printer.FontSize = 10
        For i = 1 To LastDocLine + 1
            If Trim(DocLines(i - 1)) <> "" Then
                Printer.Print Left$(DocLines(i - 1), 80)
            Else
                Printer.Print " "
            End If
        Next i
        Printer.EndDoc
    ElseIf cPassbookPrinter <> 0 Then
        If cPassbookPrinter = 5 Then
            alldata = ""
            For i = 1 To LastDocLine + 1
                alldata = alldata & IIf(Trim(Left(DocLines(i - 1), 80)) = "", vbCrLf, Left(DocLines(i - 1), 80) & vbCrLf)
            Next i
            alldata = "EX$WDRIVERS" & Chr(13) & Chr(10) & "PRINTERSHINE4905" & vbCrLf & alldata
            Status = gPanel.PrintText(alldata)
        End If
    End If
    Exit Sub
errorHandling:
    MsgBox "Πρόβλημα στην εκτύπωση: " & Err.number & " " & Err.description
End Sub


Public Sub xClearDoc_()
Dim i As Integer
    If cPassbookPrinter = 0 Then Exit Sub
    For i = 0 To DocumentLines - 1: DocLines(i) = String(255, " "): Next i
    LastDocLine = -1
End Sub

Public Sub xSetDocLine_(inLineNo As Integer, inLineData As String)
    If cPassbookPrinter = 0 Then Exit Sub
    DocLines(inLineNo - 1) = RTrim(inLineData)
    If inLineNo - 1 > LastDocLine Then LastDocLine = inLineNo - 1
End Sub

Public Sub xSetInDocLine_(inLineNo As Integer, inLineData As String, inX As Integer, inW As Integer, inAlign As String)
    If cPassbookPrinter = 0 Then Exit Sub
    Dim astr As String, bstr As String
    astr = DocLines(inLineNo - 1)
    If inAlign = "L" Then inLineData = StrPad_(inLineData, inW, " ", "R") _
    Else inLineData = StrPad_(inLineData, inW, " ", "L")
    bstr = Left$(inLineData, inW)
    astr = Left$(astr, inX - 1) & bstr & Right$(astr, 255 - inX - inW + 1)
        
    DocLines(inLineNo - 1) = RTrim(astr)
    If inLineNo - 1 > LastDocLine Then LastDocLine = inLineNo - 1
End Sub

Public Sub xPrintDoc_(owner, Optional inMsg, Optional PrintOCR, Optional IsCondensed)
Dim Status As Long, i As Integer, alldata As String, DocumentExist As Boolean
Dim OCRFlag As Boolean
    
    If IsMissing(PrintOCR) Then OCRFlag = False Else OCRFlag = PrintOCR
    If IsMissing(IsCondensed) Then IsCondensed = False
    
    If Not (owner Is Nothing) Then
        If IsMissing(inMsg) Then PrintMsg = owner.PrintPromptMessage Else PrintMsg = CStr(inMsg)
    Else
        If IsMissing(inMsg) Then PrintMsg = "ΕΚΤΥΠΩΣΗ ΠΑΡΑΣΤΑΤΙΚΟΥ...." Else PrintMsg = CStr(inMsg)
    End If
    
    'αν καταφέρουμε να εκτυπώσουμε παραστατικο σε Α4 τότε το παρακάτω πρέπει να γίνει If Trim(cLaserDocumentsPrinter) = "" Then
    If Not (Trim(cLaserDocumentsPrinter) <> "" And IsCondensed) Then
        If cPassbookPrinter = 0 Then Exit Sub
        If cPassbookPrinter = 9 Then PrintDocLines_ owner, PrintMsg:   Exit Sub
    End If
    
    DocumentExist = False
    For i = 1 To LastDocLine + 1
        If Trim(DocLines(i - 1)) <> "" Then
            DocumentExist = True
            Exit For
        End If
    Next i
    If Not DocumentExist Then Exit Sub
    
    If cPassbookPrinter = 5 Then
        If Not owner Is Nothing Then owner.SPCPanel.LockPrinter
    End If
    
    If G0Data.count > 0 Then
        DocForm.Show vbModal
    Else
        If Not owner Is Nothing Then
            If G0Data.count > 0 Then DocForm.Show vbModal _
            Else NBG_MsgBox PrintMsg, True, PrintMsg
        End If
    End If
     
    If Not owner Is Nothing Then
        If TypeOf owner Is shine.TRNFrm Then
            owner.sbWriteStatusMessage "Εκτύπωση Παραστατικών...."
        End If
    End If
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then DocLines(1) = "********* Α Κ Υ Ρ Ο *********"
    alldata = ""
    For i = 1 To LastDocLine + 1
        alldata = alldata & Left(DocLines(i - 1), 80) & Chr(13) & Chr(10)
    Next i
    
    'αν καταφέρουμε να εκτυπώσουμε παραστατικο σε Α4 τότε το παρακάτω πρέπει να γίνει If Trim(cLaserDocumentsPrinter) <> "" Then
    If Trim(cLaserDocumentsPrinter) <> "" And IsCondensed Then
        alldata = "EX$LCDRIVERS" & Chr(13) & Chr(10) & "PRINTER" & cLaserDocumentsPrinter & Chr(13) & Chr(10) & alldata
        
'ο παρακάτω κώδικας πρέπει να ενεργοποιηθεί όταν θα εκτυπώνωνται σε laser παραστατικά Α4
'        If IsCondensed Then
'            alldata = "EX$LCDRIVERS" & Chr(13) & Chr(10) & "PRINTER" & cLaserDocumentsPrinter & Chr(13) & Chr(10) & alldata
'        Else
'            alldata = "EX$LDRIVERS" & Chr(13) & Chr(10) & "PRINTER" & cLaserDocumentsPrinter & Chr(13) & Chr(10) & alldata
'        End If
        
        Dim apanel As SPCPanelX
        Set apanel = CreateObject("SPCPanelXControl.SPCPanelX")
        apanel.host = MachineName
        apanel.Port = cPrinterPort
        Status = apanel.PrintText(alldata)
        Set apanel = Nothing
    ElseIf cPassbookPrinter = 5 Then
        If IsCondensed Then
            alldata = "EX$CDRIVERS" & Chr(13) & Chr(10) & "PRINTERSHINE4905" & Chr(13) & Chr(10) & alldata
        ElseIf Not OCRFlag Then
            alldata = "EX$WDRIVERS" & Chr(13) & Chr(10) & "PRINTERSHINE4905" & Chr(13) & Chr(10) & alldata
        Else
            alldata = "EX$OCRDRIVERS" & Chr(13) & Chr(10) & "PRINTERSHINE4905" & Chr(13) & Chr(10) & alldata
        End If
        
        If owner Is Nothing Then
            If gPanel Is Nothing Then
                Dim SPCPanel
                Set SPCPanel = CreateObject("SPCPanelXControl.SPCPanelX")
                SPCPanel.host = cPRINTERSERVER
                SPCPanel.Port = cPrinterPort
                SPCPanel.LockPrinter
                
                NBG_MsgBox PrintMsg, True, "Εκτύπωση"
                Status = SPCPanel.PrintText(alldata)
                
                Set SPCPanel = Nothing
             Else
                gPanel.LockPrinter
                NBG_MsgBox PrintMsg, True, "Εκτύπωση"
                Status = gPanel.PrintText(alldata)
             End If
        Else
            Status = owner.SPCPanel.PrintText(alldata)
        End If
    End If
End Sub

Public Function GetPassbookAmount_(inAmount As Double) As String
' μορφοποίηση ποσού για εκτύπωση σε βιβλιάριο (συμπληρωμένο με *)
Dim astr As String, bstr As String
    astr = Trim(CStr(inAmount))
    If Len(astr) < 3 Then astr = "000" & astr
    bstr = Left(astr, Len(astr) - 2)
    GetPassbookAmount_ = Right(String(10, " ") & "*" & bstr & "," & Right(astr, 2), 13)
End Function

Public Function GetStrAmount_(inAmount As Double, flength As Integer, decpart As Integer) As String
    Dim astr As String, bstr As String
    astr = Trim(CStr(inAmount))
    If Len(astr) <= decpart Then astr = StrPad_(astr, decpart + 1, "0")
    bstr = Left(astr, Len(astr) - decpart) & "," & Right(astr, decpart)
    GetStrAmount_ = StrPad_(bstr, flength + 1)
End Function

Public Function GetPassbookLargeAmount_(inAmount As Double) As String
' μορφοποίηση ποσού για εκτύπωση σε βιβλιάριο (συμπληρωμένο με *)
Dim astr As String, bstr As String
    astr = Trim(CStr(inAmount))
    If Len(astr) < 3 Then astr = "000" & astr
    bstr = Left(astr, Len(astr) - 2)
    GetPassbookLargeAmount_ = Right(String(15, " ") & "*" & bstr & "," & Right(astr, 2), 15)
End Function

Public Sub PrintSinglePassbookLine_(owner As Form, inAccount As String, inTrnDate As String, inTrnCode As Integer, inTrnAmount1 As Double, _
    inTrnAmount2 As Double, fromLine As Integer, fromAmount As Double, inTerm As String)
'inTrnAmount1: ποσο καταθεσης
'inTrnAmount2: ποσο αναληψης

'Εκτύπωση μιάς μόνο γραμμής στο βιβλιάριο για κατάθεση ανάληψη ή ενημέρωση
Dim astr As String, EUROAmount As String
Dim linedata As String, cline As Integer, cAmount As Double
Dim bamount As String
Dim aValue As String, bvalue As String
Dim adate As String, asign As String, aTerminal As String, aaccount As String, aCode As String
Dim Flag02 As Boolean

Flag02 = (CInt(Right(inTrnDate, 2)) >= 1 And CInt(Right(inTrnDate, 2)) < 50)
cline = fromLine
cAmount = fromAmount
If Len(inTrnDate) = 8 Then '01012010
    adate = Mid(inTrnDate, 1, 2) & Mid(inTrnDate, 3, 2) & Mid(inTrnDate, 7, 2) '010110
ElseIf Len(inTrnDate) <= 6 Then
    adate = inTrnDate
End If

If inTerm = "" Then
    aTerminal = owner.GetTerminalID
Else
    aTerminal = inTerm
 End If

aTerminal = Left(aTerminal & String(5, " "), 5)
aaccount = inAccount
clearDoc_


    If cline = 0 Then cline = 1
   

    If inTrnAmount1 > 0 Then
        'ΚΑΤΑΘΕΣΗ
        If inTrnCode = 0 Then aCode = "010" Else aCode = Right("000" & inTrnCode, 3)
        
        bvalue = GetPassbookAmount_(CDbl(inTrnAmount1))
    ElseIf inTrnAmount2 > 0 Then
        'ΑΝΑΛΗΨΗ
         If inTrnCode = 0 Then aCode = "020" Else aCode = Right("000" & inTrnCode, 3)
        
        bvalue = GetPassbookAmount_(CDbl(inTrnAmount2))
    End If
        
    bamount = GetPassbookAmount_(cAmount)
        
    linedata = adate & String(1, " ") & aTerminal & " " & aaccount & " " & aCode
    
    If inTrnAmount1 > 0 Then
        linedata = linedata & String(8, " ") & bvalue & bamount
    ElseIf inTrnAmount2 > 0 Then
        linedata = linedata & String(1, " ") & bvalue & String(7, " ") & bamount
    Else
        linedata = adate & String(2, " ") & aTerminal & "  ΕΝΗΜΕΡΩΘΗΚΕ" & String(24, " ") & bamount
    End If
    DocLines(cline + 2) = "   " & linedata & "."
    
    LastDocLine = cline + 2
    If Trim(DocLines(3)) <> "" Then
        DocLines(2) = String(59, " ") & IIf(Flag02, "ΕΥΡΩ", "ΔΡΧ")
    End If
    
    L2PrintDocLines_ owner
        
End Sub
Public Function L2PrintPassbook_(owner As Form, inTrnType As Integer, fromLine As Integer, fromAmount As Double, inDocument As IXMLDOMElement) As String

Dim cline As Integer, PageCounter As Integer
Dim cAmount As Double
Dim aFlag As Integer, Flag01 As Boolean, Flag02 As Boolean, LastFlag As Integer  '01: εγγραφές 2001, 02: εγγραφές 2002

cline = fromLine
cAmount = fromAmount

clearDoc_
LastFlag = 0
Dim linedata As String
Dim adate As String, aterm As String, abranch As String, aaccount As String, aCode As String, aposo As String, asign As String
Dim bvalue As String, bamount As String, EUROAmount As String
Dim Row As IXMLDOMElement
Dim formatposo As String, decpart As String, intpart As String
   
Dim maxrows As String
Dim EURORate As Double
EURORate = EURORate_

maxrows = inDocument.SelectNodes("//ROWS").length
If (maxrows > 0) Then
    For Each Row In inDocument.SelectNodes("//ROWS")
        If cline = 0 Then
            cline = 1
            fromLine = 1
        ElseIf cline = 1 Then
            If fromLine > 0 Then
                bamount = GetPassbookLargeAmount_(cAmount)
                linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                DocLines(cline + 2) = "  " & linedata
                cline = 2
            Else
                fromLine = 1
            End If
        
        End If
        
        If (Row.selectSingleNode("DCUR").Text = "") Then  'not empty row
            Exit For
        End If
        aFlag = Row.selectSingleNode("DCUR").Text '1:ΔΡΧ 2:ΕΥΡΩ
        
        Flag01 = Flag01 Or (aFlag = "1")
        Flag02 = Flag02 Or (aFlag = "2")
        If Flag01 And Flag02 And LastFlag = 1 Then
            EUROAmount = EUROAmount_(cAmount)
            linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
            DocLines(cline + 2) = "  " & linedata
            DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
            LastDocLine = cline + 3
            
         
            cAmount = Round(cAmount / EURORate)
            PageCounter = PageCounter + 1
            
            If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
            
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
            bamount = GetPassbookLargeAmount_(cAmount)
            linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
            DocLines(cline + 2) = "  " & linedata
            cline = 2
        
        End If
        
        If aFlag = "1" Then LastFlag = 1
        If aFlag = "2" Then LastFlag = 2
        
        
        adate = Row.selectSingleNode("DD_TRANS").Text
        If Len(adate) = 10 Then adate = Left(adate, 2) & Mid(adate, 4, 2) & Right(adate, 2)
        If Len(adate) <> 6 Then adate = Left(adate & String(6, "0"), 6)
        
        'aterm = Decode_Greek_(Right(Row.selectSingleNode("DBRANCH_SND").Text & Row.selectSingleNode("DTERM_ID").Text, 5))
        aterm = Right(Row.selectSingleNode("DBRANCH_SND").Text & Row.selectSingleNode("DTERM_ID").Text, 5)
        aterm = Left(aterm & "     ", 5)
        abranch = Row.selectSingleNode("DBRANCH").Text
        aaccount = abranch & Row.selectSingleNode("DACCOUNT").Text
        aCode = Right("000" & Row.selectSingleNode("DREASON_CODE").Text, 3)
        aposo = Row.selectSingleNode("DUNP_AMOUNT").Text
        
        linedata = " " & adate & " " & aterm & " " & aaccount & " " & aCode
        
        If CDbl(aposo) > 0 Then
          asign = "+"
          bvalue = CDbl(aposo)
        Else
          asign = "-"
          bvalue = (-1) * CDbl(aposo)
          
        End If
        
        cAmount = cAmount + CDbl(asign & "1") * (-1) * CDbl(bvalue)
        bamount = GetPassbookAmount_(cAmount)

      
        formatposo = "": intpart = "": decpart = ""
        formatposo = Right("000000000000" & CStr(bvalue), 12)
        intpart = CStr(CDbl(Left(formatposo, 10)))
        decpart = Right(formatposo, 2)
        bvalue = Right(String(10, " ") & "*" & intpart & "," & decpart, 13)
        
         
        
        If asign = "-" Then
            linedata = linedata & String(8, " ") & bvalue & bamount
        Else
            linedata = linedata & " " & bvalue & String(7, " ") & bamount
        End If

        DocLines(cline + 2) = "  " & linedata & "@"
        
         If CInt(cline) = 20 Then
            LastDocLine = 22
            If Trim(DocLines(3)) <> "" Then
                'If Flag02 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
                If aFlag = 2 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
            End If
            
            PageCounter = PageCounter + 1
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
        Else
            LastDocLine = cline + 2
            cline = cline + 1
        End If
        
         
    Next
    
    'τυπωσε τελευταια σελιδα
    
    adate = format(cPOSTDATE, "DDMMYY")
    Flag02 = True

    If Flag01 And LastFlag = 1 Then
        'EUROAmount = owner.EUROAmount(cAmount)
        EUROAmount = EUROAmount_(cAmount)
        linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
        DocLines(cline + 2) = "  " & linedata
        DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
        LastDocLine = cline + 3
        'cAmount = Round(cAmount / owner.EURORate)
        cAmount = Round(cAmount / EURORate)
        PageCounter = PageCounter + 1
        If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
        If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
        clearDoc_
        cline = 1
        bamount = GetPassbookLargeAmount_(cAmount)
        linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
        DocLines(cline + 2) = "  " & linedata
        cline = 2
        LastFlag = 2
    End If

   Dim aTerminal   As String
   aTerminal = cTERMINALID
   
   If inTrnType = 0 Or inTrnType = 3 Then 'ενημερωση,εξοφληση
        If cline = 1 Then
             If fromLine > 0 Then
                 Dim lcldata As String
                 bamount = GetPassbookLargeAmount_(cAmount)
                 lcldata = String(34, " ") & "EK MET          " & Right(bamount, 13)
                 DocLines(cline + 2) = " " & lcldata
                 cline = 2
             Else
                 fromLine = 1
             End If
        End If
        
        If inTrnType = 0 Then 'ενημερωση
            bamount = GetPassbookAmount_(cAmount)
            linedata = " " & adate & String(1, " ") & aTerminal & "  ΕΝΗΜΕΡΩΘΗΚΕ" & String(23, " ") & bamount & "@"
        ElseIf inTrnType = 3 Then 'εξοφληση
            linedata = " " & adate & String(1, " ") & aTerminal & " " & "ΕΞΟΦΛΗΘΗΚΕ"
        
        End If
        DocLines(cline + 2) = "  " & linedata
        LastDocLine = cline + 2
   
   End If
   
    If Trim(DocLines(3)) <> "" Then
       DocLines(2) = String(59, " ") & "ΕΥΡΩ"
    End If
    PageCounter = PageCounter + 1
    If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
    Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
 
End If
    L2PrintPassbook_ = cAmount
End Function

Public Sub PrintPassbook_(owner As Form, inAccount As String, inTrnType As Integer, inTrnCode As String, inTrnAmount As Double, _
    fromLine As Integer, fromAmount As Double, Optional PrintEUROText As Integer)
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

' εκτύπωση βιβλιαρίου πρόγραμμα 1
Dim i As Integer, k As Integer, astr As String
Dim linedata As String, cline As Integer, cAmount As Double
Dim bamount As String, EUROAmount As String
Dim aValue As String, bvalue As String
Dim adate As String, asign As String, aTerminal As String, aaccount As String, aCode As String

Dim PageCounter, LastDate As String, Flag01 As Boolean, Flag02 As Boolean, LastFlag As Integer   '01: εγγραφές 2001, 02: εγγραφές 2002
Dim aFlag As String, useLineG5 As Boolean
useLineG5 = True

cline = fromLine
cAmount = fromAmount

clearDoc_
LastFlag = 0
If owner.ListData.count >= 1 Then
    For i = 1 To owner.ListData.count
        astr = owner.ListData.item(i)
        If Left(astr, 1) = "7" Or Left(astr, 1) = "5" Then
            LastDate = adate
            astr = owner.ListData.item(i)
            
            If Left(astr, 1) = "7" Or (Left(astr, 1) = "5" And inTrnType <> 0) Then   ' useLineG5) Then
                If cline = 0 Then
                        cline = 1
                        ''fromLine = 1
                ElseIf cline = 1 Then
                    If fromLine > 0 Then
                        bamount = GetPassbookLargeAmount_(cAmount)
                        linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                        DocLines(cline + 2) = "  " & linedata
                        cline = 2
                    Else
                        fromLine = 1
                    End If
                End If
            End If
            
            If Left(astr, 1) = "7" Then
                If cline = 1 And fromLine = 0 Then
                   fromLine = 1
                End If
                useLineG5 = False
                
                astr = Mid(astr, 6, Len(astr) - 5)
                
                asign = Mid(astr, 13, 1)
                adate = Mid(astr, 15, 6)
                aFlag = Mid(astr, 21, 1) '1:ΔΡΧ 2:ΕΥΡΩ
    '--------------------------------------------
                Flag01 = Flag01 Or (aFlag = "1")
                Flag02 = Flag02 Or (aFlag = "2")
                If Flag01 And Flag02 And LastFlag = 1 Then
                    EUROAmount = owner.EUROAmount(cAmount)
                    linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
                    DocLines(cline + 2) = "  " & linedata
                    DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
                    LastDocLine = cline + 3
                    
                    cAmount = Round(cAmount / owner.EURORate)
                    PageCounter = PageCounter + 1
                    
                    If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
                    
                    If PageCounter = 1 Then PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
                    Else PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
                    clearDoc_
                    cline = 1
                    bamount = GetPassbookLargeAmount_(cAmount)
                    linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                    DocLines(cline + 2) = "  " & linedata
                    cline = 2
                
                End If
                
                If aFlag = "1" Then LastFlag = 1
                If aFlag = "2" Then LastFlag = 2
    '--------------------------------------------
                
                aValue = Mid(astr, 1, 12) '2 δεκαδικά πάντα
                
                bvalue = CStr(CDbl(Left(aValue, 10)))
                
                bvalue = Right(String(10, " ") & "*" & bvalue & "," & Right(aValue, 2), 13)
                
                aTerminal = Mid(astr, 22, 5)
                aaccount = Mid(astr, 28, 10)
                aCode = Mid(astr, 39, 3)     'κωδικος αιτιολογίας
                
                cAmount = cAmount + CDbl(asign & "1") * (-1) * CDbl(aValue)
                
                linedata = " " & adate & " " & aTerminal & " " & aaccount & " " & aCode
                
                bamount = GetPassbookAmount_(cAmount)
                
                If asign = "-" Then
                    linedata = linedata & String(8, " ") & bvalue & bamount
                Else
                    linedata = linedata & " " & bvalue & String(7, " ") & bamount
                End If
                
                DocLines(cline + 2) = "  " & linedata & "@"
                If CInt(cline) = 20 Then
                    LastDocLine = 22
                    If Trim(DocLines(3)) <> "" Then
                        If Flag02 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
                    End If
                    
                    PageCounter = PageCounter + 1
                    If PageCounter = 1 Then PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
                    Else PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
                    clearDoc_
                    cline = 1
                Else
                    LastDocLine = cline + 2
                    cline = cline + 1
                End If
            ElseIf Left(astr, 1) = "5" Then
                
                adate = owner.GetPostDate_U6
                Flag02 = True
                
                If Flag01 And LastFlag = 1 Then
                    EUROAmount = owner.EUROAmount(cAmount)
                    linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
                    DocLines(cline + 2) = "  " & linedata
                    DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
                    LastDocLine = cline + 3
                    cAmount = Round(cAmount / owner.EURORate)
                    PageCounter = PageCounter + 1
                    If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
                    If PageCounter = 1 Then PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" Else PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
                    clearDoc_
                    cline = 1
                    bamount = GetPassbookLargeAmount_(cAmount)
                    linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                    DocLines(cline + 2) = "  " & linedata
                    cline = 2
                    LastFlag = 2
                End If
                
                aTerminal = owner.GetTerminalID
                aaccount = Left(inAccount, 10)
                
                If inTrnType = 1 Then
                    aCode = Right("010" & inTrnCode, 3)
                ElseIf inTrnType = 2 Then
                    aCode = Right("020" & inTrnCode, 3)
                ElseIf inTrnType = 3 Then
                    aCode = "056"
                End If
                
                astr = Mid(astr, 6, Len(astr) - 5)
                
                If inTrnType = 3 Then
                    bvalue = GetPassbookAmount_(CDbl(Mid(astr, 1, 12)))
                Else
                    bvalue = GetPassbookAmount_(inTrnAmount)
                End If
                
                If inTrnType <> 0 Then cAmount = CDbl(Mid(astr, 1, 12))
                
                bamount = GetPassbookAmount_(cAmount)
                EUROAmount = owner.EUROAmount(cAmount)
                
                linedata = " " & adate & " " & aTerminal & " " & aaccount & " " & aCode
                If inTrnType = 1 Then
                    linedata = linedata & String(8, " ") & bvalue & bamount
                ElseIf inTrnType = 2 Then
                    linedata = linedata & " " & bvalue & String(7, " ") & bamount
                ElseIf inTrnType = 3 Then
                    linedata = linedata & " " & bvalue & String(15, " ") & "*0,00"
                ElseIf inTrnType = 0 Then
                    linedata = " " & adate & String(1, " ") & aTerminal & "  ΕΝΗΜΕΡΩΘΗΚΕ" & String(23, " ") & bamount
                End If
                
                If cline = 1 Then
                    If fromLine > 0 Then
                        Dim lcldata As String
                        bamount = GetPassbookLargeAmount_(cAmount)
                        lcldata = String(34, " ") & "EK MET          " & Right(bamount, 13)
                        DocLines(cline + 2) = " " & lcldata
                        cline = 2
                    Else
                        fromLine = 1
                    End If
                End If
                DocLines(cline + 2) = "  " & linedata & "@"
                LastDocLine = cline + 2
                
                If inTrnType = 3 Then
                    cline = cline + 1
                    If cline > 20 Then
                    
                        If Trim(DocLines(3)) <> "" Then
                            DocLines(2) = String(59, " ") & "ΕΥΡΩ"
                        End If
                        PageCounter = PageCounter + 1
                        If PageCounter = 1 Then PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
                        Else PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
                        
                        cline = 1
                        
                        linedata = String(34, " ") & "EK MET          " & "*********0,00"
                        DocLines(cline + 2) = "  " & linedata
                        LastDocLine = cline + 2
                        cline = cline + 1
                    End If
                    linedata = adate & String(2, " ") & aTerminal & " " & "ΕΞΟΦΛΗΘΗΚΕ"
                    DocLines(cline + 2) = "  " & linedata
                    LastDocLine = cline + 2
                End If
                
                If Trim(DocLines(3)) <> "" Then
                    DocLines(2) = String(59, " ") & "ΕΥΡΩ"
                End If
                PageCounter = PageCounter + 1
                If PageCounter = 1 Then PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
                Else PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
                
                Dim ChkDate As String
                ChkDate = owner.GetPostDate_U8
            End If
        End If
    Next
End If
End Sub

Public Sub PrintPassbook5_(owner As Form, inAccount As String, inTrnType As Integer, _
    inTrnCode As String, inTrnAmount As Double, inTrnDRXAmount As Double, _
    fromLine As Integer, fromAmount As Double, Optional inTrnEuroFinalAmount As Double)
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

'εκτύπωση βιβλιαρίου πρόγραμμα 5

Dim i As Integer, k As Integer, astr As String
Dim linedata As String, cline As Integer, cAmount As Double
Dim bamount As String, eamount As String
Dim aValue As String, bvalue As String
Dim adate As String, asign As String, aTerminal As String, aaccount As String, aCode As String
Dim lastamount As Double

    cline = fromLine
    cAmount = fromAmount
    
    clearDoc_
    linedata = ""
    
    If owner.ListData.count >= 1 Then
        For i = 1 To owner.ListData.count
            If cline = 0 Then
                cline = 1
            ElseIf cline = 1 Then
                If fromLine > 0 Then
                    If InStr(1, linedata, "           ΙΣΟΤΙΜΟ ΣΕ ") Then bamount = GetPassbookLargeAmount_(lastamount) _
                    Else bamount = GetPassbookLargeAmount_(cAmount)
                    linedata = String(34, " ") & "EK MET " & bamount
                    DocLines(cline + 2) = " " & linedata
                    cline = 2
                Else
                    fromLine = 1
                End If
            End If
            lastamount = cAmount
            astr = owner.ListData.item(i)
'            If (owner.Module28PoolLink And Left(astr, 1) <> "5") Then
'                astr = "70000" & astr
'            End If
            If Left(astr, 1) = "7" Then
                If Len(astr) >= 81 Then
                    If Mid(astr, 81, 1) = "1" Then
                        astr = Mid(astr, 6, Len(astr) - 5)
                        linedata = Left(astr, 75)
                        'cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                        If Mid(astr, 69, 1) = "," Then
                            cAmount = CDbl(Mid(astr, 56, 13) & Mid(astr, 70, 2))
                        Else
                            cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                        End If
                    ElseIf Mid(astr, 81, 1) = "2" Then
                        linedata = Right(astr, Len(astr) - 5)
                    Else
                        astr = Mid(astr, 6, Len(astr) - 5)
                        linedata = Left(astr, 75)
                        'cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                        If Mid(astr, 69, 1) = "," Then
                            cAmount = CDbl(Mid(astr, 56, 13) & Mid(astr, 70, 2))
                        Else
                            cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                        End If
                        
                    End If
                Else
                    astr = Mid(astr, 6, Len(astr) - 5)
                    linedata = Left(astr, 75)
                    'cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                    If Mid(astr, 69, 1) = "," Then
                        cAmount = CDbl(Mid(astr, 56, 13) & Mid(astr, 70, 2))
                    Else
                        cAmount = CDbl(Mid(astr, 56, 16) & Mid(astr, 73, 2))
                    End If
                End If
                DocLines(cline + 2) = " " & linedata
                LastDocLine = cline + 2
                If CInt(cline) = 20 Then
                    LastDocLine = 22
                    PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                    clearDoc_
                    cline = 1
                Else
                    LastDocLine = cline + 2
                    cline = cline + 1
                End If
            ElseIf Left(astr, 1) = "5" Then
                Dim CurrName As String
                CurrName = Right(astr, 3)
                If inTrnType = 0 Then
                '    DocLines(cline + 2) = "  " & linedata
                '    LastDocLine = cline + 2
                    
                    
                    adate = owner.GetPostDate_U6
                    aTerminal = StrPad_(cTERMINALID, 5, " ", "L")
                    
                    linedata = adate & aTerminal & " " & " ΕΝΗΜΕΡΩΘΗΚΕ *****"
    
                    DocLines(cline + 2) = " " & linedata
                    LastDocLine = cline + 2
                    
                    PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                    
                Else
                    adate = owner.GetPostDate_U6
                    aTerminal = StrPad_(cTERMINALID, 5, " ", "L")
                    bvalue = CStr(CDbl(Trim(Mid(astr, 73, 13)))) + "," + Mid(astr, 86, 2)
                    If inTrnType = 1 Then
                        bvalue = StrPad_(bvalue, 18, " ", "L") + "+"
                        cAmount = cAmount + CDbl(Mid(astr, 73, 13) & Mid(astr, 86, 2))
                    ElseIf inTrnType = 2 Then
                        bvalue = StrPad_(bvalue, 18, " ", "L") + "-"
                        cAmount = cAmount - CDbl(Mid(astr, 73, 13) & Mid(astr, 86, 2))
                    End If
                    
                    If (inTrnCode = 80) Or (inTrnCode = 180) Or (inTrnCode = 89) Or (inTrnCode = 189) Then
                        aValue = Trim(CStr(inTrnDRXAmount))
                    Else
                        aValue = Trim(CStr(inTrnAmount))
                    End If
                    
                    If Len(aValue) > 2 Then
                        aValue = Left(aValue, Len(aValue) - 2) & "," & Right(aValue, 2)
                    Else
                        aValue = "0," & StrPad_(aValue, 2, "0", "L")
                    End If
                    linedata = adate & aTerminal & " " & Mid(astr, 120, 3) & _
                        StrPad_(aValue, 17, " ", "L") & " " & _
                        StrPad_(inTrnCode, 3, "0", "L") & bvalue
                        
                    bvalue = CStr(cAmount)
                    If Len(bvalue) < 3 Then bvalue = Right("000" & bvalue, 3)
                    bvalue = StrPad_(bvalue, 18, " ", "L")
                    
                    linedata = linedata & Left(bvalue, 16) & "," & Right(bvalue, 2)
    
                    DocLines(cline + 2) = " " & linedata
                    LastDocLine = cline + 2
                    
                    If Not IsMissing(inTrnEuroFinalAmount) And inTrnEuroFinalAmount <> 0 Then 'Ισότιμο ΕΥΡΩ
                        If cline = 20 Then
                            LastDocLine = 22
                            PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                            clearDoc_
                            cline = 1
                        Else
                            LastDocLine = cline + 2
                            cline = cline + 1
                        End If
                        
                        If cline = 1 Then
                            If fromLine > 0 Then
                                bamount = GetPassbookLargeAmount_(cAmount)
                                linedata = String(34, " ") & "EK MET " & bamount
                                DocLines(cline + 2) = " " & linedata
                                cline = 2
                            Else
                                fromLine = 1
                            End If
                        End If
                        
                        eamount = CStr(inTrnEuroFinalAmount)
                        If Len(eamount) < 3 Then eamount = Right("000" & eamount, 3)
                        eamount = StrPad_(eamount, 19, " ", "L")

'                        eamount = StrPad_(CStr(inTrnEuroFinalAmount), 19, " ", "L")
                        
                        linedata = "   " & String(22, " ") & _
                                    IIf(Date >= DateSerial(2001, 10, 29) Or owner.CurVer >= 20011029, "ΙΣΟΤΙΜΟ ΣΕ " & CurrName & ":" & String(18, " "), owner.EUROText) & _
                                    Right(Left(eamount, Len(eamount) - 2) & "," & Right(eamount, 2), 16)
                        'owner.EUROAmount5(inTrnEuroFinalAmount)
                        
                        DocLines(cline + 2) = " " & linedata
                        LastDocLine = cline + 2
                        PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                        
                    Else
                        PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                    End If
                End If
            End If
        Next
    End If
End Sub
Public Function L2PrintPassbook5_(owner As Form, inAccount As String, inTrnType As Integer, _
    inTrnCode As String, inTrnAmount As Double, inTrnDRXAmount As Double, _
    fromLine As Integer, fromAmount As Double, _
    inDocument As IXMLDOMElement, failedtrnflag As Boolean, _
    Optional inTrnEuroFinalAmount As Double) As String
    
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

'εκτύπωση βιβλιαρίου πρόγραμμα 5

Dim i As Integer, k As Integer, astr As String
Dim linedata As String, cline As Integer, cAmount As Double
Dim bamount As String, eamount As String
Dim aValue As String, bvalue As String
Dim adate As String, asign As String, aTerminal As String, aaccount As String, aCode As String

cline = fromLine
cAmount = fromAmount

clearDoc_
linedata = ""
    
Dim maxrows As String
Dim Row As IXMLDOMElement
Dim rowcounter As Integer
rowcounter = 0

Dim print_type As String, trans_dt As String, value_dt As String, term_dep_exp_dt As String, rsn_cd As String
Dim ent_amnt As String, aent_amnt As Double, aent_amntsign As String
Dim psbk_ball As String, apsbk_ball As Double, apsbk_ballsign As String
Dim trans_amnt As String, atrans_amnt As Double, atrans_amntsign As String
Dim int_rate As String, aint_rate As Double
Dim cur_iso As String, send_br As String, term_id As String, star_id As String
Dim parite As String, aparite As Double

If gPanel Is Nothing Then
    Set gPanel = New GlobalSPCPanel
End If

maxrows = inDocument.SelectNodes("//ROWS").length
If (maxrows > 0) Then
    For rowcounter = 0 To maxrows
        
        If cline = 0 Then
            cline = 1
        ElseIf cline = 1 Then
            If fromLine > 0 Then
                bamount = GetPassbookLargeAmount_(cAmount)
                linedata = String(34, " ") & "EK MET " & bamount
                DocLines(cline + 2) = " " & linedata
                cline = 2
            Else
                fromLine = 1
            End If
        End If
            
        If rowcounter < maxrows Then
            Set Row = inDocument.SelectNodes("//ROWS")(rowcounter)
            
            print_type = Row.selectSingleNode("PRINT_TYPE").Text
            trans_dt = Row.selectSingleNode("TRANS_DT").Text & String(10, " ")
            value_dt = Row.selectSingleNode("VALUE_DT").Text & String(10, " ")
            term_dep_exp_dt = Row.selectSingleNode("TERM_DEP_EXP_DT").Text & String(10, " ")
            rsn_cd = Right("000" & Row.selectSingleNode("RSN_CD").Text, 3)
            ent_amnt = Row.selectSingleNode("ENT_AMNT").Text
            psbk_ball = Row.selectSingleNode("PSBK_BAL").Text
            cur_iso = Row.selectSingleNode("CUR_ISO").Text
            trans_amnt = Row.selectSingleNode("TRANS_AMNT").Text
            send_br = Row.selectSingleNode("SEND_BR").Text & String(3, " ")
            term_id = Row.selectSingleNode("TERM_ID").Text & String(2, " ")
            int_rate = Row.selectSingleNode("INT_RATE").Text
            parite = Row.selectSingleNode("PARITE").Text
            star_id = Row.selectSingleNode("STAR_IND").Text
            
            trans_dt = Mid(trans_dt, 1, 2) & Mid(trans_dt, 4, 2) & Mid(trans_dt, 9, 2)
            value_dt = Mid(value_dt, 1, 2) & Mid(value_dt, 4, 2) & Mid(value_dt, 9, 2)
            term_dep_exp_dt = Mid(term_dep_exp_dt, 1, 2) & Mid(term_dep_exp_dt, 4, 2) & Mid(term_dep_exp_dt, 9, 2)
            send_br = Mid(send_br, 1, 3)
            term_id = Mid(term_id, 1, 2)
            
            atrans_amnt = CDbl(trans_amnt)
            atrans_amntsign = " "
            If atrans_amnt < 0 Then
                atrans_amntsign = "-"
            End If
            aent_amnt = CDbl(ent_amnt)
            aent_amntsign = " "
            If aent_amnt > 0 Then
                aent_amntsign = "+"
            Else
                aent_amntsign = "-"
            End If
            apsbk_ball = CDbl(psbk_ball)
            apsbk_ballsign = " "
            If apsbk_ball < 0 Then
                apsbk_ballsign = "-"
            End If
            aint_rate = CDbl(int_rate)
            
            aparite = CDbl(parite)
                        
            If print_type = "0" Then
                linedata = gFormat_("%38ST% ΥΠΟΛΟΙΠΟ ΣΕ %3ST%:%16SR%", _
                    Array("", cur_iso, GetStrAmount_(aent_amnt, 15, 2)))
                
                DocLines(cline + 2) = " " & linedata
                LastDocLine = cline + 2
                If CInt(cline) = 20 Then
                    LastDocLine = 22
                    L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                    clearDoc_
                    cline = 1
                Else
                    LastDocLine = cline + 2
                    cline = cline + 1
                End If
                             
                linedata = gFormat_("** ΝΕΟ ΝΟΜΙΣΜΑ EUR ΤΙΜΗ %14SR% ΥΠΟΛΟΙΠΟ ΣΕ EUR:%16SR%", _
                    Array(GetStrAmount_(aparite, 10, 6), GetStrAmount_(atrans_amnt, 15, 2)))
                
                cAmount = atrans_amnt
            ElseIf print_type = "1" Then
                linedata = gFormat_("%6ST% %3ST%%2ST%%3ST% %16SR% %3ST% %17SR%%1ST% %18SR%%1ST% %1ST%", _
                    Array(trans_dt, send_br, term_id, cur_iso, GetStrAmount_(atrans_amnt, 15, 2), rsn_cd, GetStrAmount_(Abs(aent_amnt), 15, 2), aent_amntsign, GetStrAmount_(Abs(apsbk_ball), 15, 2), apsbk_ballsign, star_id))
                
                cAmount = apsbk_ball
            ElseIf print_type = "2" Then
                linedata = gFormat_("ΗΜΕΡΟΜΗΝΙΑ ΙΣΧΥΟΣ: %6ST% ΕΠΙΤΟΚΙΟ: %6SR%%35SR%", _
                    Array(value_dt, GetStrAmount_(aint_rate, 5, 3), star_id))
                
                cAmount = apsbk_ball
            ElseIf print_type = "3" Then
                linedata = gFormat_("%6ST% %3ST%%2ST% ΑΝΑΝΕΩΣΗ V %6ST% ΛΗΞΗ %6ST% ΕΠ %6SR% ΠΟΣΟ %16SR%%1ST% %1ST%", _
                    Array(trans_dt, send_br, term_id, value_dt, term_dep_exp_dt, GetStrAmount_(aint_rate, 5, 3), GetStrAmount_(Abs(atrans_amnt), 15, 2), atrans_amntsign, star_id))
                
                cAmount = apsbk_ball
            End If
                
            DocLines(cline + 2) = " " & linedata
            LastDocLine = cline + 2
            If CInt(cline) = 20 Then
                LastDocLine = 22
                L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                clearDoc_
                cline = 1
            Else
                LastDocLine = cline + 2
                cline = cline + 1
            End If
        
        Else
            If Not failedtrnflag Then
                If inTrnType = 0 Then
                    adate = format(cPOSTDATE, "DDMMYY")
                    aTerminal = StrPad_(cTERMINALID, 5, " ", "L")
    
                    linedata = adate & aTerminal & " " & " ΕΝΗΜΕΡΩΘΗΚΕ *****"
    
                    DocLines(cline + 2) = " " & linedata
                    LastDocLine = cline + 2
                End If
            End If
            
            L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
            clearDoc_
        End If

        Next
    End If
    
'    If Not (gPanel Is Nothing) Then
'        gPanel.UnlockPrinter
'    End If
    Set gPanel = Nothing

    L2PrintPassbook5_ = cAmount
End Function

Public Function GetBatchList_(owner As Form) As Boolean
' επιστρέφει στο ListData collection τα περιεχόμενα του batch
    GetBatchList_ = False
End Function

Public Function ClearBatchList_(lastline As Long) As Boolean
' διαγράφει τις εγγραφές του batch με sn <= LastLine
    ClearBatchList_ = False
End Function

Public Sub DecodeRange_(inRange As String, ByRef lowvalue As Integer, ByRef highvalue As Integer)
Dim aPos As Integer, astr As String, bstr As String
    aPos = InStr(inRange, "-")
    If aPos > 1 Then astr = Left(inRange, aPos - 1) Else astr = "0"
    If aPos < Len(inRange) Then bstr = Right(inRange, Len(inRange) - aPos) Else bstr = "0"
    If IsNumeric(astr) Then lowvalue = CInt(astr) Else lowvalue = 0
    If IsNumeric(bstr) Then highvalue = CInt(bstr) Else highvalue = 0
End Sub

Public Function L2PrintPassbookVersion3_(owner As Form, inTrnType As Integer, fromLine As Integer, fromAmount As Double, inDocument As IXMLDOMElement, branch As String, account As String) As String

Dim cline As Integer, PageCounter As Integer
Dim cAmount As Double
Dim aFlag As Integer, Flag01 As Boolean, Flag02 As Boolean, LastFlag As Integer  '01: εγγραφές 2001, 02: εγγραφές 2002


 'aflag = Mid(astr, 21, 1) '1:ΔΡΧ 2:ΕΥΡΩ
 
cline = fromLine
cAmount = fromAmount

clearDoc_
LastFlag = 0
Dim linedata As String
Dim adate As String, aterm As String, abranch As String, aaccount As String, aCode As String, aposo As String, asign As String
Dim bvalue As String, bamount As String, EUROAmount As String
Dim Row As IXMLDOMElement
Dim formatposo As String, decpart As String, intpart As String
   
Dim maxrows As String
Dim EURORate As Double
EURORate = EURORate_

Dim Cur As String

If gPanel Is Nothing Then
    Set gPanel = New GlobalSPCPanel
End If

maxrows = inDocument.SelectNodes("//ROWS").length
If (maxrows > 0) Then
    For Each Row In inDocument.SelectNodes("//ROWS")
        If cline = 0 Then
            cline = 1
            fromLine = 1
        ElseIf cline = 1 Then
            If fromLine > 0 Then
                bamount = GetPassbookLargeAmount_(cAmount)
                linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                DocLines(cline + 2) = "  " & linedata
                cline = 2
            Else
                fromLine = 1
            End If
        
        End If
        
        If (Row.selectSingleNode("CURRENCY").Text = "") Then  'not empty row
            Exit For
        End If
        Cur = Row.selectSingleNode("CURRENCY").Text
        If Cur = "GRD" Then aFlag = 1
        If Cur = "EUR" Then aFlag = 2
        'aflag = Row.selectSingleNode("DCUR").Text '1:ΔΡΧ 2:ΕΥΡΩ
        
        Flag01 = Flag01 Or (aFlag = "1")
        Flag02 = Flag02 Or (aFlag = "2")
        If Flag01 And Flag02 And LastFlag = 1 Then

            EUROAmount = EUROAmount_(cAmount)
            linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
            DocLines(cline + 2) = "  " & linedata
            DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
            LastDocLine = cline + 3
            
         
            'cAmount = Round(cAmount / owner.EURORate)
            cAmount = Round(cAmount / EURORate)
            PageCounter = PageCounter + 1
            
            If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
            
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
            bamount = GetPassbookLargeAmount_(cAmount)
            linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
            DocLines(cline + 2) = "  " & linedata
            cline = 2
        
        End If
        
        If aFlag = "1" Then LastFlag = 1
        If aFlag = "2" Then LastFlag = 2
        
        
        adate = Row.selectSingleNode("D_TRANS").Text
        If Len(adate) = 10 Then adate = Left(adate, 2) & Mid(adate, 4, 2) & Right(adate, 2)
        If Len(adate) <> 6 Then adate = Left(adate & String(6, "0"), 6)
        
        
        aterm = Right(Row.selectSingleNode("BRANCH_SND").Text & Row.selectSingleNode("TERM_ID").Text, 5)
        aterm = Left(aterm & "     ", 5)
        'abranch = Row.selectSingleNode("DBRANCH").Text
        'aaccount = abranch & Row.selectSingleNode("DACCOUNT").Text
        abranch = branch
        aaccount = account
        
        aCode = Right("000" & Row.selectSingleNode("REASON_CODE").Text, 3)
        aposo = Row.selectSingleNode("UNP_AMOUNT").Text
        
        'linedata = " " & adate & " " & aterm & " " & aaccount & " " & aCode
        linedata = " " & adate & " " & aterm & " " & branch & aaccount & " " & aCode
        
        If CDbl(aposo) > 0 Then
          asign = "+"
          bvalue = CDbl(aposo)
        Else
          asign = "-"
          bvalue = (-1) * CDbl(aposo)
          
        End If
        
        cAmount = cAmount + CDbl(asign & "1") * (-1) * CDbl(bvalue)
        bamount = GetPassbookAmount_(cAmount)

      
        formatposo = "": intpart = "": decpart = ""
        formatposo = Right("000000000000" & CStr(bvalue), 12)
        intpart = CStr(CDbl(Left(formatposo, 10)))
        decpart = Right(formatposo, 2)
        bvalue = Right(String(10, " ") & "*" & intpart & "," & decpart, 13)
        
         
        
        If asign = "-" Then
            linedata = linedata & String(8, " ") & bvalue & bamount
        Else
            linedata = linedata & " " & bvalue & String(7, " ") & bamount
        End If

        DocLines(cline + 2) = "  " & linedata & "@"
        
         If CInt(cline) = 20 Then
            LastDocLine = 22
            If Trim(DocLines(3)) <> "" Then
                'If Flag02 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
                If aFlag = 2 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
            End If
            
            PageCounter = PageCounter + 1
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
        Else
            LastDocLine = cline + 2
            cline = cline + 1
        End If
        
         
    Next
    
    'τυπωσε τελευταια σελιδα
    
    adate = format(cPOSTDATE, "DDMMYY")
    Flag02 = True

    If Flag01 And LastFlag = 1 Then
        'EUROAmount = owner.EUROAmount(cAmount)
        EUROAmount = EUROAmount_(cAmount)
        linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
        DocLines(cline + 2) = "  " & linedata
        DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
        LastDocLine = cline + 3
        'cAmount = Round(cAmount / owner.EURORate)
        cAmount = Round(cAmount / EURORate)
        PageCounter = PageCounter + 1
        If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
        If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
        clearDoc_
        cline = 1
        bamount = GetPassbookLargeAmount_(cAmount)
        linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
        DocLines(cline + 2) = "  " & linedata
        cline = 2
        LastFlag = 2
    End If

   Dim aTerminal   As String
   aTerminal = cTERMINALID
   
   If inTrnType = 0 Or inTrnType = 3 Then 'ενημερωση,εξοφληση
        If cline = 1 Then
             If fromLine > 0 Then
                 Dim lcldata As String
                 bamount = GetPassbookLargeAmount_(cAmount)
                 lcldata = String(34, " ") & "EK MET          " & Right(bamount, 13)
                 DocLines(cline + 2) = " " & lcldata
                 cline = 2
             Else
                 fromLine = 1
             End If
        End If
        
        If inTrnType = 0 Then 'ενημερωση
            bamount = GetPassbookAmount_(cAmount)
            linedata = " " & adate & String(1, " ") & aTerminal & "  ΕΝΗΜΕΡΩΘΗΚΕ" & String(23, " ") & bamount & "@"
        ElseIf inTrnType = 3 Then 'εξοφληση
            linedata = " " & adate & String(1, " ") & aTerminal & " " & "ΕΞΟΦΛΗΘΗΚΕ"
        
        End If
        DocLines(cline + 2) = "  " & linedata
        LastDocLine = cline + 2
   
   End If
   
    If Trim(DocLines(3)) <> "" Then
       DocLines(2) = String(59, " ") & "ΕΥΡΩ"
    End If
    PageCounter = PageCounter + 1
    If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
    Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
    
  
End If
    
'    If Not (gPanel Is Nothing) Then
'        gPanel.UnlockPrinter
'    End If
    Set gPanel = Nothing
    
    L2PrintPassbookVersion3_ = cAmount
End Function

Public Function L2PrintPassbookVersion4_(owner As Form, inTrnType As Integer, fromLine As Integer, fromAmount As Double, inDocument As IXMLDOMElement, branch As String, account As String) As String

Dim cline As Integer, PageCounter As Integer
Dim cAmount As Double
Dim aFlag As Integer, Flag01 As Boolean, Flag02 As Boolean, LastFlag As Integer  '01: εγγραφές 2001, 02: εγγραφές 2002


 'aflag = Mid(astr, 21, 1) '1:ΔΡΧ 2:ΕΥΡΩ
 
cline = fromLine
cAmount = fromAmount

clearDoc_
LastFlag = 0
Dim linedata As String
Dim adate As String, aterm As String, abranch As String, aaccount As String, aCode As String, aposo As String, asign As String
Dim bvalue As String, bamount As String, EUROAmount As String
Dim Row As IXMLDOMElement
Dim formatposo As String, decpart As String, intpart As String
   
Dim maxrows As String
Dim EURORate As Double
EURORate = EURORate_

Dim Cur As String

If gPanel Is Nothing Then
    Set gPanel = New GlobalSPCPanel
End If

maxrows = inDocument.SelectNodes("//ROWS").length
If (maxrows > 0) Then
    For Each Row In inDocument.SelectNodes("//ROWS")
        If cline = 0 Then
            cline = 1
            fromLine = 1
        ElseIf cline = 1 Then
            If fromLine > 0 Then
                bamount = GetPassbookLargeAmount_(cAmount)
                linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
                DocLines(cline + 2) = "  " & linedata
                cline = 2
            Else
                fromLine = 1
            End If
        
        End If
        
        If (Row.selectSingleNode("CURRENCY").Text = "") Then  'not empty row
            Exit For
        End If
        Cur = Row.selectSingleNode("CURRENCY").Text
        If Cur = "GRD" Then aFlag = 1
        If Cur = "EUR" Then aFlag = 2
        'aflag = Row.selectSingleNode("DCUR").Text '1:ΔΡΧ 2:ΕΥΡΩ
        
        Flag01 = Flag01 Or (aFlag = "1")
        Flag02 = Flag02 Or (aFlag = "2")
        If Flag01 And Flag02 And LastFlag = 1 Then

            EUROAmount = EUROAmount_(cAmount)
            linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
            DocLines(cline + 2) = "  " & linedata
            DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
            LastDocLine = cline + 3
            
         
            'cAmount = Round(cAmount / owner.EURORate)
            cAmount = Round(cAmount / EURORate)
            PageCounter = PageCounter + 1
            
            If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
            
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
            bamount = GetPassbookLargeAmount_(cAmount)
            linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
            DocLines(cline + 2) = "  " & linedata
            cline = 2
        
        End If
        
        If aFlag = "1" Then LastFlag = 1
        If aFlag = "2" Then LastFlag = 2
        
        
        adate = Row.selectSingleNode("TRDATE").Text
        If Len(adate) = 10 Then adate = Left(adate, 2) & Mid(adate, 4, 2) & Right(adate, 2)
        If Len(adate) <> 6 Then adate = Left(adate & String(6, "0"), 6)
        
        
        aterm = Right(Row.selectSingleNode("BRANCH_SND").Text & Row.selectSingleNode("ATERM_ID").Text, 5)
        aterm = Left(aterm & "     ", 5)
        'abranch = Row.selectSingleNode("DBRANCH").Text
        'aaccount = abranch & Row.selectSingleNode("DACCOUNT").Text
        abranch = branch
        aaccount = account
        
        aCode = Right("000" & Row.selectSingleNode("REASON_CODE").Text, 3)
        aposo = Row.selectSingleNode("ENT_AMNT").Text
        
        'linedata = " " & adate & " " & aterm & " " & aaccount & " " & aCode
        linedata = " " & adate & " " & aterm & " " & aaccount & " " & aCode
        
        If CDbl(aposo) > 0 Then
          asign = "+"
          bvalue = CDbl(aposo)
        Else
          asign = "-"
          bvalue = (-1) * CDbl(aposo)
        End If
        
        cAmount = cAmount + CDbl(asign & "1") * CDbl(bvalue)
        bamount = GetPassbookAmount_(cAmount)

      
        formatposo = "": intpart = "": decpart = ""
        formatposo = Right("000000000000" & CStr(bvalue), 12)
        intpart = CStr(CDbl(Left(formatposo, 10)))
        decpart = Right(formatposo, 2)
        bvalue = Right(String(10, " ") & "*" & intpart & "," & decpart, 13)
        
        If asign = "-" Then
            linedata = linedata & " " & bvalue & String(7, " ") & bamount
        Else
            linedata = linedata & String(8, " ") & bvalue & bamount
        End If

        DocLines(cline + 2) = "  " & linedata & "@"
        
         If CInt(cline) = 20 Then
            LastDocLine = 22
            If Trim(DocLines(3)) <> "" Then
                'If Flag02 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
                If aFlag = 2 Then DocLines(2) = String(59, " ") & "ΕΥΡΩ" Else DocLines(2) = String(59, " ") & "ΔΡΧ"
            End If
            
            PageCounter = PageCounter + 1
            If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
            Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
            clearDoc_
            cline = 1
        Else
            LastDocLine = cline + 2
            cline = cline + 1
        End If
        
         
    Next
    
    'τυπωσε τελευταια σελιδα
    
    adate = format(cPOSTDATE, "DDMMYY")
    Flag02 = True

    If Flag01 And LastFlag = 1 Then
        'EUROAmount = owner.EUROAmount(cAmount)
        EUROAmount = EUROAmount_(cAmount)
        linedata = " " & "      ΙΣΟΤΙΜΟ ΥΠΟΛ. ΣΕ ΕΥΡΩ: " & EUROAmount
        DocLines(cline + 2) = "  " & linedata
        DocLines(cline + 3) = "  ΜΕΤΑΤΡΟΠΗ ΥΠΟΛΟΙΠΟΥ ΣΕ ΕΥΡΩ ΜΕ ΙΣΟΤΙΜΙΑ 1 ΕΥΡΩ=340,75 ΔΡΑΧΜΕΣ"
        LastDocLine = cline + 3
        'cAmount = Round(cAmount / owner.EURORate)
        cAmount = Round(cAmount / EURORate)
        PageCounter = PageCounter + 1
        If Trim(DocLines(3)) <> "" Then DocLines(2) = String(59, " ") & "ΔΡΧ"
        If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
        clearDoc_
        cline = 1
        bamount = GetPassbookLargeAmount_(cAmount)
        linedata = String(34, " ") & "EK MET         " & Right(bamount, 13)
        DocLines(cline + 2) = "  " & linedata
        cline = 2
        LastFlag = 2
    End If

   Dim aTerminal   As String
   aTerminal = cTERMINALID
   
   If inTrnType = 0 Or inTrnType = 3 Then 'ενημερωση,εξοφληση
        If cline = 1 Then
             If fromLine > 0 Then
                 Dim lcldata As String
                 bamount = GetPassbookLargeAmount_(cAmount)
                 lcldata = String(34, " ") & "EK MET          " & Right(bamount, 13)
                 DocLines(cline + 2) = " " & lcldata
                 cline = 2
             Else
                 fromLine = 1
             End If
        End If
        
        If inTrnType = 0 Then 'ενημερωση
            bamount = GetPassbookAmount_(cAmount)
            linedata = " " & adate & String(1, " ") & aTerminal & "  ΕΝΗΜΕΡΩΘΗΚΕ" & String(23, " ") & bamount & "@"
        ElseIf inTrnType = 3 Then 'εξοφληση
            linedata = " " & adate & String(1, " ") & aTerminal & " " & "ΕΞΟΦΛΗΘΗΚΕ"
        
        End If
        DocLines(cline + 2) = "  " & linedata
        LastDocLine = cline + 2
   
   End If
   
    If Trim(DocLines(3)) <> "" Then
       DocLines(2) = String(59, " ") & "ΕΥΡΩ"
    End If
    PageCounter = PageCounter + 1
    If PageCounter = 1 Then L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου" _
    Else L2PrintDocLines_ owner, "Αλλαγή Σελίδας Βιβλιαρίου"
    
  
End If
    
'    If Not (gPanel Is Nothing) Then
'        gPanel.UnlockPrinter
'    End If
    Set gPanel = Nothing
    
    L2PrintPassbookVersion4_ = cAmount
End Function


Public Function L2PrintPassbook6_(owner As Form, inAccount As String, inTrnType As Integer, _
    inTrnCode As String, inTrnAmount As Double, inTrnDRXAmount As Double, _
    fromLine As Integer, fromAmount As Double, inDocument As IXMLDOMElement, _
    failedtrnflag As Boolean, Optional inTrnEuroFinalAmount As Double) As String
    
' inTrnType 0:Ενημέρωση 1: Καταθεση 2: Ανάληψη 3: Εξόφληση

'εκτύπωση βιβλιαρίου πρόγραμμα 5

Dim i As Integer, k As Integer, astr As String
Dim linedata As String, cline As Integer, cAmount As Double
Dim bamount As String, eamount As String
Dim aValue As String, bvalue As String
Dim adate As String, asign As String, aTerminal As String, aaccount As String, aCode As String

cline = fromLine
cAmount = fromAmount

clearDoc_
linedata = ""
    
Dim maxrows As String
Dim Row As IXMLDOMElement
Dim rowcounter As Integer
rowcounter = 0

Dim print_type As String, trans_dt As String, value_dt As String, term_dep_exp_dt As String, rsn_cd As String
Dim ent_amnt As String, aent_amnt As Double, aent_amntsign As String
Dim psbk_ball As String, apsbk_ball As Double, apsbk_ballsign As String
Dim trans_amnt As String, atrans_amnt As Double, atrans_amntsign As String
Dim int_rate As String, aint_rate As Double
Dim cur_iso As String, send_br As String, term_id As String, star_id As String
Dim parite As String, aparite As Double

If gPanel Is Nothing Then
    Set gPanel = New GlobalSPCPanel
End If

maxrows = inDocument.SelectNodes("//ROWS").length
If (maxrows > 0) Then
    For rowcounter = 0 To maxrows
        
        If cline = 0 Then
            cline = 1
        ElseIf cline = 1 Then
            If fromLine > 0 Then
                bamount = GetPassbookLargeAmount_(cAmount)
                linedata = String(34, " ") & "EK MET " & bamount
                DocLines(cline + 2) = " " & linedata
                cline = 2
            Else
                fromLine = 1
            End If
        End If
            
        If rowcounter < maxrows Then
            Set Row = inDocument.SelectNodes("//ROWS")(rowcounter)
            
            'print_type = Row.selectSingleNode("PRINT_TYPE").Text
            trans_dt = Row.selectSingleNode("TRDATE").Text & String(10, " ")
            'value_dt = Row.selectSingleNode("VALUE_DT").Text & String(10, " ")
            'term_dep_exp_dt = Row.selectSingleNode("TERM_DEP_EXP_DT").Text & String(10, " ")
            rsn_cd = Right("000" & Row.selectSingleNode("REASON_CODE").Text, 3)
            ent_amnt = Row.selectSingleNode("ENT_AMNT").Text
           ' psbk_ball = Row.selectSingleNode("PSBK_BALANCE").Text
            cur_iso = Row.selectSingleNode("CURRENCY_TRANS").Text
            trans_amnt = Row.selectSingleNode("TRANS_AMOUNT").Text
            send_br = Row.selectSingleNode("BRANCH_SND").Text & String(3, " ")
            'term_id = Row.selectSingleNode("ATERM_ID").Text & String(2, " ")
            term_id = Row.selectSingleNode("ATERM_ID").Text
            'int_rate = Row.selectSingleNode("INT_RATE").Text
            'parite = Row.selectSingleNode("PARITE").Text
            'star_id = Row.selectSingleNode("STAR_IND").Text
            
            trans_dt = Mid(trans_dt, 1, 2) & Mid(trans_dt, 4, 2) & Mid(trans_dt, 9, 2)
           ' value_dt = Mid(value_dt, 1, 2) & Mid(value_dt, 4, 2) & Mid(value_dt, 9, 2)
            'term_dep_exp_dt = Mid(term_dep_exp_dt, 1, 2) & Mid(term_dep_exp_dt, 4, 2) & Mid(term_dep_exp_dt, 9, 2)
            send_br = Mid(send_br, 1, 3)
            'term_id = Mid(term_id, 1, 2)
            
            atrans_amnt = CDbl(trans_amnt)
            atrans_amntsign = " "
            If atrans_amnt < 0 Then
                atrans_amntsign = "-"
            End If
            aent_amnt = CDbl(ent_amnt)
            aent_amntsign = " "
            If aent_amnt > 0 Then
                aent_amntsign = "+"
            Else
                aent_amntsign = "-"
            End If
              
            If CDbl(ent_amnt) > 0 Then
              asign = "+"
              bvalue = CDbl(ent_amnt)
            Else
              asign = "-"
              bvalue = (-1) * CDbl(ent_amnt)
            End If
            
            cAmount = cAmount + CDbl(asign & "1") * CDbl(bvalue)
            bamount = GetPassbookAmount_(cAmount)
          
            linedata = gFormat_("%6ST% %5ST% %3ST% %16SR% %3ST% %17SR%%1ST% %18SR%", _
                Array(trans_dt, term_id, cur_iso, GetStrAmount_(Abs(atrans_amnt), 15, 2), rsn_cd, GetStrAmount_(Abs(aent_amnt), 15, 2), aent_amntsign, bamount))
                
            DocLines(cline + 2) = " " & linedata
            LastDocLine = cline + 2
            If CInt(cline) = 20 Then
                LastDocLine = 22
                L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
                clearDoc_
                cline = 1
            Else
                LastDocLine = cline + 2
                cline = cline + 1
            End If
        
        Else
            If Not failedtrnflag Then
                If inTrnType = 0 Then
                    adate = format(cPOSTDATE, "DDMMYY")
                    aTerminal = StrPad_(cTERMINALID, 5, " ", "L")
    
                    linedata = adate & aTerminal & " " & " ΕΝΗΜΕΡΩΘΗΚΕ *****"
    
                    DocLines(cline + 2) = " " & linedata
                    LastDocLine = cline + 2
                End If
            End If
            
            L2PrintDocLines_ owner, "Εισαγωγή Βιβλιαρίου"
            clearDoc_
     End If

        Next
    End If

    Set gPanel = Nothing

    L2PrintPassbook6_ = cAmount
End Function

Public Function CalculateIRISTime(IRISTime As String) As String
    On Error GoTo ErrorPos
    If Trim(IRISTime) <> "" Then
        Dim ti As Long
        Dim h, m, s As Integer
        Dim tTime
        ti = CLng(IRISTime) / 1000
        h = Int(ti / 3600)
        m = Int((ti Mod 3600) / 60)
        s = Int((ti Mod 3600) Mod 60)
        tTime = TimeSerial(h, m, s)
        CalculateIRISTime = tTime
    Else
        CalculateIRISTime = IRISTime
    End If
    Exit Function
ErrorPos:
    

End Function

