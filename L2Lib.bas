Attribute VB_Name = "L2Lib"
Option Explicit

Public L2TrnListFile As MSXML2.DOMDocument30
Public L2ModelFile As MSXML2.DOMDocument30
Public ActiveL2TrnHandler As L2TrnHandler


Public Const bvFalse = "0"
Public Const bvTrue = "-1"

Public Function L2ChkTaxID(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim aTaxid As String, res As Boolean
    If Not (inDocument.selectSingleNode("//TaxID") Is Nothing) Then
        aTaxid = inDocument.selectSingleNode("//TaxID").Text
    End If
    res = TRNFrm.ChkTaxID(aTaxid)
    If Not res Then
        L2ChkTaxID = "<MESSAGE><ERROR><LINE>ΛΑΝΘΑΣΜΕΝΟ ΑΦΜ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkTaxID = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkTaxID = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΤΟΥ ΑΦΜ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

'Function TrnFileFromTrnCode(handler As L2TrnHandler, inCode As String)
'    Dim atrnnode As IXMLDOMElement
'    Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "']")
'    If atrnnode Is Nothing Then
'        L2TrnListFile.Load ReadDir & "\" & "L2TrnList.xml"
'        Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "']")
'    End If
'    If Not (atrnnode Is Nothing) Then
'        If atrnnode.Attributes.getNamedItem("filename") Is Nothing Then
'        Else
'            TrnFileFromTrnCode = ReadDir & "\" & atrnnode.getAttribute("filename") & ".xml"
'            Dim child As IXMLDOMNode
'            For Each child In atrnnode.childNodes
'                If child.nodeType = NODE_ELEMENT Then
'                    If Not child.Attributes.getNamedItem("name") Is Nothing Then
'                        Dim inDoc As MSXML2.DOMDocument30
'                        Set inDoc = New MSXML2.DOMDocument30
'                        inDoc.LoadXML child.XML
'                        handler.addFormUpdate inDoc, child.Attributes.getNamedItem("name").Text
'                    End If
'                End If
'            Next child
'        End If
'    End If
'    Set atrnnode = Nothing
'End Function

Function TrnNodeFromTrnCode(inCode As String) As IXMLDOMElement
    Dim atrnnode As IXMLDOMElement
    If L2TrnListFile Is Nothing Then
        L2TrnListFile.Load ReadDir & "\" & "L2TrnList.xml"
    End If
    If cKMODEFlag = True And cKMODEValue <> "" Then
        Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "' and @kmode='" & cKMODEValue & "']")
    End If
    If atrnnode Is Nothing Then
        Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "']")
    End If
    Set TrnNodeFromTrnCode = atrnnode
End Function

Function TrnFileFromTrnNode(handler As L2TrnHandler, atrnnode As IXMLDOMElement)
    If Not (atrnnode Is Nothing) Then
        If atrnnode.Attributes.getNamedItem("filename") Is Nothing Then
        Else
            TrnFileFromTrnNode = ReadDir & "\" & atrnnode.getAttribute("filename") & ".xml"
            Dim child As IXMLDOMNode
            For Each child In atrnnode.childNodes
                If child.nodeType = NODE_ELEMENT Then
                    If Not child.Attributes.getNamedItem("name") Is Nothing Then
                        Dim inDoc As MSXML2.DOMDocument30
                        Set inDoc = New MSXML2.DOMDocument30
                        inDoc.LoadXML child.XML
                        handler.addFormUpdate inDoc, child.Attributes.getNamedItem("name").Text
                    End If
                End If
            Next child
        End If
    End If
End Function

'Function HiddenFlagFromTrnCode(inCode As String)
'    HiddenFlagFromTrnCode = False
'    Dim atrnnode As IXMLDOMElement
'    Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "']")
'    If atrnnode Is Nothing Then
'        L2TrnListFile.Load ReadDir & "\" & "L2TrnList.xml"
'    End If
'    Set atrnnode = L2TrnListFile.documentElement.selectSingleNode("//trn[@id='" & inCode & "']")
'    If Not (atrnnode Is Nothing) Then
'        Dim hiddenattr As IXMLDOMAttribute
'        Set hiddenattr = atrnnode.getAttributeNode("hidden")
'        If Not (hiddenattr Is Nothing) Then
'            If hiddenattr.value = "-1" Then HiddenFlagFromTrnCode = True
'        End If
'    End If
'End Function
Function HiddenFlagFromTrnNode(atrnnode As IXMLDOMElement)
    HiddenFlagFromTrnNode = False
    If Not (atrnnode Is Nothing) Then
        Dim hiddenattr As IXMLDOMAttribute
        Set hiddenattr = atrnnode.getAttributeNode("hidden")
        If Not (hiddenattr Is Nothing) Then
            If hiddenattr.Value = "-1" Then
                HiddenFlagFromTrnNode = True
            End If
        End If
    End If
End Function


Public Function L2ChkFldType(ByRef Message As String, invalue As String, inValidation As String)
'01: ΧΩΡΙΣ VALIDATION
'04: Δάνειο με CD,
'05: Δάνειο χωρίς CD,
'07: Αριθμός Εγγραφής
'08: Ειδικός με CD
'09: Γενικός Λογαριασμός Δανείου
'10: Λογαριασμός Καταθέσεων με 1 CD
'11: Τραπεζική Επιταγή
'12: ΕΘΝΟΚΑΡΤΑ
'13: Τραπεζική Εντολή

Dim astr As String, ares As Integer, aFlag As Boolean
Dim sec As Integer, min As Integer, hour As Integer
On Error GoTo chkFailed
    astr = invalue
    L2ChkFldType = True
    If Trim(astr) = "" Then Exit Function
    
    Select Case UCase(Replace(inValidation, "ς", "Σ"))
    Case "DATE"
        If Len(astr) <= 6 Then
            astr = StrPad_(astr, 6, "0", "L")
            astr = Left(astr, 2) & "/" & Mid(astr, 3, 2) & "/" & Right(astr, 2)
        ElseIf Len(astr) > 6 And Len(astr) <= 8 Then
            astr = StrPad_(astr, 8, "0", "L")
            If Mid(astr, 3, 1) <> "/" Then
                astr = Left(astr, 2) & "/" & Mid(astr, 3, 2) & "/" & Right(astr, 4)
            End If
        End If
        
        If Not IsDate(astr) Then
            Message = "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
            Message = "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
            Message = "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
    Case "TIME06"
        Message = "Λάθος Μορφή Ώρας"
        If Len(astr) > 6 Then GoTo chkFailed
        If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
        If Not IsNumeric(astr) Then GoTo chkFailed
        If CLng(astr) < 0 Then GoTo chkFailed
        sec = CInt(Right(astr, 2))
        min = CInt(Mid(astr, 3, 2))
        hour = CInt(Left(astr, 2))
        If sec < 0 Or sec > 59 Then GoTo chkFailed
        If min < 0 Or min > 59 Then GoTo chkFailed
        If hour < 0 Or hour > 23 Then GoTo chkFailed
    Case "TIME04"
        Message = "Λάθος Μορφή Ώρας"
        If Len(astr) > 4 Then GoTo chkFailed
        If Len(astr) < 4 Then astr = StrPad_(astr, 4, "0", "L")
        If Not IsNumeric(astr) Then GoTo chkFailed
        If CLng(astr) < 0 Then GoTo chkFailed
        min = CInt(Right(astr, 2))
        hour = CInt(Left(astr, 2))
        If min < 0 Or min > 59 Then GoTo chkFailed
        If hour < 0 Or hour > 23 Then GoTo chkFailed
    Case "ACCOUNT2CD"
        If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
        If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 11, "0", "L")

        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If
        
        ares = CalcCd2_(Left(astr, 10))
        If CInt(Mid(astr, 11, 1)) <> ares Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "ACCOUNT1CD"
        If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
        If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 10, "0", "L")
        
        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If
        
        ares = CalcCd2_(Left(astr, 10))
    Case "ACCOUNTIRISCD"
       If Len(astr) <= 10 Then
            astr = StrPad_(astr, 10, "0", "L")
            ares = CalcCd1_(Left(astr, 9), 9)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
       Else
            astr = StrPad_(astr, 11, "0", "L")
            ares = CalcCd2_(Left(astr, 10))
            If CInt(Mid(astr, 11, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
       End If
       
    Case "ETEBANKCHECK" 'Τραπεζική Επιταγή
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "ΙΔΙΩΤΙΚΉ ΕΠΙΤΑΓΉ"
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "ETEPRIVATECHECK" 'ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "13" 'Τραπεζική Εντολή
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
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
        
    Case "4"
        If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
        ares = CalcCd1_(Left(astr, 9), 9)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "7"
        If Len(astr) < 13 Then astr = StrPad_(astr, 13, "0", "L")
        ares = CalcCd1_(Left(astr, 12), 12)
        If CInt(Right(astr, 1)) <> ares Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "9"
        astr = Trim(astr)
        If Len(astr) > 3 Then
            ares = CalcSAccCd(Left(astr, Len(astr) - 1))
            If CInt(Right(astr, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        End If
    Case "12"
        astr = Trim(astr)
        If Not ChkCard(astr) Then
            Message = "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case "21" 'Λογαριασμος Καταθέσεων Γερμανία
        If Len(astr) < 8 Then 'astr = StrPad_(astr, 8, "0", "L")
            astr = StrPad_(astr, 7, "0", "L")
            ares = CalcCd1_(Mid(astr, 1, 6), 6)
            If CInt(Mid(astr, 7, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
        Else
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Message = "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
            ares = CalcCd2_(Left(astr, 10))
            If CInt(Mid(astr, 11, 1)) <> ares Then
                Message = "Υποχρεωτικό πεδίο"
                GoTo chkFailed
            End If
        End If
    Case "NOMISMA_IRIS"
        If CInt(astr) < 1 Or CInt(astr) > 96 Then
            Message = "Λάθος Νόμισμα"
            GoTo chkFailed
        End If
    Case "ΛΟΓΑΡΙΑΣΜΌΣ ΕΘΝΟΚΑΡΤΑΣ"
        If Not ChkCard(astr) Then
            Message = "Λάθος Ψηφίο Ελέχγου"
            GoTo chkFailed
        End If
    Case "ΑΜ ΣΥΝΤΑΞΙΟΥΧΟΥ 4ΧΧΧ"
        If Not ChkPensionCD(astr) Then
            Message = "Λάθος Αριθμός Μητρώου Συνταξιούχου"
            GoTo chkFailed
        End If
    End Select
    
    L2ChkFldType = True
    
    Exit Function
chkFailed:
    L2ChkFldType = False

End Function

Public Function L2ChkBank(inDocument As IXMLDOMElement) As String
'έλεγχος κωδικού τράπεζας
    On Error GoTo ErrorPos
    Dim abank As String, abranch As String, abankname As String
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        abank = inDocument.selectSingleNode("//bank").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        abranch = inDocument.selectSingleNode("//branch").Text
    End If
    abankname = GetBankName_(CLng(abank))
    If abankname = "" Then
        L2ChkBank = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΤΡΑΠΕΖΑΣ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkBank = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkBank = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΤΡΑΠΕΖΑΣ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function


Public Function L2ChkBankAccount(inDocument As IXMLDOMElement) As String
'έλεγχος check digit λογαριασμού τράπεζας
    On Error GoTo ErrorPos
    Dim abank As String, abranch As String, aaccount As String
    Dim ChequeType As Integer
    Dim found As Boolean
    found = False
    
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        abank = inDocument.selectSingleNode("//bank").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        abranch = inDocument.selectSingleNode("//branch").Text
    End If
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        aaccount = inDocument.selectSingleNode("//account").Text
    End If
    If Not (inDocument.selectSingleNode("//type") Is Nothing) Then
        If Trim(inDocument.selectSingleNode("//type").Text) <> "" Then
            ChequeType = CInt(inDocument.selectSingleNode("//type").Text)
            found = True
        End If
    End If
    
    Dim res As Boolean
    If found Then
        res = ChkBankAccount_(abank, abranch, aaccount, ChequeType)
    Else
        res = ChkBankAccount_(abank, abranch, aaccount)
    End If
    If Not res Then
        L2ChkBankAccount = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkBankAccount = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkBankAccount = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2ChkBankCheque(inDocument As IXMLDOMElement) As String
'έλεγχος check digit επιταγής τράπεζας
    On Error GoTo ErrorPos
    Dim abank As String, abranch As String, aaccount As String, acheque As String
    Dim ChequeType As Integer
    Dim found As Boolean
    found = False
    
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        abank = inDocument.selectSingleNode("//bank").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        abranch = inDocument.selectSingleNode("//branch").Text
    End If
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        aaccount = inDocument.selectSingleNode("//account").Text
    End If
    If Not (inDocument.selectSingleNode("//cheque") Is Nothing) Then
        acheque = inDocument.selectSingleNode("//cheque").Text
    End If
    If Not (inDocument.selectSingleNode("//type") Is Nothing) Then
        If Trim(inDocument.selectSingleNode("//type").Text) <> "" Then
            ChequeType = CInt(inDocument.selectSingleNode("//type").Text)
            found = True
        End If
    End If
    
    Dim res As Boolean
    If found Then
        res = ChkBankCheque_(abank, abranch, aaccount, acheque, ChequeType)
    Else
        res = ChkBankCheque_(abank, abranch, aaccount, acheque)
    End If
    If Not res Then
        res = ChkGenBankCheque_(acheque)
    End If
    If Not res Then
        L2ChkBankCheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΕΠΙΤΑΓΗΣ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkBankCheque = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkBankCheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΕΠΙΤΑΓΗΣ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2ChkETECheque(inDocument As IXMLDOMElement) As String
'έλεγχος check digit επιταγής ετε
    On Error GoTo ErrorPos
    Dim abank As String, abranch As String, aaccount As String, acheque As String
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        abank = inDocument.selectSingleNode("//bank").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        abranch = inDocument.selectSingleNode("//branch").Text
    End If
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        aaccount = inDocument.selectSingleNode("//account").Text
    End If
    If Not (inDocument.selectSingleNode("//cheque") Is Nothing) Then
        acheque = inDocument.selectSingleNode("//cheque").Text
    End If
    Dim res As Boolean
    res = ChkETECheque_(CLng(acheque))
    If Not res Then
        L2ChkETECheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΕΠΙΤΑΓΗΣ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkETECheque = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkETECheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΕΠΙΤΑΓΗΣ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function
Public Function L2CDETECheque(inDocument As IXMLDOMElement) As String
'υπολογισμός check digit επιταγής ετε
    On Error GoTo ErrorPos
    Dim acheque As String
    If Not (inDocument.selectSingleNode("//cheque") Is Nothing) Then
        acheque = inDocument.selectSingleNode("//cheque").Text
    End If
    
    L2CDETECheque = "<MESSAGE></MESSAGE>"
    
    Dim res As String
    Dim i As Integer
    For i = 0 To 9
        res = acheque & CStr(i)
        If TRNFrm.ChkETECheque(res) Then
            L2CDETECheque = "<MESSAGE><ChequeNoCD>" + acheque + "</ChequeNoCD><CD>" + CStr(i) + "</CD><ChequeWithCD>" + res + "</ChequeWithCD></MESSAGE>"
            Exit For
        End If
    Next
    
    Exit Function
ErrorPos:
    L2CDETECheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΕΠΙΤΑΓΗΣ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function


Public Function L2ChkETEAccount(inDocument As IXMLDOMElement) As String
'έλεγχος check digit λογαριασμου ετε
    On Error GoTo ErrorPos
    Dim abank As String, abranch As String, aaccount As String, acheque As String
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        abank = inDocument.selectSingleNode("//bank").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        abranch = inDocument.selectSingleNode("//branch").Text
    End If
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        aaccount = inDocument.selectSingleNode("//account").Text
    End If
    Dim ares1 As Integer, ares2 As Integer, res As Boolean
    aaccount = Right("00000000000" & aaccount, 11)
    ares1 = CalcCd1_(Mid(aaccount, 4, 6), 6)
    ares2 = CalcCd2_(Left(aaccount, 10))
    If CInt(Mid(aaccount, 10, 1)) <> ares1 Or CInt(Mid(aaccount, 11, 1)) <> ares2 Then
        L2ChkETEAccount = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkETEAccount = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkETEAccount = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΛΟΓΑΡΙΑΣΜΟΥ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2ChiefKey(inDocument As IXMLDOMElement) As String
'αίτηση για κλειδι Chief Teller
    On Error GoTo ErrorPos
    ManagerRequest = False
    Dim aChiefUserName As String
    
    Dim DoNotAcceptCurrentUserAsAuthUser As Boolean
    DoNotAcceptCurrentUserAsAuthUser = False
    If Not inDocument Is Nothing Then
        If Not inDocument.selectSingleNode("./DoNotAcceptCurrentUserAsAuthUser") Is Nothing Then
            If UCase(inDocument.selectSingleNode("./DoNotAcceptCurrentUserAsAuthUser").Text) = "TRUE" Then
                DoNotAcceptCurrentUserAsAuthUser = True
            End If
        End If
    End If
    
    If isChiefTeller And DoNotAcceptCurrentUserAsAuthUser = False Then
        KeyAccepted = False: ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = Screen.activeform
        KeyWarning.Show vbModal, Screen.activeform
    Else
        KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: ChiefRequest = True
        SelKeyFrm.Show vbModal, Screen.activeform
    End If
    If Not KeyAccepted Then
        L2ChiefKey = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΚΛΕΙΔΙ CHIEF TELLER</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        Exit Function
    Else
        If isChiefTeller And DoNotAcceptCurrentUserAsAuthUser = False Then
            aChiefUserName = cUserName
        Else
            aChiefUserName = cCHIEFUserName
        End If
        
        UpdateChiefKey aChiefUserName
        L2ChiefKey = "<MESSAGE><CHIEF>" & aChiefUserName & "</CHIEF></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChiefKey = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΚΛΕΙΔΙ CHIEF TELLER" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2ManagerKey(inDocument As IXMLDOMElement) As String
'αίτηση για κλειδι Manager Teller
    On Error GoTo ErrorPos
    ChiefRequest = False
    Dim aManagerUserName As String
    
    Dim DoNotAcceptCurrentUserAsAuthUser As Boolean
    DoNotAcceptCurrentUserAsAuthUser = False
    If Not inDocument Is Nothing Then
        If Not inDocument.selectSingleNode("./DoNotAcceptCurrentUserAsAuthUser") Is Nothing Then
            If UCase(inDocument.selectSingleNode("./DoNotAcceptCurrentUserAsAuthUser").Text) = "TRUE" Then
                DoNotAcceptCurrentUserAsAuthUser = True
            End If
        End If
    End If
    
    If isManager And DoNotAcceptCurrentUserAsAuthUser = False Then
        KeyAccepted = False: ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = Screen.activeform
        KeyWarning.Show vbModal, Screen.activeform
    Else
        KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: ManagerRequest = True
        SelKeyFrm.Show vbModal, Screen.activeform
    End If
    If Not KeyAccepted Then
        L2ManagerKey = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΚΛΕΙΔΙ MANAGER</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        Exit Function
    Else
        If isManager And DoNotAcceptCurrentUserAsAuthUser = False Then
            aManagerUserName = cUserName
        Else
            aManagerUserName = cMANAGERUserName
        End If
        
        UpdateManagerKey aManagerUserName
        L2ManagerKey = "<MESSAGE><MANAGER>" & aManagerUserName & "</MANAGER></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ManagerKey = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΚΛΕΙΔΙ MANAGER" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2AnyKey(inDocument As IXMLDOMElement) As String
'αίτηση για κλειδι Chief ή Manager Teller
    Dim aManagerUserName As String
    Dim aChiefUserName As String
    ChiefRequest = False
    ManagerRequest = False
    AnyRequest = False
    cANYKEY = ""
    SaveJournal
    On Error GoTo ErrorPos
    If isManager Then
        KeyAccepted = False: ManagerRequest = True: Load KeyWarning: Set KeyWarning.owner = Screen.activeform
        KeyWarning.Show vbModal, Screen.activeform
        cANYKEY = "MANAGER"
    ElseIf isChiefTeller Then
        KeyAccepted = False: ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = Screen.activeform
        KeyWarning.Show vbModal, Screen.activeform
        cANYKEY = "CHIEF"
    Else
        KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: AnyRequest = True
        SelKeyFrm.Show vbModal, Screen.activeform
    End If
    If Not KeyAccepted Then
        L2AnyKey = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        AnyRequest = False
        Exit Function
    Else
        If cANYKEY = "MANAGER" Then
            ChiefRequest = False
            ManagerRequest = True
            If Not (isManager) Then
                aManagerUserName = cMANAGERUserName
            Else
                aManagerUserName = cUserName
            End If
            UpdateManagerKey aManagerUserName
        ElseIf cANYKEY = "CHIEF" Then
            ChiefRequest = True
            ManagerRequest = False
            If Not (isChiefTeller) Then
                aChiefUserName = cCHIEFUserName
            Else
                aChiefUserName = cUserName
            End If
            UpdateChiefKey aChiefUserName
        End If
        L2AnyKey = "<MESSAGE><CHIEF>" & aChiefUserName & "</CHIEF><MANAGER>" & aManagerUserName & "</MANAGER><KEY>" & cANYKEY & "</KEY></MESSAGE>"
        AnyRequest = False
    End If
    Exit Function
ErrorPos:
    L2AnyKey = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    AnyRequest = False
End Function

Public Function L2AnyKeyUnconditional(inDocument As IXMLDOMElement) As String
'αίτηση για κλειδι Chief Teller ή Manager χωρίς έλεγχο του χρήστη
    Dim aUserName As String
    cCHIEFUserName = ""
    cMANAGERUserName = ""
    cANYKEY = ""
    SaveJournal
    On Error GoTo ErrorPos
    
    KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: AnyRequest = True
    SelKeyFrm.Show vbModal, Screen.activeform
    
    If Not KeyAccepted Then
        L2AnyKeyUnconditional = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        AnyRequest = False
        Exit Function
    Else
        If cANYKEY = "MANAGER" Then
            ChiefRequest = False
            ManagerRequest = True
            aUserName = cMANAGERUserName
            UpdateManagerKey aUserName
        ElseIf cANYKEY = "CHIEF" Then
            ChiefRequest = True
            ManagerRequest = False
            aUserName = cCHIEFUserName
            UpdateChiefKey aUserName
        End If
        
        L2AnyKeyUnconditional = "<MESSAGE><AUTHUSER>" & aUserName & "</AUTHUSER><KEY>" & cANYKEY & "</KEY></MESSAGE>"
        AnyRequest = False
    End If
    Exit Function
ErrorPos:
    L2AnyKeyUnconditional = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    AnyRequest = False
End Function
Public Function L24EyesKey(info As String) As String
'αίτηση για κλειδι Chief Teller ή Manager χωρίς έλεγχο του χρήστη
    Dim aUserName As String
    cCHIEFUserName = ""
    cMANAGERUserName = ""
    cANYKEY = ""
    SaveJournal
    On Error GoTo ErrorPos
    
    KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: AnyRequest = True:
    SelKeyFrm.SetReasonText info
    SelKeyFrm.Show vbModal, Screen.activeform
    
    If Not KeyAccepted Then
        L24EyesKey = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        AnyRequest = False
        Exit Function
    Else
        If cANYKEY = "MANAGER" Then
            ChiefRequest = False
            ManagerRequest = True
            aUserName = cMANAGERUserName
            UpdateManagerKey aUserName
        ElseIf cANYKEY = "CHIEF" Then
            ChiefRequest = True
            ManagerRequest = False
            aUserName = cCHIEFUserName
            UpdateChiefKey aUserName
        End If
        
        L24EyesKey = "<MESSAGE><AUTHUSER>" & aUserName & "</AUTHUSER><KEY>" & cANYKEY & "</KEY></MESSAGE>"
        AnyRequest = False
    End If
    Exit Function
ErrorPos:
    L24EyesKey = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓΚΡΙΣΗ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    AnyRequest = False
End Function


Public Function L2IRISAuth(inDocument As IXMLDOMElement) As String
'αίτηση για έγκριση
    On Error GoTo ErrorPos
    KeyAccepted = False
    IRISSelKeyFrm.Show vbModal, Screen.activeform
    If Not KeyAccepted Then
        L2IRISAuth = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓKΡΙΣΗ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        Exit Function
    Else
        L2IRISAuth = "<MESSAGE><IRISAUTH>" & cIRISAuthUserName & "</IRISAUTH></MESSAGE>"
    End If

    Exit Function
ErrorPos:
    L2IRISAuth = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓKΡΙΣΗ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function
Public Function L2IRISAuthLevel(inDocument As IXMLDOMElement) As String
'αίτηση για έγκριση
    On Error GoTo ErrorPos
    Dim sourcenode As IXMLDOMElement
    Set sourcenode = inDocument.selectSingleNode("//input/level")
    KeyAccepted = False
    If Not (sourcenode Is Nothing) Then
        IRISSelKeyFrm.levelAuth = sourcenode.Text
    End If
    IRISSelKeyFrm.Show vbModal, Screen.activeform
    If Not KeyAccepted Then
        L2IRISAuthLevel = "<MESSAGE><ERROR><LINE>ΑΠΟΡΡΙΦΘΗΚΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓKΡΙΣΗ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
        Exit Function
    Else
        L2IRISAuthLevel = "<MESSAGE><IRISAUTH>" & cIRISAuthUserName & "</IRISAUTH></MESSAGE>"
    End If

    Exit Function
ErrorPos:
    L2IRISAuthLevel = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΑΙΤΗΣΗ ΓΙΑ ΕΓKΡΙΣΗ" & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2FTFilaName(inDocument As IXMLDOMElement) As String
    
    Dim documentnode As IXMLDOMElement
    Set documentnode = inDocument.selectSingleNode("//input/document")
    Dim sourcenode As IXMLDOMElement
    Dim destinationnode As IXMLDOMElement
    Set sourcenode = inDocument.selectSingleNode("//input/source")
    Set destinationnode = inDocument.selectSingleNode("//input/destination")
    
    Dim sourcelist As IXMLDOMNodeList
    Dim destinationlist As IXMLDOMNodeList
    If Not (documentnode Is Nothing Or sourcenode Is Nothing Or destinationnode Is Nothing) Then
        Set sourcelist = documentnode.SelectNodes(sourcenode.Text)
        Set destinationlist = documentnode.SelectNodes(destinationnode.Text)
        
        If Not (sourcelist Is Nothing Or destinationlist Is Nothing) Then
            If (sourcelist.length > 0 And sourcelist.length = destinationlist.length) Then
                Dim tablenode As IXMLDOMElement
                Dim applnode As IXMLDOMElement
                Set tablenode = inDocument.selectSingleNode("//input/table")
                Set applnode = inDocument.selectSingleNode("//input/application")
                
                If Not (tablenode Is Nothing Or applnode Is Nothing) Then
                    Dim i As Integer
                    For i = 0 To sourcelist.length - 1
                        destinationlist(i).Text = GetFTFilaName_(tablenode.Text, applnode.Text, sourcelist(i).Text)
                    Next i
                Else
                
                
                End If
            
            Else
            
            
            End If
        
        Else
        
        
        End If
    Else
    
    End If
    
    
    L2FTFilaName = documentnode.XML
    
End Function

Public Function L2FTFilaRecordset(inDocument As IXMLDOMElement) As String
    
    Dim filternode As IXMLDOMElement
    Dim sortnode As IXMLDOMElement
    Set filternode = inDocument.selectSingleNode("//input/filter")
    Set sortnode = inDocument.selectSingleNode("//input/sort")
    
    Dim inFilter As String, inSort As String
    If Not filternode Is Nothing Then inFilter = filternode.Text
    If Not sortnode Is Nothing Then inSort = sortnode.Text
        
    Dim aRecs As ADODB.Recordset
    Dim a_Recs As RecordsetEntry
    
    Set aRecs = AddFTFilaRecordset_("FTFilaRS", inFilter, inSort)
    Set a_Recs = AppRSEntryByName_("FTFilaRS")
    If aRecs.RecordCount > 0 Then
        Dim resultdoc As New MSXML2.DOMDocument30
        resultdoc.appendChild resultdoc.createElement("RECORDSET")
        aRecs.MoveFirst
        While Not aRecs.Eof
            Dim Row As IXMLDOMElement
            Set Row = resultdoc.createElement("row")
            resultdoc.documentElement.appendChild Row
            Row.Attributes.setNamedItem resultdoc.createAttribute("clave_fila")
            Row.Attributes.setNamedItem resultdoc.createAttribute("descr_corta")
            Row.Attributes.setNamedItem resultdoc.createAttribute("descr_larga")
            Row.getAttributeNode("clave_fila").Value = Trim(a_Recs.NVLString("CLAVE_FILA"))
            Row.getAttributeNode("descr_corta").Value = Trim(a_Recs.NVLString("DESCR_CORTA"))
            Row.getAttributeNode("descr_larga").Value = Trim(a_Recs.NVLString("DESCR_LARGA"))
            aRecs.MoveNext
        Wend
        L2FTFilaRecordset = resultdoc.XML
        Exit Function
    End If
    
    L2FTFilaRecordset = "<row/>"
End Function
Public Function L2GetSingleFTFilaName(inDocument As IXMLDOMElement) As String
    Dim tablenode As IXMLDOMElement
    Dim applnode As IXMLDOMElement
     Dim clavenode As IXMLDOMElement
    Set tablenode = inDocument.selectSingleNode("//input/table")
    Set applnode = inDocument.selectSingleNode("//input/application")
    Set clavenode = inDocument.selectSingleNode("//input/clave_fila")
    Dim res As String
    If Not (tablenode Is Nothing Or applnode Is Nothing Or clavenode Is Nothing) Then
        res = GetFTFilaName_(tablenode.Text, applnode.Text, clavenode.Text)
    End If
    L2GetSingleFTFilaName = "<MESSAGE><table>" & tablenode.Text & "</table>" & "<application>" & applnode.Text & "</application>" & "<clave_fila>" & clavenode.Text & "</clave_fila>" & "<descr_corta>" & res & "</descr_corta>" & "</MESSAGE>"
End Function

Public Function L2GetIRISErrorData(inDocument As IXMLDOMElement) As String
    Dim returnstr As String
    returnstr = "<MESSAGE>"
    If Not inDocument Is Nothing Then
      Dim i As Integer
      If inDocument.childNodes.length > 0 Then
         Dim iKeys() As String
         ReDim iKeys(inDocument.childNodes.length - 1)
         For i = 0 To inDocument.childNodes.length - 1
            iKeys(i) = inDocument.childNodes(i).Text
         Next i
         Dim ReturnArray() As String
         ReDim ReturnArray(inDocument.childNodes.length - 1)
         ReturnArray = GetIRISErrorData_(iKeys)
         For i = LBound(ReturnArray) To UBound(ReturnArray)
             returnstr = returnstr & "<DATA>" + ReturnArray(i) + "</DATA>"
         Next i
      End If
    End If
    returnstr = returnstr & "</MESSAGE>"
    L2GetIRISErrorData = returnstr
End Function

Public Function L2Recordset(inDocument As IXMLDOMElement) As String
    Dim cursortypenode As IXMLDOMElement
    Dim locktypenode As IXMLDOMElement
    Dim connectionnode As IXMLDOMElement
    Dim querynode As IXMLDOMElement
    
    Set cursortypenode = inDocument.selectSingleNode("//input/cursortype")
    Set locktypenode = inDocument.selectSingleNode("//input/locktype")
    Set connectionnode = inDocument.selectSingleNode("//input/connection")
    Set querynode = inDocument.selectSingleNode("//input/query")
    
    Dim ars As New ADODB.Recordset, aRSRow As RecordsetEntry
    Dim aConnection As ADODB.Connection
    Dim acursortype, alocktype
    
    Set aConnection = ado_DB
    
    If Not connectionnode Is Nothing Then
        Set aConnection = New ADODB.Connection
        aConnection.ConnectionString = connectionnode.Text
        aConnection.open
    End If
        
    acursortype = adOpenForwardOnly
    alocktype = adLockReadOnly
    If Not locktypenode Is Nothing Then alocktype = locktypenode.Text
    If Not cursortypenode Is Nothing Then acursortype = cursortypenode.Text
    
    Dim aquery As String
    If Not querynode Is Nothing Then aquery = querynode.Text

    If aquery <> "" Then
        ars.open aquery, aConnection, acursortype, alocktype
        
        Dim resultdoc As New MSXML2.DOMDocument30
        resultdoc.appendChild resultdoc.createElement("RECORDSET")
        
        If Not ars.Eof Then
            ars.MoveFirst
            While Not ars.Eof
                Dim Row As IXMLDOMElement
                Set Row = resultdoc.createElement("row")
                resultdoc.documentElement.appendChild Row
                Dim i As Integer
                For i = 0 To ars.fields.Count - 1
                    Row.Attributes.setNamedItem resultdoc.createAttribute(LCase(ars.fields(i).name))
                    Row.getAttributeNode(LCase(ars.fields(i).name)).Value = NVLString_(ars.fields(i).Value, "")
                Next i
                ars.MoveNext
            Wend
        End If
        L2Recordset = resultdoc.XML
        Exit Function
    End If
    L2Recordset = ""
End Function

Public Function L2GetAmountText2002(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim aAmount As String, alinelength As Integer, res As String
    If Not (inDocument.selectSingleNode("//amount") Is Nothing) Then
        aAmount = inDocument.selectSingleNode("//amount").Text
    End If
    If Not (inDocument.selectSingleNode("//linelength") Is Nothing) Then
        alinelength = CInt(inDocument.selectSingleNode("//linelength").Text)
    Else
        alinelength = 0
    End If
    res = TRNFrm.GetAmountText2002(aAmount)
    If Len(res) <= alinelength Or alinelength = 0 Then
        L2GetAmountText2002 = "<MESSAGE><AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE></MESSAGE>"
    Else
        Dim xmlstr As String
        While (Len(res) > alinelength)
            xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & Mid(res, 1, alinelength) & "</AMOUNTTEXTLINE>"
            res = Mid(res, alinelength + 1)
        Wend
        xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE>"
        L2GetAmountText2002 = "<MESSAGE>" & xmlstr & "</MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2GetAmountText2002 = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΜΕΤΑΤΡΟΠΗ ΤΟΥ ΠΟΣΟΥ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2GetAmountText(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim aAmount As String, alinelength As Integer, res As String
    Dim aFlag1, aFlag2 As String
    Dim pFlag1, pFlag2
    pFlag1 = False
    pFlag2 = False
    If Not (inDocument.selectSingleNode("//amount") Is Nothing) Then
        aAmount = inDocument.selectSingleNode("//amount").Text
    End If
    If Not (inDocument.selectSingleNode("//linelength") Is Nothing) Then
        alinelength = CInt(inDocument.selectSingleNode("//linelength").Text)
    Else
        alinelength = 0
    End If
    If Not (inDocument.selectSingleNode("//lang") Is Nothing) Then
       If UCase(inDocument.selectSingleNode("//lang").Text) = "TRUE" Then
          pFlag1 = True
       End If
    End If
    If Not (inDocument.selectSingleNode("//curr") Is Nothing) Then
       If UCase(inDocument.selectSingleNode("//curr").Text) = "TRUE" Then
          pFlag2 = True
       End If
    End If
    res = TRNFrm.GetAmountTextGen(aAmount)
    If Len(res) <= alinelength Or alinelength = 0 Then
        L2GetAmountText = "<MESSAGE><AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE></MESSAGE>"
    Else
        Dim xmlstr As String
        While (Len(res) > alinelength)
            xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & Mid(res, 1, alinelength) & "</AMOUNTTEXTLINE>"
            res = Mid(res, alinelength + 1)
        Wend
        xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE>"
        L2GetAmountText = "<MESSAGE>" & xmlstr & "</MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2GetAmountText = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΜΕΤΑΤΡΟΠΗ ΤΟΥ ΠΟΣΟΥ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function
Public Function L2PrintPassbookLine(inDocument As IXMLDOMElement) As String
'(inAccount, inTrnDate, inTrnCode, inTrnAmount1, _
'    inTrnAmount2, fromLine, fromAmount)
'inTrnAmount1: ποσο καταθεσης
'inTrnAmount2: ποσο αναληψης
    Dim account As String
    Dim inTrnDate As String
    Dim inTrnCode As String
    Dim depositamount As String
    Dim withdrawamount As String
    Dim fromLine As String
    Dim fromAmount As String
    Dim PrintPromptMessage As String
    Dim inTerm As String
    
    On Error GoTo ErrorPos
    
    If Not (inDocument.selectSingleNode("//printpromptmessage") Is Nothing) Then
        PrintPromptMessage = inDocument.selectSingleNode("//printpromptmessage").Text
    End If
    
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        account = inDocument.selectSingleNode("//account").Text
    End If
    If Not (inDocument.selectSingleNode("//interm") Is Nothing) Then
        inTerm = inDocument.selectSingleNode("//interm").Text
    End If
    If Not (inDocument.selectSingleNode("//intrndate") Is Nothing) Then
        inTrnDate = inDocument.selectSingleNode("//intrndate").Text
    End If
    
    If Not (inDocument.selectSingleNode("//intrncode") Is Nothing) Then
        inTrnCode = inDocument.selectSingleNode("//intrncode").Text
    Else
        inTrnCode = "0"
    End If
    
    If Not (inDocument.selectSingleNode("//depositamount") Is Nothing) Then
        depositamount = inDocument.selectSingleNode("//depositamount").Text
    End If
    
    If Not (inDocument.selectSingleNode("//withdrawamount") Is Nothing) Then
        withdrawamount = inDocument.selectSingleNode("//withdrawamount").Text
    End If
    
    If Not (inDocument.selectSingleNode("//fromline") Is Nothing) Then
        fromLine = inDocument.selectSingleNode("//fromline").Text
    End If
    
    If Not (inDocument.selectSingleNode("//fromamount") Is Nothing) Then
        fromAmount = inDocument.selectSingleNode("//fromamount").Text
    End If
    
    If gPanel Is Nothing Then Set gPanel = New GlobalSPCPanel

    Dim aform As TRNFrm
    Set aform = New TRNFrm

    aform.PrintPromptMessage = PrintPromptMessage '"Εισαγωγή Βιβλιαρίου"

    PrintSinglePassbookLine_ aform, CStr(account), CStr(inTrnDate), CInt(inTrnCode), CDbl(depositamount), _
        CDbl(withdrawamount), CInt(fromLine), CDbl(fromAmount), CStr(inTerm)

     Set gPanel = Nothing
     Unload aform
     Set aform = Nothing
     
     
    L2PrintPassbookLine = "<MESSAGE></MESSAGE>"
    Exit Function
ErrorPos:
    L2PrintPassbookLine = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΚΤΥΠΩΣΗ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2PrintPassbook(inDocument As IXMLDOMElement) As String
'Dim inAccount As String, inTrnType As Integer, inTrnCode As String, inTrnAmount As Double
'Dim fromLine As Integer, fromAmount As Double
  Dim inTrnType As Integer, fromLine As Integer
  Dim fromAmount As String
 
  
  If Not (inDocument.selectSingleNode("//intrntype") Is Nothing) Then
        inTrnType = inDocument.selectSingleNode("//intrntype").Text
    Else
        inTrnType = "0"
  End If
    
  If Not (inDocument.selectSingleNode("//fromline") Is Nothing) Then
        fromLine = inDocument.selectSingleNode("//fromline").Text
    Else
        fromLine = "0"
  End If
  If Not (inDocument.selectSingleNode("//fromamount") Is Nothing) Then
        fromAmount = inDocument.selectSingleNode("//fromamount").Text
    End If
    
  Dim res As String
  res = L2PrintPassbook_(ActiveL2TrnHandler.activeform, inTrnType, fromLine, CDbl(fromAmount), inDocument)
  
  L2PrintPassbook = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
  Exit Function

ErrorPos:
    L2PrintPassbook = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΚΤΥΠΩΣΗ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2Show1041Messages(inDocument As IXMLDOMElement) As String

Dim inLineNo As Integer
Dim inLineData As String

On Error GoTo ErrorPos

    Dim resultdocument As New MSXML2.DOMDocument30
    resultdocument.LoadXML inDocument.XML
        
    Load XMLMessageForm
    Set XMLMessageForm.MessageDocument = resultdocument
    XMLMessageForm.Show vbModal
    

    L2Show1041Messages = "<MESSAGE>F12</MESSAGE>" ' για να διαβιβασει ξανα
    Exit Function
ErrorPos:
    L2Show1041Messages = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΚΤΥΠΩΣΗ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2ShowDepositMessages(inDocument As IXMLDOMElement) As String

Dim inLineNo As Integer
Dim inLineData As String

On Error GoTo ErrorPos

    Dim resultdocument As New MSXML2.DOMDocument30
    resultdocument.LoadXML inDocument.XML
        
'    Load XMLMessageForm
'    Set XMLMessageForm.MessageDocument = resultdocument
'    XMLMessageForm.Show vbModal
    
    
    'Load DepositMessageForm
    Set DepositMessageForm.MessageDocument = resultdocument
    DepositMessageForm.Show vbModal, ActiveL2TrnHandler.activeform

    L2ShowDepositMessages = resultdocument.XML    '"<MESSAGE>F12</MESSAGE>" ' για να διαβιβασει ξανα
    Exit Function
ErrorPos:
    L2ShowDepositMessages = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΚΤΥΠΩΣΗ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function
Public Function L2ShowDepositMassiveMessages(inDocument As IXMLDOMElement) As String

Dim inLineNo As Integer
Dim inLineData As String

On Error GoTo ErrorPos

    Dim resultdocument As New MSXML2.DOMDocument30
    resultdocument.LoadXML inDocument.XML
        
'    Load XMLMessageForm
'    Set XMLMessageForm.MessageDocument = resultdocument
'    XMLMessageForm.Show vbModal
    
    
    'Load DepositMessageForm
    Set DepositMassiveMessageForm.MessageDocument = resultdocument
    DepositMassiveMessageForm.Show vbModal, ActiveL2TrnHandler.activeform

    L2ShowDepositMassiveMessages = resultdocument.XML    '"<MESSAGE>F12</MESSAGE>" ' για να διαβιβασει ξανα
    Exit Function
ErrorPos:
    L2ShowDepositMassiveMessages = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΚΤΥΠΩΣΗ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2SCreateIBAN(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    'αν ειδος λογ/μου προγραμμα 1 ή 2 τοτε λογ/μος = 0000066860168393
    'αν ειδος λογ/μου προγραμμα 5 τοτε λογ/μος = 1000066860168393
    Dim branch As String '0668
    Dim account As String '0000066860168393 ή 1000066860168393
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        branch = inDocument.selectSingleNode("//branch").Text
    End If
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        account = inDocument.selectSingleNode("//account").Text
    End If
    Dim IBAN As String
    IBAN = CreateIBAN_(branch, account)
    L2SCreateIBAN = "<MESSAGE><IBAN>" & IBAN & "</IBAN></MESSAGE>"
 
    Exit Function
ErrorPos:
    L2SCreateIBAN = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΔΗΜΙΟΥΡΓΙΑ IBAN ΛΟΓΑΡΙΑΣΜΟΥ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function
Public Function L2FormatIBAN(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim IBAN As String, fiban As String
    
    If Not (inDocument.selectSingleNode("//iban") Is Nothing) Then
        IBAN = inDocument.selectSingleNode("//iban").Text
    End If
    fiban = FormatIBAN_(IBAN)
    L2FormatIBAN = "<MESSAGE><FORMATEDIBAN>" & fiban & "</FORMATEDIBAN></MESSAGE>"
 
    Exit Function
ErrorPos:
    L2FormatIBAN = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΜΟΡΦΟΠΟΙΗΣΗ IBAN ΛΟΓΑΡΙΑΣΜΟΥ " & Err.number & " " & Err.description & "</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
End Function

Public Function L2GetChequeAmountText(inDocument As IXMLDOMElement) As String
   On Error GoTo ErrorPos
    Dim aAmount As String, alinelength As Integer, res As String
    If Not (inDocument.selectSingleNode("//amount") Is Nothing) Then
        aAmount = inDocument.selectSingleNode("//amount").Text
    End If
    If Not (inDocument.selectSingleNode("//linelength") Is Nothing) Then
        alinelength = CInt(inDocument.selectSingleNode("//linelength").Text)
    Else
        alinelength = 0
    End If
    res = TRNFrm.GetChequeAmountText(aAmount)
    If Len(res) <= alinelength Or alinelength = 0 Then
        L2GetChequeAmountText = "<MESSAGE><AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE></MESSAGE>"
    Else
        Dim xmlstr As String
        While (Len(res) > alinelength)
            xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & Mid(res, 1, alinelength) & "</AMOUNTTEXTLINE>"
            res = Mid(res, alinelength + 1)
        Wend
        xmlstr = xmlstr & "<AMOUNTTEXTLINE>" & res & "</AMOUNTTEXTLINE>"
        L2GetChequeAmountText = "<MESSAGE>" & xmlstr & "</MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2GetChequeAmountText = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΜΕΤΑΤΡΟΠΗ ΠΟΣΟΥ " & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2ReadOCR() As String
    Dim res As String
    Dim clearRes As String
    Dim aaccount As String
    Dim abranch As String
    Dim abank As String
    Dim aAm As String
    Dim aCheck As String
    Dim aPos As Integer
    Dim B As String
    Dim aIban As String
    Dim chkType As String
    res = TRNFrm.ReadOCR
    clearRes = Replace(Replace(Replace(res, "<", "&lt;"), ">", "&gt;"), Chr(0), "")
    
    If Len(res) >= 53 Then
     If Left(res, 10) = "RES:000*;+" Then
        aPos = InStr(1, res, "<")
        If aPos > 0 And Len(res) > aPos + 28 Then
           aaccount = Mid(res, aPos, 28): B = Left(aaccount, 12): aaccount = Right(aaccount, 16)
           abranch = Right(B, 4): B = Left(B, 8)
           abank = Right(B, 3)
           aIban = "GR" & Mid(res, aPos + 3, 25)
           If aPos > 1 Then
            chkType = Mid(res, aPos - 1, 1)
           End If
           res = Right(res, Len(res) - aPos - 28)
        End If
        
        aPos = InStr(1, res, "<")
        If aPos > 0 And Len(res) > aPos + 10 Then
           aCheck = Mid(res, aPos, 10): aCheck = Right(aCheck, 9)
           res = Right(res, Len(res) - aPos - 10)
        End If

        If Len(res) > 2 Then
           If Left(res, 1) = ">" Then
              res = Right(res, Len(res) - 1)
              aPos = InStr(1, res, ">")
              If aPos > 1 Then
                 aAm = Left(res, aPos - 1)
              End If
           End If
        End If
     End If
    End If
    L2ReadOCR = "<MESSAGE>" & "<OCR>" & clearRes & "</OCR>" & "<ACCOUNT>" & aaccount & "</ACCOUNT>" & _
                "<BRANCH>" & abranch & "</BRANCH>" & "<BANK>" & abank & "</BANK>" & _
                "<CHECK>" & aCheck & "</CHECK>" & "<AM>" & aAm & "</AM>" & "<IBAN>" & aIban & "</IBAN>" & "<TYPE>" & chkType & "</TYPE>" & "</MESSAGE>"
                
    Exit Function
End Function

Public Function L2ChkETEBankCheque(inDocument As IXMLDOMElement) As String
   On Error GoTo ErrorPos
   Dim acheque As String
   If Not (inDocument.selectSingleNode("//cheque") Is Nothing) Then
        acheque = inDocument.selectSingleNode("//cheque").Text
   End If
   Dim res As Boolean
   Dim ares As Integer
   res = False
   If acheque = "" Or acheque = "0" Then
      res = True
   Else
     If CLng(acheque) >= 550000000 And CLng(acheque) <= 600000000 Then
        res = ChkGenBankCheque_(acheque)
     Else
       ''FBB  1706/2013
        'res = ChkGenBankCheque_(acheque)
        
       If Len(acheque) > 1 Then
          ares = CLng(Left(acheque, Len(acheque) - 1)) Mod 11
          If ares = 10 Then ares = 0
          If CInt(Right(acheque, 1)) = ares Then res = True
       Else
          res = True
       End If
     End If
   End If
   If Not res Then
        L2ChkETEBankCheque = "<MESSAGE><ERROR><LINE>ΛΑΘΟΣ ΑΡΙΘΜΟΣ ΕΝΤΟΛΗΣ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkETEBankCheque = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2ChkETEBankCheque = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΟΣ ΕΝΤΟΛΗΣ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2ChkCard(inDocument As IXMLDOMElement)
    On Error GoTo ErrorPos
    Dim cardaccount As String
    Dim res As Boolean
    If Not (inDocument.selectSingleNode("//cardaccount") Is Nothing) Then
        cardaccount = inDocument.selectSingleNode("//cardaccount").Text
    End If
    If cardaccount = "" Or cardaccount = "0" Then
        res = True
    Else
        res = ChkCard(cardaccount)
    End If
    If Not res Then
        L2ChkCard = "<MESSAGE><ERROR><LINE>ΛΑΘΟΣ ΛΟΓΑΡΙΑΣΜΟΣ ΕΘΝΟΚΑΡΤΑΣ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    Else
        L2ChkCard = "<MESSAGE></MESSAGE>"
    End If
    Exit Function
ErrorPos: L2ChkCard = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΟΣ ΛΟΓΑΡΙΑΣΜΟΥ ΕΘΝΟΚΑΡΤΑΣ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"""
End Function


Public Function L2ChkValidIBAN(inDocument As IXMLDOMElement) As String
   On Error GoTo ErrorPos
   Dim IBAN As String
   If Not (inDocument.selectSingleNode("//iban") Is Nothing) Then
        IBAN = inDocument.selectSingleNode("//iban").Text
   End If
   
   L2ChkValidIBAN = "<MESSAGE></MESSAGE>"
   
   If IBAN <> "" Then
      Dim res As Integer
      res = TRNFrm.ChkValidIBAN(IBAN)
   
      If res = 0 Then
            L2ChkValidIBAN = "<MESSAGE></MESSAGE>"
      ElseIf res = 1 Then
            L2ChkValidIBAN = "<MESSAGE><ERROR><LINE>ΜΗ ΑΠΟΔΕΚΤΟΣ IBAN</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
      Else
            L2ChkValidIBAN = "<MESSAGE><ERROR><LINE>Ο ΛΟΓΑΡΙΑΣΜΟΣ ΔΕΝ ΕΙΝΑΙ IBAN</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
      End If
   End If
   
   Exit Function
ErrorPos:
    L2ChkValidIBAN = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΟΣ IBAN" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2ChkSATNo(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim satnumber As String
    Dim asum As Long
    Dim CD As Integer
    Dim res As Boolean
    Dim i As Integer
    If Not (inDocument.selectSingleNode("//satnumber") Is Nothing) Then
        satnumber = inDocument.selectSingleNode("//satnumber").Text
   End If
   If satnumber = "" Then
      res = True
   Else
      If Len(satnumber) < 11 Then
        res = False
      Else
        asum = 0
        For i = 1 To 10
            If i Mod 2 = 0 Then
               asum = asum + CInt(Mid(satnumber, i, 1))
            Else
               asum = asum + CInt(Mid(satnumber, i, 1)) * 2
            If CInt(Mid(satnumber, i, 1)) >= 5 Then asum = asum - 9
            End If
       
        Next
        CD = 10 - (asum Mod 10)
        If CD = 10 Then CD = 0
        res = (CD = CInt(Mid(satnumber, 11, 1)))
      End If
   End If
   If Not res Then
       L2ChkSATNo = "<MESSAGE><ERROR><LINE>ΛΑΘΟΣ ΨΗΦΙΟ ΕΛΕΓΧΟΥ</LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
   Else
       L2ChkSATNo = "<MESSAGE></MESSAGE>"
   End If
   Exit Function
ErrorPos:
L2ChkSATNo = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2AreaBuilder(inDocument As IXMLDOMElement) As String
Dim area As String
Dim ComArea As cXmlComArea

'    <comarea name="CALL_P49E4" id="P49E4" filename="@P49E4">
'        <method name="P49E4" trncall="49E4" inputname="IDATA" outputname="ODATA"/>
'    </comarea>


    If Not (inDocument.selectSingleNode("//comarea") Is Nothing) Then
         Set ComArea = New cXmlComArea
         Set ComArea.content = inDocument.selectSingleNode("//comarea")
         
         'διορθωση cXmlComArea.Container
         'Set ComArea.owner = ActiveL2TrnHandler.DocumentManager
         Set ComArea.Container = ActiveL2TrnHandler.DocumentManager.TrnBuffers
         
        L2AreaBuilder = "<MESSAGE>" + ComArea.BuilderParseArea(ComArea.content) + "</MESSAGE>"
         Set ComArea = Nothing
    
    End If
    
Exit Function
ErrorPos:
    L2AreaBuilder = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΔΗΜΙΟΥΡΓΙΑ ΤΗΣ COM AREA" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"

End Function

Public Function L2SendXmlBuffer(inDocument As IXMLDOMElement) As String
'Dim newbuffer As buffer
Dim ComArea As cXmlComArea
Dim aname As String
Dim content As IXMLDOMElement
Dim owner As Object
    If Not (inDocument.selectSingleNode("//comarea") Is Nothing) Then
        Set ComArea = New cXmlComArea
        Set ComArea.content = inDocument.selectSingleNode("//comarea")
        
        Set ComArea.Container = ActiveL2TrnHandler.DocumentManager.TrnBuffers
    End If
    
    If Not (inDocument.selectSingleNode("//buffer") Is Nothing) Then
        If Not (inDocument.selectSingleNode("//buffer/*").nodename = "") Then
            aname = inDocument.selectSingleNode("//buffer/*").nodename
            Set content = inDocument.selectSingleNode("//buffer/" & aname)
        End If
        Dim aresult As String
        aresult = ComArea.ParseCallWithID(content, aname, "ODATA")
        ComArea.ComResult.UpdateXmlDocumentManager ActiveL2TrnHandler.DocumentManager
        L2SendXmlBuffer = "<MESSAGE>" + aresult + "</MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2SendXmlBuffer = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΔΗΜΙΟΥΡΓΙΑ ΤΗΣ COM AREA" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"

End Function

Public Function L2MessageBox(inDocument As IXMLDOMElement) As String
    Dim title_elm As IXMLDOMElement
    Dim Message As String
    
    Set title_elm = inDocument.selectSingleNode(".//title")
    Dim Title As String
    If Not title_elm Is Nothing Then Title = title_elm.Text
    
    Dim Line As IXMLDOMElement
    
    For Each Line In inDocument.SelectNodes(".//line")
        If Message <> "" Then Message = Message & vbCrLf
        Message = Message & Line.Text
    Next Line
    
    Dim Buttons As VbMsgBoxStyle
    Dim Button As IXMLDOMElement
    For Each Button In inDocument.SelectNodes(".//button")
        If Not Button.Attributes.getNamedItem("type") Is Nothing Then
            If UCase(Button.Attributes.getNamedItem("type").Text) = UCase("AbortRetryIgnore") Then
                Buttons = Buttons + vbAbortRetryIgnore
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("ApplicationModal") Then
                Buttons = Buttons + vbApplicationModal
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("Critical") Then
                Buttons = Buttons + vbCritical
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("DefaultButton1") Then
                Buttons = Buttons + vbDefaultButton1
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("DefaultButton2") Then
                Buttons = Buttons + vbDefaultButton2
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("DefaultButton3") Then
                Buttons = Buttons + vbDefaultButton3
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("DefaultButton4") Then
                Buttons = Buttons + vbDefaultButton4
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("Exclamation") Then
                Buttons = Buttons + vbExclamation
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("Information") Then
                Buttons = Buttons + vbInformation
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("MsgBoxHelpButton") Then
                Buttons = Buttons + vbMsgBoxHelpButton
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("MsgBoxRight") Then
                Buttons = Buttons + vbMsgBoxRight
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("MsgBoxRtlReading") Then
                Buttons = Buttons + vbMsgBoxRtlReading
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("MsgBoxSetForeground") Then
                Buttons = Buttons + vbMsgBoxSetForeground
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("OkCancel") Then
                Buttons = Buttons + vbOKCancel
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("OkOnly") Then
                Buttons = Buttons + vbOKOnly
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("Question") Then
                Buttons = Buttons + vbQuestion
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("RetryCancel") Then
                Buttons = Buttons + vbRetryCancel
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("SystemModal") Then
                Buttons = Buttons + vbSystemModal
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("YesNo") Then
                Buttons = Buttons + vbYesNo
            ElseIf UCase(Button.Attributes.getNamedItem("type").Text) = UCase("YesNoCancel") Then
                Buttons = Buttons + vbYesNoCancel
            End If
        End If
    Next Button
    
    Dim res As VbMsgBoxResult
    res = MsgBox(Message, Buttons, Title)
    L2MessageBox = "<RESULT>" & CStr(res) & "</RESULT>"
    
End Function

Public Function L2PrintPassbookNewVersion3(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    sbWriteLogFile "s1231.input.xml", inDocument.XML
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim CallerNode As IXMLDOMElement
    Set CallerNode = GetXmlNode(inDocument, "//I_CALLER", "CALLER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim finishNode As IXMLDOMElement
    Set finishNode = GetXmlNode(inDocument, "//I_FINISH", "FINISH", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca1231 As cXmlComArea
    Set ca1231 = DeclareComArea_("CALL_1233", "S1233", "S1233", "1233", "@S1233", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    'διορθωση cXmlComArea.Container
    'Set ca1231.owner = owner.Manager
    Set ca1231.Container = owner.Manager.TrnBuffers
    
    With ca1231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_CALLER") = CallerNode.Text
        .v2Value("IDATA/I_FINISH") = finishNode.Text
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    Do
        Dim response As String
        response = ca1231.ParseCall(Nothing)
        'sbWriteLogFile "s1231.response.xml", response
       
        If response = "" Then
             L2PrintPassbookNewVersion3 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Function
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion3
        Set messager.ComArea = ca1231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        Dim resp As IXMLDOMElement
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
        'If Not ca1231.HandleResp(ResponseDoc, "Λάθος...") Then
            L2PrintPassbookNewVersion3 = response
            Exit Function
        End If
        'sbWriteLogFile "s1231.response.xml", response
        
        
        'Dim allrowscount, notemptyrows As Integer
        'allrowscount = ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS").length
        'notemptyrows = ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[DCUR!='']").length
        Dim rowsfetched As Integer
        rowsfetched = ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("D_TRANS").Text & " " & _
                    rownode.selectSingleNode("CURRENCY").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("TERM_ID").Text & " " & _
                    rownode.selectSingleNode("UNP_AMOUNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca1231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
    Dim res As String
    Dim rowsdoc As New MSXML2.DOMDocument30
    rowsdoc.LoadXML "<root>" & allrows & "</root>"
    
  
   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
   res = L2PrintPassbookVersion3_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 7))
    L2PrintPassbookNewVersion3 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function

Public Function L2PrintPassbookNewVersion4(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    sbWriteLogFile "s1231.input.xml", inDocument.XML
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim CallerNode As IXMLDOMElement
    Set CallerNode = GetXmlNode(inDocument, "//I_CALLER", "CALLER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim finishNode As IXMLDOMElement
    Set finishNode = GetXmlNode(inDocument, "//I_FINISH", "FINISH", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca1231 As cXmlComArea
    Set ca1231 = DeclareComArea_("CALL_1231", "S1231", "S1231", "1231", "@S1231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    'διορθωση cXmlComArea.Container
    'Set ca1231.owner = owner.Manager
    Set ca1231.Container = owner.Manager.TrnBuffers
    
    With ca1231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_CALLER") = CallerNode.Text
        .v2Value("IDATA/I_FINISH") = finishNode.Text
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    Do
        Dim response As String
        response = ca1231.ParseCall(Nothing)
        'sbWriteLogFile "s1231.response.xml", response
       
        If response = "" Then
             L2PrintPassbookNewVersion4 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Function
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion3
        Set messager.ComArea = ca1231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            L2PrintPassbookNewVersion4 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Function
'            response = "<MESSAGE><RC>999</RC><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
'            ResponseDoc.LoadXml response
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
        'If Not ca1231.HandleResp(ResponseDoc, "Λάθος...") Then
            L2PrintPassbookNewVersion4 = response
            Exit Function
        End If
        'sbWriteLogFile "s1231.response.xml", response
        
        
        'Dim allrowscount, notemptyrows As Integer
        'allrowscount = ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS").length
        'notemptyrows = ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[DCUR!='']").length
        Dim rowsfetched As Integer
        rowsfetched = ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("D_TRANS").Text & " " & _
                    rownode.selectSingleNode("CURRENCY").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("TERM_ID").Text & " " & _
                    rownode.selectSingleNode("UNP_AMOUNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca1231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
    Dim res As String
    Dim rowsdoc As New MSXML2.DOMDocument30
    rowsdoc.LoadXML "<root>" & allrows & "</root>"
    
  
   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
   res = L2PrintPassbookVersion3_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 7))
    L2PrintPassbookNewVersion4 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function


Public Function L2PrintPassbookNewVersion6(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim I_IPNode As IXMLDOMElement
    Set I_IPNode = GetXmlNode(inDocument, "//I_IP", "I_IP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim I_ACC_X_TP_ORDERNode As IXMLDOMElement
    Set I_ACC_X_TP_ORDERNode = GetXmlNode(inDocument, "//I_ACC_X_TP_ORDER", "I_ACC_X_TP_ORDER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim T_TIMESTAMPNode As IXMLDOMElement
    Set T_TIMESTAMPNode = GetXmlNode(inDocument, "//T_TIMESTAMP", "T_TIMESTAMP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim OPERATIONNode As IXMLDOMElement
    Set OPERATIONNode = GetXmlNode(inDocument, "//OPERATION", "OPERATION", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim SYSTEMNode As IXMLDOMElement
    Set SYSTEMNode = GetXmlNode(inDocument, "//SYSTEM", "SYSTEM", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca3231 As cXmlComArea
    Set ca3231 = DeclareComArea_("CALL_3231", "S3231", "S3231", "3231", "@S3231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    'διορθωση cXmlComArea.Container
    'Set ca1231.owner = owner.Manager
    Set ca3231.Container = owner.Manager.TrnBuffers
    
    With ca3231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_IP") = I_IPNode.Text
        .v2Value("IDATA/I_ACC_X_TP_ORDER") = I_ACC_X_TP_ORDERNode.Text
        .v2Value("IDATA/T_TIMESTAMP") = T_TIMESTAMPNode.Text
        .v2Value("IDATA/OPERATION") = OPERATIONNode.Text
        .v2Value("IDATA/SYSTEM") = SYSTEMNode.Text
        
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    Do
        Dim response As String
        response = ca3231.ParseCall(Nothing)
       
        If response = "" Then
             L2PrintPassbookNewVersion6 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Function
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion4
        Set messager.ComArea = ca3231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            L2PrintPassbookNewVersion6 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Function
'            response = "<MESSAGE><RC>999</RC><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
'            ResponseDoc.LoadXml response
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
        'If Not ca1231.HandleResp(ResponseDoc, "Λάθος...") Then
            L2PrintPassbookNewVersion6 = response
            Exit Function
        End If
      
        Dim rowsfetched As Integer
        rowsfetched = ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca3231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("TRDATE").Text & " " & _
                    rownode.selectSingleNode("CURRENCY").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("ATERM_ID").Text & " " & _
                    rownode.selectSingleNode("ENT_AMNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca3231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .v2Value("IDATA/T_TIMESTAMP") = .v2Value("ODATA//T_TIMESTAMP")
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
    Dim res As String
    Dim rowsdoc As New MSXML2.DOMDocument30
    rowsdoc.LoadXML "<root>" & allrows & "</root>"
    
  
   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
   res = L2PrintPassbookVersion4_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 10))
    L2PrintPassbookNewVersion6 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function

Public Function L2PrintPassbookNewVersion7(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim I_IPNode As IXMLDOMElement
    Set I_IPNode = GetXmlNode(inDocument, "//I_IP", "I_IP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim I_ACC_X_TP_ORDERNode As IXMLDOMElement
    Set I_ACC_X_TP_ORDERNode = GetXmlNode(inDocument, "//I_ACC_X_TP_ORDER", "I_ACC_X_TP_ORDER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim T_TIMESTAMPNode As IXMLDOMElement
    Set T_TIMESTAMPNode = GetXmlNode(inDocument, "//T_TIMESTAMP", "T_TIMESTAMP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim OPERATIONIDNode As IXMLDOMElement
    Set OPERATIONIDNode = GetXmlNode(inDocument, "//OPERATION_ID", "OPERATION_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim SYSTEMNode As IXMLDOMElement
    Set SYSTEMNode = GetXmlNode(inDocument, "//SYSTEM", "SYSTEM", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca3231 As cXmlComArea
    Set ca3231 = DeclareComArea_("CALL_3231", "S3231", "S3231", "3231", "@S3231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    'διορθωση cXmlComArea.Container
    'Set ca1231.owner = owner.Manager
    Set ca3231.Container = owner.Manager.TrnBuffers
    
    With ca3231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_IP") = I_IPNode.Text
        .v2Value("IDATA/I_ACC_X_TP_ORDER") = I_ACC_X_TP_ORDERNode.Text
        .v2Value("IDATA/T_TIMESTAMP") = T_TIMESTAMPNode.Text
        .v2Value("IDATA/OPERATION_ID") = OPERATIONIDNode.Text
        .v2Value("IDATA/SYSTEM") = SYSTEMNode.Text
        
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    Do
        Dim response As String
        response = ca3231.ParseCall(Nothing)
       
        If response = "" Then
             L2PrintPassbookNewVersion7 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Function
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion4
        Set messager.ComArea = ca3231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            L2PrintPassbookNewVersion7 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Function
'            response = "<MESSAGE><RC>999</RC><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
'            ResponseDoc.LoadXml response
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
        'If Not ca1231.HandleResp(ResponseDoc, "Λάθος...") Then
            L2PrintPassbookNewVersion7 = response
            Exit Function
        End If
      
        Dim rowsfetched As Integer
        rowsfetched = ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca3231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("TRDATE").Text & " " & _
                    rownode.selectSingleNode("CURRENCY").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("ATERM_ID").Text & " " & _
                    rownode.selectSingleNode("ENT_AMNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca3231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .v2Value("IDATA/T_TIMESTAMP") = .v2Value("ODATA//T_TIMESTAMP")
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
    Dim res As String
    Dim rowsdoc As New MSXML2.DOMDocument30
    rowsdoc.LoadXML "<root>" & allrows & "</root>"
    
  
   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
   res = L2PrintPassbookVersion4_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 10))
    L2PrintPassbookNewVersion7 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function


Public Function L2PrintExchangePassbookNew(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim I_IPNode As IXMLDOMElement
    Set I_IPNode = GetXmlNode(inDocument, "//I_IP", "I_IP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim I_ACC_X_TP_ORDERNode As IXMLDOMElement
    Set I_ACC_X_TP_ORDERNode = GetXmlNode(inDocument, "//I_ACC_X_TP_ORDER", "I_ACC_X_TP_ORDER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim T_TIMESTAMPNode As IXMLDOMElement
    Set T_TIMESTAMPNode = GetXmlNode(inDocument, "//T_TIMESTAMP", "T_TIMESTAMP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim OPERATIONNode As IXMLDOMElement
    Set OPERATIONNode = GetXmlNode(inDocument, "//OPERATION", "OPERATION", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim SYSTEMNode As IXMLDOMElement
    Set SYSTEMNode = GetXmlNode(inDocument, "//SYSTEM", "SYSTEM", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    allrows = ""
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca3231 As cXmlComArea
    Set ca3231 = DeclareComArea_("CALL_3231", "S3231", "S3231", "3231", "@S3231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    Set ca3231.Container = owner.Manager.TrnBuffers
    
    With ca3231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_IP") = I_IPNode.Text
        .v2Value("IDATA/I_ACC_X_TP_ORDER") = I_ACC_X_TP_ORDERNode.Text
        .v2Value("IDATA/T_TIMESTAMP") = T_TIMESTAMPNode.Text
        .v2Value("IDATA/OPERATION") = OPERATIONNode.Text
        .v2Value("IDATA/SYSTEM") = SYSTEMNode.Text
        
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    
    Dim failed As Boolean
    Dim err_message5 As String
    failed = False
    err_message5 = ""
    
    Do
        Dim response As String
        response = ca3231.ParseCall(Nothing)
       
        If response = "" Then
             failed = True
             err_message5 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Do
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion4
        Set messager.ComArea = ca3231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            failed = True
            err_message5 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Do
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
            failed = True
            err_message5 = response
            Exit Do
        End If
      
        Dim rowsfetched As Integer
        rowsfetched = ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca3231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("TRDATE").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("ATERM_ID").Text & " " & _
                    rownode.selectSingleNode("CURRENCY_TRANS").Text & " " & _
                    rownode.selectSingleNode("TRANS_AMOUNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text & " " & _
                    rownode.selectSingleNode("ENT_AMNT").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca3231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .v2Value("IDATA/T_TIMESTAMP") = .v2Value("ODATA//T_TIMESTAMP")
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
      If failed = False Or (failed = True And allrows <> "") Then
        Dim res As String
        Dim rowsdoc As New MSXML2.DOMDocument30
        rowsdoc.LoadXML "<root>" & allrows & "</root>"
        
        'res = L2PrintPassbook6_(owner.Manager.activeform, Mid(AccountNode.Text, 1, 10), 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, TRNTypeNode.Text, failed)
        res = L2PrintPassbook6_(owner.Manager.activeform, Mid(AccountNode.Text, 1, 10), 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, failed)
        'res = L2PrintPassbook5_(owner.Manager.activeform, AccountNode.Text, 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, failed)
'    Else
'        L2PrintExchangePassbook = err_message5
'        Exit Function
    End If
    
    If failed = False Then err_message5 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
    L2PrintExchangePassbookNew = err_message5

    
    
'    Dim res As String
'    Dim rowsdoc As New MSXML2.DOMDocument30
'    rowsdoc.LoadXML "<root>" & allrows & "</root>"
'
'
'
'
'   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
'   'res = L2PrintPassbook6_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 10))
'   res = L2PrintPassbook6_(owner.Manager.activeform, Mid(AccountNode.Text, 1, 10), 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, TRNTypeNode.Text)
'    L2PrintExchangePassbookNew = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function

Public Function L2PrintExchangePassbookNew1(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim I_IPNode As IXMLDOMElement
    Set I_IPNode = GetXmlNode(inDocument, "//I_IP", "I_IP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim I_ACC_X_TP_ORDERNode As IXMLDOMElement
    Set I_ACC_X_TP_ORDERNode = GetXmlNode(inDocument, "//I_ACC_X_TP_ORDER", "I_ACC_X_TP_ORDER", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim T_TIMESTAMPNode As IXMLDOMElement
    Set T_TIMESTAMPNode = GetXmlNode(inDocument, "//T_TIMESTAMP", "T_TIMESTAMP", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim OPERATIONIDNode As IXMLDOMElement
    Set OPERATIONIDNode = GetXmlNode(inDocument, "//OPERATION_ID", "OPERATION_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim SYSTEMNode As IXMLDOMElement
    Set SYSTEMNode = GetXmlNode(inDocument, "//SYSTEM", "SYSTEM", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    allrows = ""
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca3231 As cXmlComArea
    Set ca3231 = DeclareComArea_("CALL_3231", "S3231", "S3231", "3231", "@S3231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    Set ca3231.Container = owner.Manager.TrnBuffers
    
    With ca3231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        .v2Value("IDATA/I_IP") = I_IPNode.Text
        .v2Value("IDATA/I_ACC_X_TP_ORDER") = I_ACC_X_TP_ORDERNode.Text
        .v2Value("IDATA/T_TIMESTAMP") = T_TIMESTAMPNode.Text
        .v2Value("IDATA/OPERATION_ID") = OPERATIONIDNode.Text
        .v2Value("IDATA/SYSTEM") = SYSTEMNode.Text
        
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    
    Dim failed As Boolean
    Dim err_message5 As String
    failed = False
    err_message5 = ""
    
    Do
        Dim response As String
        response = ca3231.ParseCall(Nothing)
       
        If response = "" Then
             failed = True
             err_message5 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Do
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion4
        Set messager.ComArea = ca3231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            failed = True
            err_message5 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Do
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
            failed = True
            err_message5 = response
            Exit Do
        End If
      
        Dim rowsfetched As Integer
        rowsfetched = ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca3231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("TRDATE").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("ATERM_ID").Text & " " & _
                    rownode.selectSingleNode("CURRENCY_TRANS").Text & " " & _
                    rownode.selectSingleNode("TRANS_AMOUNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text & " " & _
                    rownode.selectSingleNode("ENT_AMNT").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca3231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca3231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .v2Value("IDATA/T_TIMESTAMP") = .v2Value("ODATA//T_TIMESTAMP")
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
      If failed = False Or (failed = True And allrows <> "") Then
        Dim res As String
        Dim rowsdoc As New MSXML2.DOMDocument30
        rowsdoc.LoadXML "<root>" & allrows & "</root>"
        
        res = L2PrintPassbook6_(owner.Manager.activeform, Mid(AccountNode.Text, 1, 10), 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, failed)

    End If
    
    If failed = False Then err_message5 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
    L2PrintExchangePassbookNew1 = err_message5


End Function


Public Function L2PrintPassbookNewVersion5(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim TRNTypeNode As IXMLDOMElement
    Set TRNTypeNode = GetXmlNode(inDocument, "//TRN_TYPE", "TRN_TYPE", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If TRNTypeNode Is Nothing Then Exit Function
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BRANCH", "Branch", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCOUNT", "Account", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BALANCE", "Balance", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//F_LINE", "Line", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim pageNode As IXMLDOMElement
    Set pageNode = GetXmlNode(inDocument, "//F_PAGE", "Page", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim ReasonNode As IXMLDOMElement
    Set ReasonNode = GetXmlNode(inDocument, "//REASON_CODE", "REASON", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AmountNode As IXMLDOMElement
    Set AmountNode = GetXmlNode(inDocument, "//TRANS_AMOUNT", "AMOUNT", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim TrnIDNode As IXMLDOMElement
    Set TrnIDNode = GetXmlNode(inDocument, "//TRANS_ID", "TRANS_ID", "Document L2PrintPassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    
    Dim allrows As String
    
    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca1231 As cXmlComArea
    Set ca1231 = DeclareComArea_("CALL_1231", "S1231", "S1231", "1231", "@S1231", "IDATA", "ODATA", owner.Manager.TrnBuffers)
    
    'διορθωση cXmlComArea.Container
    'Set ca1231.owner = owner.Manager
    Set ca1231.Container = owner.Manager.TrnBuffers
    
    With ca1231.Buffer
        .ClearData
        .v2Value("IDATA/TRANS_ID") = TrnIDNode.Text
        .v2Value("IDATA/BRANCH") = BranchNode.Text
        .v2Value("IDATA/ACCOUNT") = AccountNode.Text
        .v2Value("IDATA/PSBK_BALANCE") = BalanceNode.Text
        .v2Value("IDATA/F_LINE") = LineNode.Text
        .v2Value("IDATA/F_PAGE") = pageNode.Text
        If Not (AmountNode Is Nothing) Then .v2Value("IDATA/TRANS_AMOUNT") = AmountNode.Text
        If Not (ReasonNode Is Nothing) Then .v2Value("IDATA/REASON_CODE") = ReasonNode.Text
       
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    Do
        Dim response As String
        response = ca1231.ParseCall(Nothing)
       
        If response = "" Then
             L2PrintPassbookNewVersion5 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
             Exit Function
        End If
        
        Dim messager As New cXmlDepositMessageHandlerVersion3
        Set messager.ComArea = ca1231
        response = messager.LoadXML(response)
        
        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response
        
        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            L2PrintPassbookNewVersion5 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Function
'            response = "<MESSAGE><RC>999</RC><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
'            ResponseDoc.LoadXml response
        End If
        
        Dim resp As IXMLDOMElement
       
        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
        'If Not ca1231.HandleResp(ResponseDoc, "Λάθος...") Then
            L2PrintPassbookNewVersion5 = response
            Exit Function
        End If
      
        Dim rowsfetched As Integer
        rowsfetched = ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched
        
        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca1231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[CURRENCY!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("D_TRANS").Text & " " & _
                    rownode.selectSingleNode("CURRENCY").Text & " " & _
                    rownode.selectSingleNode("BRANCH_SND").Text & rownode.selectSingleNode("TERM_ID").Text & " " & _
                    rownode.selectSingleNode("UNP_AMOUNT").Text & " " & rownode.selectSingleNode("REASON_CODE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
            
            allrows = allrows & rownode.XML
        Next rownode
        
        If (ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text <> "0") Then
            linedata = "ΓΡΑΜΜΗ ΤΕΛΟΥΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA//L_LINE").Text
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        End If
        linedata = "ΑΛΛΕΣ ΓΡΑΜΜΕΣ: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/MORE_ROWS").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        linedata = "ΣΥΝ.ΓΡΑΜ.ΔΙΑΒ.: " & ca1231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/ROWS_FETCHED").Text
        eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)
        
        With ca1231.Buffer
            If Trim(.v2Value("ODATA/CDATA/MORE_ROWS")) = "Y" Then
                continueflag = True
                '.v2Data("IDATA/IROW") = .v2Data("ODATA//ROWS[" & rowsfetched - 1 & "]")
                .v2Value("IDATA/CNV_DATA/ROWS_COUNTER") = CStr(rows_counter)
                .ByName("ODATA").ClearData
            
            Else
                continueflag = False
            End If
        End With
        

    Loop While continueflag
    
    Dim res As String
    Dim rowsdoc As New MSXML2.DOMDocument30
    rowsdoc.LoadXML "<root>" & allrows & "</root>"
    
  
   ' res = L2PrintPassbook_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement)
   res = L2PrintPassbookVersion3_(owner.Manager.activeform, TRNTypeNode.Text, LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, BranchNode.Text, Mid(AccountNode.Text, 1, 7))
    L2PrintPassbookNewVersion5 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
End Function


Public Function L2PrintExchangePassbook(owner As cXMLLocalMethod, inDocument As IXMLDOMElement) As String
    
    Dim BranchNode As IXMLDOMElement
    Set BranchNode = GetXmlNode(inDocument, "//BR", "Branch", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    Dim AccountNode As IXMLDOMElement
    Set AccountNode = GetXmlNode(inDocument, "//ACCT", "Account", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If AccountNode Is Nothing Then Exit Function
    Dim IndBalanceNode As IXMLDOMElement
    Set IndBalanceNode = GetXmlNode(inDocument, "//IND_PSBK_BAL", "IndBal", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If IndBalanceNode Is Nothing Then Exit Function
    Dim BalanceNode As IXMLDOMElement
    Set BalanceNode = GetXmlNode(inDocument, "//PSBK_BAL", "Balance", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If BalanceNode Is Nothing Then Exit Function
    Dim ReservedNode As IXMLDOMElement
    Set ReservedNode = GetXmlNode(inDocument, "//RESERVED", "Reserved", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If ReservedNode Is Nothing Then Exit Function
    Dim LineNode As IXMLDOMElement
    Set LineNode = GetXmlNode(inDocument, "//PLINE", "Line", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If LineNode Is Nothing Then Exit Function
    Dim SPGR5Node As IXMLDOMElement
    Set SPGR5Node = GetXmlNode(inDocument, "//SPGR5", "SPGR5", "Document L2PrintExchangePassbook", "Πρόβλημα στη διαδικασία Ενημέρωσης Βιβλιαρίου...")
    If SPGR5Node Is Nothing Then Exit Function
    
    Dim allrows As String
    allrows = ""

    Dim ComareaDoc As New MSXML2.DOMDocument30
    Dim ca5231 As cXmlComArea
    Set ca5231 = DeclareComArea_("CALL_5231", "S5231", "S5231", "5231", "@S5231", "IDATA", "ODATA", owner.Manager.TrnBuffers)

    Set ca5231.Container = owner.Manager.TrnBuffers

    With ca5231.Buffer
        .ClearData
        .v2Value("IDATA/BR") = BranchNode.Text
        .v2Value("IDATA/ACCT") = AccountNode.Text
        .v2Value("IDATA/IND_PSBK_BAL") = IndBalanceNode.Text
        .v2Value("IDATA/PSBK_BAL") = BalanceNode.Text
        .v2Value("IDATA/RESERVED") = ReservedNode.Text
        
        .v2Value("IDATA/IMSG_DATA/OUT_MAX_MSG") = "1"
        .v2Value("IDATA/IMSG_DATA/CNTR_ROWS") = SPGR5Node.selectSingleNode("//ODATA/CDATA/OMSG_DATA/CNTR_ROWS").Text
        
        Dim i As Integer
        For i = 0 To SPGR5Node.SelectNodes("//ODATA/CDATA/OMSG_DATA/OMSG_TAB/STRMSG").length - 1
            Dim msgnode As IXMLDOMElement
            Set msgnode = SPGR5Node.SelectNodes("//ODATA/CDATA/OMSG_DATA/OMSG_TAB/STRMSG").Item(i)
            If Not msgnode Is Nothing Then
                If Trim(msgnode.selectSingleNode("./MSG_CODE").Text) <> "" Then
                    .v2Value("IDATA/IMSG_DATA/IMSG_TAB/STRMSG/MSG_CODE", i + 1) = msgnode.selectSingleNode("./MSG_CODE").Text
                    .v2Value("IDATA/IMSG_DATA/IMSG_TAB/STRMSG/MSG_KEYS", i + 1) = msgnode.selectSingleNode("./MSG_KEYS").Text
                End If
            End If
        Next i
    End With
    
    Dim repeatcount As Integer
    Dim continueflag As Boolean
    Dim rows_counter As Integer
    rows_counter = 0
    
    Dim failed As Boolean
    Dim err_message5 As String
    failed = False
    err_message5 = ""
    
    Do
        Dim response As String
        response = ca5231.ParseCall(Nothing)

        If response = "" Then
            failed = True
            err_message5 = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
            Exit Do
'             L2PrintExchangePassbook = "<MESSAGE><ERROR><LINE>Απέτυχε η κλήση της Συναλλαγής</LINE></ERROR></MESSAGE>"
'             Exit Function
        End If

        Dim messager As New cXmlExchangeDepositMessageHandler
        Set messager.ComArea = ca5231
        response = messager.LoadXML(response)

        Dim ResponseDoc As New MSXML2.DOMDocument30
        ResponseDoc.LoadXML response

        If (ResponseDoc.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            failed = True
            err_message5 = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
            Exit Do
'            L2PrintExchangePassbook = "<MESSAGE><ERROR><LINE>ΔΙΑΚΟΠΗ ΣΥΝΑΛΛΑΓΗΣ</LINE></ERROR></MESSAGE>"
'            Exit Function
        End If

        Dim resp As IXMLDOMElement

        Set resp = GetXmlNode(ResponseDoc.documentElement, "//RC", "RC", , "Πρόβλημα στο HandleResp...")
        If resp.Text <> "0" Then
            failed = True
            err_message5 = response
            Exit Do
'            L2PrintExchangePassbook = response
'            Exit Function
        End If

        Dim rowsfetched As Integer
        rowsfetched = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/ROWS_FETCHED").Text
        rows_counter = rows_counter + rowsfetched

        Dim linedata As String
        Dim rownode As IXMLDOMElement
        For Each rownode In ca5231.Buffer.xmlDocV2.SelectNodes("//ODATA/ROWS[PRINT_TYPE!='']")
            linedata = "ΓΡΑΜ: " & _
                    rownode.selectSingleNode("PRINT_TYPE").Text & " " & _
                    rownode.selectSingleNode("TRANS_DT").Text & " " & _
                    rownode.selectSingleNode("VALUE_DT").Text & " " & rownode.selectSingleNode("TERM_DEP_EXP_DT").Text & " " & _
                    rownode.selectSingleNode("RSN_CD").Text & " " & GetStrAmount_(CDbl(rownode.selectSingleNode("ENT_AMNT").Text), 15, 2) & " " & _
                    GetStrAmount_(CDbl(rownode.selectSingleNode("PSBK_BAL").Text), 15, 2) & " " & rownode.selectSingleNode("CUR_ISO").Text & " " & _
                    GetStrAmount_(CDbl(rownode.selectSingleNode("TRANS_AMNT").Text), 15, 2) & " " & Trim(rownode.selectSingleNode("SEND_BR").Text) & _
                    Trim(rownode.selectSingleNode("TERM_ID").Text) & " " & rownode.selectSingleNode("TERM_SEQ").Text & " " & _
                    GetStrAmount_(CDbl(rownode.selectSingleNode("INT_RATE").Text), 8, 3) & " " & rownode.selectSingleNode("CHEQUE_NBR").Text & " " & _
                    rownode.selectSingleNode("PRINT_STATUS").Text & " " & rownode.selectSingleNode("I_VALUTA").Text & " " & _
                    rownode.selectSingleNode("I_CALC_INT").Text & " " & rownode.selectSingleNode("I_REV").Text & " " & _
                    GetStrAmount_(CDbl(rownode.selectSingleNode("PARITE").Text), 10, 6) & " " & rownode.selectSingleNode("STAR_IND").Text

            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(linedata)

            allrows = allrows & rownode.XML
        Next rownode

        With ca5231.Buffer
            If rowsfetched = 50 Then
                continueflag = True
                .v2Value("IDATA/PSBK_BAL") = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/PSBK_BAL").Text
                .v2Value("IDATA/VALUE_DT") = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/VALUE_DT").Text
                .v2Value("IDATA/TSTMP") = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/TSTMP").Text
                .v2Value("IDATA/TRANS_DT") = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/TRANS_DT").Text
                .v2Value("IDATA/RESERVED") = ca5231.Buffer.xmlDocV2.selectSingleNode("//ODATA/CDATA/LST_DATA/RESERVED").Text
                .ByName("ODATA").ClearData
            Else
                continueflag = False
            End If
        End With

    Loop While continueflag
    
    If failed = False Or (failed = True And allrows <> "") Then
        Dim res As String
        Dim rowsdoc As New MSXML2.DOMDocument30
        rowsdoc.LoadXML "<root>" & allrows & "</root>"
        
        res = L2PrintPassbook5_(owner.Manager.activeform, AccountNode.Text, 0, "000", CDbl(0), CDbl(0), LineNode.Text, CDbl(BalanceNode.Text), rowsdoc.documentElement, failed)
'    Else
'        L2PrintExchangePassbook = err_message5
'        Exit Function
    End If
    
    If failed = False Then err_message5 = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"
    L2PrintExchangePassbook = err_message5

'    L2PrintExchangePassbook = "<MESSAGE><SUM_OF_UNPOSTED>" + res + "</SUM_OF_UNPOSTED></MESSAGE>"

End Function

Public Function L2CalcCheckDigits(inDocument As IXMLDOMElement) As String
 On Error GoTo ErrorPos
    Dim account As String
    Dim branch As String
    Dim accountforCD1 As String
    Dim accountforCD2 As String
    Dim cd1 As Integer, cd2 As Integer
    cd1 = -1
    cd2 = -1
    If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
        account = inDocument.selectSingleNode("//account").Text
    End If
   If Trim(account) <> "" And Len(account) >= 9 Then
        branch = Mid(account, 1, 3)
        accountforCD1 = Mid(account, 4, 6)
        cd1 = CalcCd1_(accountforCD1, 6)
        accountforCD2 = branch & accountforCD1 & CStr(cd1)
        cd2 = CalcCd2_(accountforCD2)
        L2CalcCheckDigits = "<MESSAGE>" + "<ACCOUNT>" & branch & accountforCD1 & "</ACCOUNT>" & _
        "<CD1>" & CStr(cd1) & "</CD1>" & "<CD2>" & CStr(cd2) & "</CD2>" & _
        "</MESSAGE>"
   Else
        L2CalcCheckDigits = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ</LINE></ERROR></MESSAGE>"
   End If

   Exit Function
ErrorPos:
    L2CalcCheckDigits = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function
Public Function L2CalcCd1(inDocument As IXMLDOMElement) As String
Dim account As String
Dim Digits As String
Dim cd1 As Integer
On Error GoTo ErrorPos
If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
   account = inDocument.selectSingleNode("//account").Text
End If
If Not (inDocument.selectSingleNode("//digits") Is Nothing) Then
   Digits = inDocument.selectSingleNode("//digits").Text
End If
If Trim(account) <> "" And Digits <> "" Then
    cd1 = CalcCd1_(account, CInt(Digits))
    L2CalcCd1 = "<MESSAGE><CD>" & cd1 & "</CD></MESSAGE>"
Else
    L2CalcCd1 = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ</LINE></ERROR></MESSAGE>"
End If
Exit Function
ErrorPos:
    L2CalcCd1 = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function
Public Function L2CalcCd2(inDocument As IXMLDOMElement) As String
Dim account As String
Dim cd2 As Integer
On Error GoTo ErrorPos
If Not (inDocument.selectSingleNode("//account") Is Nothing) Then
   account = inDocument.selectSingleNode("//account").Text
End If
If Trim(account) <> "" Then
    cd2 = CalcCd2_(Right(String("0", 10) & account, 10))
    L2CalcCd2 = "<MESSAGE><CD>" & cd2 & "</CD></MESSAGE>"
Else
    L2CalcCd2 = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ</LINE></ERROR></MESSAGE>"
End If
Exit Function
ErrorPos:
    L2CalcCd2 = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΨΗΦΙΟΥ ΕΛΕΓΧΟΥ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function


Public Function L2SuspectedTrnHandler(inDocument As IXMLDOMElement) As String
On Error GoTo ErrorPos
Dim i As Integer, messageno As Integer
Dim inPart As String, outpart As String
Dim inNode As IXMLDOMElement, outnode As IXMLDOMElement
Dim resultdoc As New MSXML2.DOMDocument30, resultnode As IXMLDOMElement, keytext As String
Dim messagedoc As New MSXML2.DOMDocument30
Dim Node As IXMLDOMElement
Dim requestedkey As String
    
    If Not (inDocument.selectSingleNode("//comarea") Is Nothing) Then
        If (inDocument.SelectNodes("//SKEYS").length <> inDocument.SelectNodes("//SKEYS1").length) Then
            L2SuspectedTrnHandler = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΥΠΟΠΤΩΝ ΣΥΝΑΛΛΑΓΩΝ:ΛΑΘΟΣ ΣΤΗΝ ΑΝΤΙΣΤΟΙΧΗΣΗ ΜΗΝΥΜΑΤΩΝ</LINE></ERROR></MESSAGE>"
            Exit Function
        End If
        
        messageno = inDocument.SelectNodes("//SKEYS1").length
        If inDocument.SelectNodes("//SKEYS").length > 0 Then inPart = inDocument.selectSingleNode("//SKEYS").parentNode.baseName
        If inDocument.SelectNodes("//SKEYS1").length > 0 Then outpart = inDocument.selectSingleNode("//SKEYS1").parentNode.baseName
        requestedkey = "0"
        For i = 0 To messageno - 1
            If inPart <> "" Then Set inNode = inDocument.selectSingleNode("//" & inPart & "[" & i & "]" & "/SKEYS/HASH_VALUE")
            If outpart <> "" Then Set outnode = inDocument.selectSingleNode("//" & outpart & "[" & i & "]" & "/SKEYS1/HASH_VALUE")
            If Not (inNode Is Nothing) And Not (outnode Is Nothing) Then inNode.Text = outnode.Text
            
            If inPart <> "" Then Set inNode = inDocument.selectSingleNode("//" & inPart & "[" & i & "]" & "/SKEYS/EFARMOGH")
            If outpart <> "" Then Set outnode = inDocument.selectSingleNode("//" & outpart & "[" & i & "]" & "/SKEYS1/EFARMOGH")
            If Not (inNode Is Nothing) And Not (outnode Is Nothing) Then inNode.Text = outnode.Text
            
            If inPart <> "" Then Set inNode = inDocument.selectSingleNode("//" & inPart & "[" & i & "]" & "/SKEYS/MSG_KWD")
            If outpart <> "" Then Set outnode = inDocument.selectSingleNode("//" & outpart & "[" & i & "]" & "/SKEYS1/MSG_KWD")
            If Not (inNode Is Nothing) And Not (outnode Is Nothing) Then inNode.Text = outnode.Text
            
            If inPart <> "" Then Set inNode = inDocument.selectSingleNode("//" & inPart & "[" & i & "]" & "/SKEYS/MSG_TYPE")
            If outpart <> "" Then Set outnode = inDocument.selectSingleNode("//" & outpart & "[" & i & "]" & "/SKEYS1/MSG_TYPE")
            If Not (inNode Is Nothing) And Not (outnode Is Nothing) Then inNode.Text = outnode.Text
                        
            If outnode.Text = "N" Then
                requestedkey = "N"
            ElseIf outnode.Text = "M" And requestedkey <> "N" Then
                requestedkey = "M"
            ElseIf outnode.Text = "C" And InStr("NM", requestedkey) = 0 Then
                requestedkey = "C"
            ElseIf outnode.Text = "Y" And InStr("NMC", requestedkey) = 0 Then
                requestedkey = "Y"
            End If
        Next
        
        Set resultnode = inDocument.ownerDocument.createElement("SUSPECTED_TRN_RESULT")
        Set Node = inDocument.selectSingleNode("//comarea")
        Node.appendChild resultnode
        If requestedkey = "Y" Then
            Set SuspectedTrnFrm.MessageDocument = Node.ownerDocument
            SuspectedTrnFrm.RequiredKey = ""
            SuspectedTrnFrm.Show vbModal
            
            If resultnode.Text <> "N" Then resultnode.Text = "Y"
        ElseIf requestedkey = "C" Then
            Set SuspectedTrnFrm.MessageDocument = Node.ownerDocument
            SuspectedTrnFrm.RequiredKey = "C"
            SuspectedTrnFrm.Show vbModal
            
            resultdoc.LoadXML SuspectedTrnFrm.KeyDoc.XML
            If resultdoc.SelectNodes("//MESSAGE/ERROR").length > 0 Then
               resultnode.Text = "N"
            Else
               If resultnode.Text <> "N" Then resultnode.Text = "Y"
            End If
        ElseIf requestedkey = "M" Then
            Set SuspectedTrnFrm.MessageDocument = Node.ownerDocument
            SuspectedTrnFrm.RequiredKey = "M"
            SuspectedTrnFrm.Show vbModal
                        
            resultdoc.LoadXML SuspectedTrnFrm.KeyDoc.XML
            If resultdoc.SelectNodes("//MESSAGE/ERROR").length > 0 Then
               resultnode.Text = "N"
            Else
               If resultnode.Text <> "N" Then resultnode.Text = "Y"
            End If
        ElseIf requestedkey = "N" Then
            Set SuspectedTrnFrm.MessageDocument = Node.ownerDocument
            SuspectedTrnFrm.RequiredKey = ""
            SuspectedTrnFrm.Show vbModal
            
            resultnode.Text = "N"
        End If

        L2SuspectedTrnHandler = inDocument.selectSingleNode("//comarea").XML
    Else
        L2SuspectedTrnHandler = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΥΠΟΠΤΩΝ ΣΥΝΑΛΛΑΓΩΝ</LINE></ERROR></MESSAGE>"
    End If
            
    Exit Function
ErrorPos:
    L2SuspectedTrnHandler = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΕΛΕΓΧΟΣ ΥΠΟΠΤΩΝ ΣΥΝΑΛΛΑΓΩΝ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"

End Function
Public Function L2BuildSwiftMessage(inDocument As IXMLDOMElement) As String
    Dim amessage As cSWIFTmessage
    Dim MessageCode As String
    Dim MessageBank As String
    Dim SenderBank As String
    Dim fieldlist As IXMLDOMNodeList
    Dim aaattr As IXMLDOMAttribute
    Dim formattedValueAttr As IXMLDOMAttribute
    Dim formattedElem As IXMLDOMCDATASection
    Dim i As Integer
    Dim ResultCode As Integer
    Dim ResultMessage As String
    
    SenderBank = "ETHNGRAAAXXX"
    ResultCode = 0
    ResultMessage = ""
    Set amessage = New cSWIFTmessage
    If Not (inDocument.selectSingleNode("//code") Is Nothing) Then
        MessageCode = inDocument.selectSingleNode("//code").Text
    Else
        L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & "ΜΗ ΣΥΜΠΛΗΡΩΜΕΝΟΣ ΚΩΔΙΚΟΣ ΜΗΝΥΜΑΤΟΣ" & "</LINE></ERROR></MESSAGE>"
        Exit Function
    End If
    If Not (inDocument.selectSingleNode("//bank") Is Nothing) Then
        MessageBank = inDocument.selectSingleNode("//bank").Text
    Else
        L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & "ΜΗ ΣΥΜΠΛΗΡΩΜΕΝΗ ΤΡΑΠΕΖΑ ΛΗΨΗΣ ΜΗΝΥΜΑΤΟΣ" & "</LINE></ERROR></MESSAGE>"
        Exit Function
    End If
    If Not (inDocument.selectSingleNode("//senderbank") Is Nothing) Then
        If inDocument.selectSingleNode("//senderbank").Text = "" Then
            L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & "ΜΗ ΣΥΜΠΛΗΡΩΜΕΝΗ ΤΡΑΠΕΖΑ ΑΠΟΣΤΟΛΗΣ ΜΗΝΥΜΑΤΟΣ" & "</LINE></ERROR></MESSAGE>"
            Exit Function
        Else
            SenderBank = inDocument.selectSingleNode("//senderbank").Text
        End If
    End If
    amessage.prepare MessageCode
     
    Set fieldlist = inDocument.SelectNodes("//field")
    If Not (fieldlist Is Nothing) Then
        For i = 0 To fieldlist.length - 1
           Set aaattr = fieldlist(i).Attributes.getNamedItem("aa")
           If Not (aaattr Is Nothing) Then
             amessage.Value(aaattr.Value) = fieldlist(i).Text
             If amessage.ResultCode <> 0 Then
                ResultCode = ResultCode + amessage.ResultCode
                ResultMessage = ResultMessage & amessage.ResultMessage & vbCrLf
             End If
             
             Set formattedValueAttr = amessage.messagedoc.createAttribute("formattedvalue")
             formattedValueAttr.Value = amessage.FormatedValue(aaattr.Value)
             amessage.messagedoc.selectSingleNode("//field[@aa=" & "'" & aaattr.Value & "'" & "]").Attributes.setNamedItem formattedValueAttr
           End If
        Next i
    Else
        L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & "ΜΗ ΣΥΜΠΛΗΡΩΜΕΝΑ ΠΕΔΙΑ ΜΗΝΥΜΑΤΟΣ" & "</LINE></ERROR></MESSAGE>"
        Exit Function
    End If
    If ResultCode <> 0 Then
       L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & ResultMessage & "</LINE></ERROR></MESSAGE>"
       Exit Function
    End If
    On Error GoTo ErrorPos
    Dim headerstr As String
    headerstr = "{1:F01" & SenderBank & "0000000000}{2:I" & Replace(MessageCode, "MT", "") & Mid(MessageBank, 1, 8) & "X" & Mid(MessageBank, 9, 3) & "N}{3:{108:EXP}}{4:" & "Ω" & vbLf
    Set formattedElem = amessage.messagedoc.createCDATASection(headerstr & Replace(amessage.FormatedText, vbCrLf, "Ω" & vbLf) & "-}")
    amessage.messagedoc.documentElement.appendChild formattedElem
    
    Dim Lines() As String
    Lines = amessage.PrintSwiftMessage
    Dim j As Integer
    Dim printElem As IXMLDOMElement
    Set printElem = amessage.messagedoc.createElement("print")
    Dim lineElem As IXMLDOMElement
    Dim lineAttr As IXMLDOMAttribute
    For j = 0 To UBound(Lines)
        Set lineElem = amessage.messagedoc.createElement("line")
        Set lineAttr = amessage.messagedoc.createAttribute("text")
        lineAttr.Value = Lines(j)
        lineElem.Attributes.setNamedItem lineAttr
        printElem.appendChild lineElem
    Next
    amessage.messagedoc.documentElement.appendChild printElem
    L2BuildSwiftMessage = amessage.messagedoc.XML
    Exit Function
ErrorPos:
    L2BuildSwiftMessage = "<MESSAGE><ERROR><LINE>" & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2CalculateTerminalID(inDocument As IXMLDOMElement) As String
 On Error GoTo ErrorPos
    Dim machine As String
    Dim res As String
    
    Dim branch As String
    
    If Not (inDocument.selectSingleNode("//machine") Is Nothing) Then
        machine = inDocument.selectSingleNode("//machine").Text
    End If
    If Not (inDocument.selectSingleNode("//branch") Is Nothing) Then
        branch = inDocument.selectSingleNode("//branch").Text
    End If
   If Trim(machine) <> "" Then
        res = MachineToTerminal_internal(machine, branch)
        If (Trim(res) <> "") Then
            L2CalculateTerminalID = res
        Else
               L2CalculateTerminalID = "<MESSAGE>" + "<MACHINE>" & machine & "</MACHINE><EL/><EN/><CICS/>" & _
             "</MESSAGE>"
        End If
        
'
'        L2CalculateTerminalID = "<MESSAGE>" + "<MACHINE>" & machine & "</MACHINE>" & _
'        "<EL>" & EL & "</EL>" & "<EN>" & EN & "</EN>" & "<CICS>" & CICS & "</CICS>" & _
'        "</MESSAGE>"
   Else
        L2CalculateTerminalID = "<MESSAGE>" + "<MACHINE>" & machine & "</MACHINE><EL/><EN/><CICS/>" & _
        "</MESSAGE>"
   End If

   Exit Function
ErrorPos:
    L2CalculateTerminalID = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ TERMID" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2CalculateIRISTime(inDocument As IXMLDOMElement) As String
    On Error GoTo ErrorPos
    Dim IRISTime As String
    If Not (inDocument.selectSingleNode("//time") Is Nothing) Then
        IRISTime = inDocument.selectSingleNode("//time").Text
    End If
    If Trim(IRISTime) <> "" Then
        Dim ti As Long
        Dim h, m, s As Integer
        Dim tTime
        ti = CLng(IRISTime) / 1000
        h = Int(ti / 3600)
        m = Int((ti Mod 3600) / 60)
        s = Int((ti Mod 3600) Mod 60)
        tTime = TimeSerial(h, m, s)
        L2CalculateIRISTime = "<MESSAGE>" + "<TIME>" & tTime & "</TIME></MESSAGE>"
    Else
        L2CalculateIRISTime = "<MESSAGE>" + "<TIME/>" & "</MESSAGE>"
    End If
    Exit Function
ErrorPos:
    L2CalculateIRISTime = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Ο ΥΠΟΛΟΓΙΣΜΟΣ ΩΡΑΣ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"

End Function

Public Function L2ShowIRISMessages(inDocument As IXMLDOMElement)
    On Error GoTo ErrorPos
    Dim MsgElementList As IXMLDOMNodeList
    Set MsgElementList = inDocument.SelectNodes("//STD_AN_AV_MSJ_LS[COD_ANTCN!='']")
    If Not (MsgElementList Is Nothing) Then
       If MsgElementList.length > 0 Then
          Dim aFrm As New IRISMsgFrm
          Set aFrm.MsgView = Nothing
          Set aFrm.MsgViewXML = inDocument
          aFrm.Show vbModal
          Set aFrm = Nothing
       End If
    End If
    L2ShowIRISMessages = "<MESSAGE/>"
    Exit Function
ErrorPos:
    L2ShowIRISMessages = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΜΦΑΝΙΣΗ ΜΗΝΥΜΑΤΩΝ ΔΑΝΕΙΩΝ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2DateAdd(inDocument As IXMLDOMElement)
    On Error GoTo ErrorPos
    Dim datepart As String
    Dim number As Double
    Dim indate As String
    If Not (inDocument.selectSingleNode("//datepart") Is Nothing) Then
        datepart = inDocument.selectSingleNode("//datepart").Text
    Else
        datepart = "d"
    End If
    If Not (inDocument.selectSingleNode("//number") Is Nothing) Then
        number = CDbl(inDocument.selectSingleNode("//number").Text)
    Else
        number = 1
    End If
    If Not (inDocument.selectSingleNode("//date") Is Nothing) Then
        indate = inDocument.selectSingleNode("//date").Text
    Else
        indate = Date
    End If
    L2DateAdd = "<MESSAGE>" & format(DateAdd(datepart, number, indate), "DDMMYYYY") & "</MESSAGE>"
    Exit Function
ErrorPos:
    L2DateAdd = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ O ΥΠΟΛΟΓΙΣΜΟΣ ΗΜΕΡΟΜΗΝΙΑΣ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2DateDiff(inDocument As IXMLDOMElement)
 On Error GoTo ErrorPos
    Dim interval As String
    Dim date1 As String
    Dim date2 As String
    If Not (inDocument.selectSingleNode("//interval") Is Nothing) Then
        interval = inDocument.selectSingleNode("//interval").Text
    Else
        interval = "d"
    End If
    If Not (inDocument.selectSingleNode("//date1") Is Nothing) Then
        date1 = inDocument.selectSingleNode("//date1").Text
    Else
        date1 = Date
    End If
    If Not (inDocument.selectSingleNode("//date2") Is Nothing) Then
        date2 = inDocument.selectSingleNode("//date2").Text
    Else
        date2 = Date
    End If
    L2DateDiff = "<MESSAGE>" & DateDiff(interval, date1, date2) & "</MESSAGE>"
    Exit Function
ErrorPos:
    L2DateDiff = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ O ΥΠΟΛΟΓΙΣΜΟΣ ΗΜΕΡΟΜΗΝΙΑΣ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function

Public Function L2ChkDocumentNo(inDocument As IXMLDOMElement)
On Error GoTo ErrorPos
  Dim Country As String
  Dim Document As String
  Dim Result As Boolean
  Result = True
  If Not (inDocument.selectSingleNode("//Country") Is Nothing) Then
        Country = inDocument.selectSingleNode("//Country").Text
  End If
  If Not (inDocument.selectSingleNode("//Document") Is Nothing) Then
        Document = inDocument.selectSingleNode("//Document").Text
  End If
  If Country <> "" And Document <> "" Then
     Result = ChkDocumentNo_(Country, Document)
     L2ChkDocumentNo = "<MESSAGE>" & Result & "</MESSAGE>"
     Exit Function
  End If
  L2ChkDocumentNo = "<MESSAGE>" & Result & "</MESSAGE>"
  Exit Function
ErrorPos:
    L2ChkDocumentNo = "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ O ΕΛΕΓΧΟΣ ΕΓΓΡΑΦΟΥ" & Err.number & " " & Err.description & "</LINE></ERROR></MESSAGE>"
End Function


Public Function L2ValidateCheckInput(inDocument As IXMLDOMElement)
    Dim BankCD As Integer
    Dim BankCheck As Long
    Dim IBANflag As String
    Dim IBAN As String
    Dim BankBranchCD As Integer
    Dim BankAccount As Double
    Dim BankAmount As Double
    Dim strError As String
    Dim astr As String
    Dim fList
    Dim atotal As Long
    Dim i, arem As Integer
    Dim ChequeType As Integer
    Dim iskeaefrequest As Integer
    
    If Not (inDocument.selectSingleNode("//BANKCD") Is Nothing) Then
        BankCD = CInt(inDocument.selectSingleNode("//BANKCD").Text)
    End If
    If Not (inDocument.selectSingleNode("//BANKCHECK") Is Nothing) Then
        BankCheck = CLng(inDocument.selectSingleNode("//BANKCHECK").Text)
    End If
    If Not (inDocument.selectSingleNode("//IBANFLAG") Is Nothing) Then
        IBANflag = inDocument.selectSingleNode("//IBANFLAG").Text
    End If
    If Not (inDocument.selectSingleNode("//IBAN") Is Nothing) Then
        IBAN = inDocument.selectSingleNode("//IBAN").Text
    End If
    If Not (inDocument.selectSingleNode("//BANKBRANCHCD") Is Nothing) Then
        BankBranchCD = CInt(inDocument.selectSingleNode("//BANKBRANCHCD").Text)
    End If
    If Not (inDocument.selectSingleNode("//BANKACCOUNT") Is Nothing) Then
        BankAccount = CDbl(inDocument.selectSingleNode("//BANKACCOUNT").Text)
    End If
    If Not (inDocument.selectSingleNode("//BANKAMOUNT") Is Nothing) Then
        BankAmount = CDbl(inDocument.selectSingleNode("//BANKAMOUNT").Text)
    End If
    If Not (inDocument.selectSingleNode("//CHECKTYPE") Is Nothing) Then
        If Trim(inDocument.selectSingleNode("//CHECKTYPE").Text) <> "" Then
            ChequeType = CInt(inDocument.selectSingleNode("//CHECKTYPE").Text)
        End If
    End If
    If Not (inDocument.selectSingleNode("//ISKEAEFREQUEST") Is Nothing) Then
        If Trim(inDocument.selectSingleNode("//ISKEAEFREQUEST").Text) <> "" Then
            iskeaefrequest = CInt(inDocument.selectSingleNode("//ISKEAEFREQUEST").Text)
        End If
    End If
    
    If BankCD = 97 Then
       astr = Right("000" & BankCD, 3) & Right("000" & BankBranchCD, 3) & _
            Right("0000000000000" & BankAccount, 13)
       fList = Array(3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
       atotal = 0
       For i = 1 To 18
         atotal = atotal + Mid(astr, i, 1) * fList(i - 1)
       Next
       arem = (11 - (atotal Mod 11)) Mod 10
       If CStr(arem) <> Right(astr, 1) Then
          strError = "ΛΑΘΟΣ ΛΟΓΑΡΙΑΣΜΟΣ": GoTo ErrorPos
       End If
       fList = Array(9, 8, 7, 6, 5, 4, 3, 2)
       astr = Right("0000000000000" & BankCheck, 9)
       atotal = 0
       For i = 1 To 8
         atotal = atotal + Mid(astr, i, 1) * fList(i - 1)
       Next
       arem = (11 - (atotal Mod 11)) Mod 10
       If CStr(arem) <> Right(astr, 1) Then
          strError = "ΛΑΘΟΣ ΑΡΙΘΜΟΣ ΕΠΙΤΑΓΗΣ": GoTo ErrorPos
       End If
    End If
    
    If BankCD <> 11 And BankAccount <> 0 Then
       If BankCD = 77 And ChequeType = 9 Then
       
       Else
           If Not TRNFrm.ChkBankAcount(BankCD, BankBranchCD, BankAccount, ChequeType) And (BankCD <> 97) Then
              If BankCD = 54 And BankBranchCD = 800 And BankAccount = 157000000 Then
                 'probank τραπεζικες επιταγές
              Else
                 strError = "ΛΑΘΟΣ ΛΟΓΑΡΙΑΣΜΟΣ": GoTo ErrorPos
              End If
           End If
       End If
    End If
    If BankCD = 11 Then
       astr = Right("000" & BankBranchCD, 3) & Right("00000000" & BankAccount, 8)
       If Not TRNFrm.ChkFldType(astr, 2) Then
          strError = "ΛΑΘΟΣ ΛΟΓΑΡΙΑΣΜΟΣ": GoTo ErrorPos
       End If
    End If
    If BankCheck <> 0 Then
       If Not TRNFrm.ChkBankCheque(BankCD, BankBranchCD, BankAccount, BankCheck, ChequeType) Then
          strError = "ΛΑΘΟΣ ΕΠΙΤΑΓΗ": GoTo ErrorPos
       End If
    End If
    If BankCD = 11 Or Right("0000" & Trim(IBANflag), 4) = "0011" Then
       If Not TRNFrm.ChkETECheque(BankCheck) Then
          strError = "ΛΑΘΟΣ ΕΠΙΤΑΓΗ": GoTo ErrorPos
       End If
    End If
'    If iskeaefrequest = 1 Then
'    Else
'        If BankCD <> 11 And Right("0000" & Trim(IBANflag), 4) <> "0011" Then
'           If BankAmount > 30000000 Then
'              strError = "ΜΗ ΑΠΟΔΕΚΤΟ ΠΟΣΟ": GoTo ErrorPos
'           End If
'        End If
'    End If
    L2ValidateCheckInput = "<MESSAGE/>"
    Exit Function
    On Error GoTo ErrorPos
ErrorPos:
    L2ValidateCheckInput = "<MESSAGE><ERROR><LINE>" & strError & "</LINE></ERROR></MESSAGE>"
End Function

