Attribute VB_Name = "TerminalIdentityFunctions"
Option Explicit

'Private ReadDir As String
Private PU As String
Private LU As String
Dim terminalsdoc As MSXML2.DOMDocument30

Public Function MachineToTerminal(MachineName As String, language As String) As String
    'language "EN", "EL", "CICS"
    
    Dim ServerLetter As String
    Dim TerminalLetter As String
    Dim branch As String
    Dim termArray As Variant
    
    termArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
   
    Set terminalsdoc = XmlLoadFile(ReadDir & "\TerminalName.xml", "MapTerminal ... Δεν βρέθηκε το TerminalName.xml", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If terminalsdoc Is Nothing Then Exit Function
    
    PU = ExtractPU(MachineName)
    LU = ExtractLU(MachineName)
    branch = ExtractBranch(MachineName)
    
    Dim servelm As IXMLDOMNode
    Dim termelm As IXMLDOMNode
    Dim luelm As IXMLDOMElement
    Dim min As IXMLDOMAttribute
    Dim max As IXMLDOMAttribute
    Dim refersto As IXMLDOMAttribute
    
    Set servelm = terminalsdoc.documentElement.selectSingleNode("//SERVER[@id='" + PU + "']")
    If Not (servelm Is Nothing) Then
        For Each luelm In servelm.childNodes
           Set min = luelm.Attributes.getNamedItem("min")
           Set max = luelm.Attributes.getNamedItem("max")
           If (CInt(LU) >= min.value And CInt(LU) <= max.value) Then
                Set refersto = luelm.Attributes.getNamedItem("refersto")
                ServerLetter = refersto.value
                Exit For
           End If
        Next
    Else
        MachineToTerminal = cTERMINALID
        Exit Function
    End If
    
    Set termelm = terminalsdoc.documentElement.selectSingleNode("//TERM[@id='" + LU + "']")
    If Not (termelm Is Nothing) Then
        TerminalLetter = termelm.Text
    Else
        Dim includeelm As IXMLDOMNode
        Dim modattr As IXMLDOMAttribute
        Dim baseattr As IXMLDOMAttribute
        Set includeelm = terminalsdoc.documentElement.selectSingleNode("//INCLUDES[@min <= " & LU & " and @max >=" & LU & "]")
        If (includeelm Is Nothing) Then MachineToTerminal = cTERMINALID: Exit Function
        Set modattr = includeelm.Attributes.getNamedItem("mod")
        Set baseattr = includeelm.Attributes.getNamedItem("min")
        Dim calc As Integer
        calc = 0
        calc = CalculateLULetter(CInt(LU), CInt(modattr.value), CInt(baseattr.value))
        If calc > 0 Then
            Set termelm = terminalsdoc.documentElement.selectSingleNode("//TERM[@id='" + Right("000" & CStr(calc), 3) + "']")
            If Not (termelm Is Nothing) Then
                TerminalLetter = termelm.Text
            End If
        End If
    End If
   
   If TerminalLetter = "" Or ServerLetter = "" Then
        MachineToTerminal = cTERMINALID
        Exit Function
   End If
   Dim res As String
   res = branch + TerminalLetter + ServerLetter
   If language = "EN" Then MachineToTerminal = res: Exit Function
   
   Dim eltranlated As String
   eltranlated = branch + TranslateHostLetter(TerminalLetter) + TranslateHostLetter(ServerLetter)
   If language = "EL" Then MachineToTerminal = eltranlated: Exit Function
   
   
   Dim hostID As String
   Dim remainder As Integer
   Dim div As Integer
   remainder = CInt(branch) Mod 36
   div = CInt(branch) \ 36
   hostID = termArray(div) + termArray(remainder) + TerminalLetter + ServerLetter
   If language = "CICS" Then MachineToTerminal = hostID: Exit Function
End Function

Function CalculateLULetter(LU As Integer, factor As Integer, start As Integer) As String
    Dim temp1, temp2, temp3 As Integer
    
    temp1 = LU - start
    temp2 = temp1 Mod factor
    temp3 = temp2 + start
    
    CalculateLULetter = temp3
End Function

'Public Function MachineToTerminal(MachineName As String, language As String) As String
'    'language "EN", "EL", "CICS"
'
'    Dim ServerLetter As String
'    Dim TerminalLetter As String
'    Dim branch As String
'    Dim termArray As Variant
'
'    termArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
'
'
'    Dim terminalsdoc As MSXML2.DOMDocument30
'    Set terminalsdoc = XmlLoadFile(ReadDir & "\hostterminals.xml", "MapTerminal", "Πρόβλημα στη διαδικασία Σύνδεσης...")
'    If terminalsdoc Is Nothing Then Exit Function
'
'    Dim serversdoc As MSXML2.DOMDocument30
'    Set serversdoc = XmlLoadFile(ReadDir & "\hostservers.xml", "MapTerminal", "Πρόβλημα στη διαδικασία Σύνδεσης...")
'    If serversdoc Is Nothing Then Exit Function
'
'    Dim lettersdoc As MSXML2.DOMDocument30
'    Set lettersdoc = XmlLoadFile(ReadDir & "\hostletters.xml", "MapTerminal", "Πρόβλημα στη διαδικασία Σύνδεσης...")
'    If lettersdoc Is Nothing Then Exit Function
'
'    PU = ExtractPU(MachineName)
'    LU = ExtractLU(MachineName)
'    branch = ExtractBranch(MachineName)
'
'    Dim servelm As IXMLDOMNode
'    Dim termelm As IXMLDOMNode
'    Dim luelm As IXMLDOMElement
'    Dim min As IXMLDOMAttribute
'    Dim max As IXMLDOMAttribute
'    Dim refersto As IXMLDOMAttribute
'
'    Set servelm = serversdoc.documentElement.selectSingleNode("//SERVER[@id='" + PU + "']")
'    If Not (servelm Is Nothing) Then
'        For Each luelm In servelm.childNodes
'           Set min = luelm.Attributes.getNamedItem("min")
'           Set max = luelm.Attributes.getNamedItem("max")
'           If (CInt(LU) >= min.value And CInt(LU) <= max.value) Then
'                Set refersto = luelm.Attributes.getNamedItem("refersto")
'                ServerLetter = refersto.value
'                Exit For
'           End If
'        Next
'    End If
'    Set termelm = terminalsdoc.documentElement.selectSingleNode("//TERM[@id='" + LU + "']")
'    If Not (termelm Is Nothing) Then
'        TerminalLetter = termelm.Text
'    End If
'
'   If TerminalLetter = "" Or ServerLetter = "" Then
'        MachineToTerminal = cTERMINALID
'        Exit Function
'   End If
'   Dim res As String
'   res = branch + TerminalLetter + ServerLetter
'   If language = "EN" Then MachineToTerminal = res: Exit Function
'
'   Dim eltranlated As String
'   eltranlated = branch + TranslateHostLetter(TerminalLetter) + TranslateHostLetter(ServerLetter)
'   If language = "EL" Then MachineToTerminal = eltranlated: Exit Function
'
'
'   Dim hostID As String
'   Dim remainder As Integer
'   Dim div As Integer
'   remainder = CInt(branch) Mod 36
'   div = CInt(branch) \ 36
'   hostID = termArray(div) + termArray(remainder) + TerminalLetter + ServerLetter
'   If language = "CICS" Then MachineToTerminal = hostID: Exit Function
'End Function

Function ExtractPU(machine As String) As String
        ExtractPU = Mid(machine, 6, 1)
End Function
Function ExtractLU(machine As String) As String
    If Len(machine) = 9 Then
        ExtractLU = Mid(machine, 7, 3)
    ElseIf Len(machine) = 10 Then
        ExtractLU = Mid(machine, 8, 3)
    End If
End Function
Function ExtractBranch(machine As String) As String
    ExtractBranch = Mid(machine, 3, 3)
End Function

Function TranslateHostLetter(letter As String) As String
    
    If IsNumeric(letter) Then TranslateHostLetter = letter: Exit Function
    
'    Dim lettersdoc As MSXML2.DOMDocument30
'    Set lettersdoc = XmlLoadFile(ReadDir & "\hostletters.xml", "MapTerminal", "Πρόβλημα στη διαδικασία Σύνδεσης...")
'    If lettersdoc Is Nothing Then Exit Function
    
    Dim elm As IXMLDOMNode
    Set elm = terminalsdoc.documentElement.selectSingleNode("//LETTER[@en='" + letter + "']")
    If Not (elm Is Nothing) Then
        Dim EL As IXMLDOMAttribute
        Set EL = elm.Attributes.getNamedItem("el")
        TranslateHostLetter = EL.value
        Exit Function
    End If
    TranslateHostLetter = ""
End Function


 Public Function MachineToTerminal_internal(MachineName As String, BranchCode As String) As String
    'language "EN", "EL", "CICS"
    Dim language As String
    
    Dim ServerLetter As String
    Dim TerminalLetter As String
    Dim branch As String
    Dim termArray As Variant
    Dim machine As String

    termArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
   
    Set terminalsdoc = XmlLoadFile(ReadDir & "\TerminalName.xml", "MapTerminal ... Δεν βρέθηκε το TerminalName.xml", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If terminalsdoc Is Nothing Then Exit Function
    
    machine = Right("_________" & MachineName, 9)
    
    PU = ExtractPU(machine)
    LU = ExtractLU(machine)
    
    If (BranchCode <> "") Then
        branch = BranchCode
    Else
        branch = ExtractBranch(machine)
    End If
    
    Dim servelm As IXMLDOMNode
    Dim termelm As IXMLDOMNode
    Dim luelm As IXMLDOMElement
    Dim min As IXMLDOMAttribute
    Dim max As IXMLDOMAttribute
    Dim refersto As IXMLDOMAttribute
    
    Set servelm = terminalsdoc.documentElement.selectSingleNode("//SERVER[@id='" + PU + "']")
    If Not (servelm Is Nothing) Then
        For Each luelm In servelm.childNodes
           Set min = luelm.Attributes.getNamedItem("min")
           Set max = luelm.Attributes.getNamedItem("max")
           If (CInt(LU) >= min.value And CInt(LU) <= max.value) Then
                Set refersto = luelm.Attributes.getNamedItem("refersto")
                ServerLetter = refersto.value
                Exit For
           End If
        Next
    Else
        MachineToTerminal_internal = ""
        Exit Function
    End If
    
    Set termelm = terminalsdoc.documentElement.selectSingleNode("//TERM[@id='" + LU + "']")
    If Not (termelm Is Nothing) Then
        TerminalLetter = termelm.Text
    Else
        Dim includeelm As IXMLDOMNode
        Dim modattr As IXMLDOMAttribute
        Dim baseattr As IXMLDOMAttribute
        Set includeelm = terminalsdoc.documentElement.selectSingleNode("//INCLUDES[@min <= " & LU & " and @max >=" & LU & "]")
        If (includeelm Is Nothing) Then MachineToTerminal_internal = "": Exit Function
        Set modattr = includeelm.Attributes.getNamedItem("mod")
        Set baseattr = includeelm.Attributes.getNamedItem("min")
        Dim calc As Integer
        calc = 0
        calc = CalculateLULetter(CInt(LU), CInt(modattr.value), CInt(baseattr.value))
        If calc > 0 Then
            Set termelm = terminalsdoc.documentElement.selectSingleNode("//TERM[@id='" + Right("000" & CStr(calc), 3) + "']")
            If Not (termelm Is Nothing) Then
                TerminalLetter = termelm.Text
            End If
        End If
    End If
   
   If TerminalLetter = "" Or ServerLetter = "" Then
        MachineToTerminal_internal = ""
        Exit Function
   End If
   
   Dim termid As String
   termid = branch + TerminalLetter + ServerLetter
   
   
    Dim EL  As String
    Dim EN  As String
    Dim CICS  As String
   
    If (termid <> "" And Len(termid) = 5) Then
        branch = Mid(termid, 1, 3)
        TerminalLetter = Mid(termid, 4, 1)
        ServerLetter = Mid(termid, 5, 1)
         
         EN = termid
         EL = branch + TranslateHostLetter(TerminalLetter) + TranslateHostLetter(ServerLetter)
         
         Dim hostID As String
         Dim remainder As Integer
         Dim div As Integer
         remainder = CInt(branch) Mod 36
         div = CInt(branch) \ 36
         CICS = termArray(div) + termArray(remainder) + TerminalLetter + ServerLetter
    
   
        MachineToTerminal_internal = "<MESSAGE>" + "<MACHINE>" & MachineName & "</MACHINE>" & _
        "<EL>" & EL & "</EL>" & "<EN>" & EN & "</EN>" & "<CICS>" & CICS & "</CICS>" & _
        "</MESSAGE>"
        Exit Function
   End If
   MachineToTerminal_internal = ""
End Function
