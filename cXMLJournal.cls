VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Manager As cXMLDocumentManager
Public StaticLink As IXMLDOMElement

'<journal trnclass="xxxxx" trncounter="before" >
'   <form>
'       <textbox name="" caption="" />
'       <textbox name="" label="" />
'       <label name="" caption="" />
'       <checkbox name="" caption="" />
'   </form>
'   <line>xxxx<line>
'   <line>xxxxxxx<line>
'</journal>

Public Function ProcessForm(inElement As IXMLDOMElement) As String
    Dim aTrnClass As String
    Dim TrnClassAttr As IXMLDOMAttribute
    Set TrnClassAttr = StaticLink.Attributes.getNamedItem("trnclass")
    
    If TrnClassAttr Is Nothing Then
        aTrnClass = cTRNCode
    ElseIf TrnClassAttr.value = "" Then
        LogMsgbox "Λάθος: Δεν περιλαμβάνεται κωδικός συναλλαγής στην δήλωση του ημερολογίου...", vbOKOnly, "Λάθος Δήλωσης(cXmlJournal.ProcessForm)"
        Exit Function
    Else
        aTrnClass = TrnClassAttr.value
    End If
    
    Dim child As IXMLDOMNode
    Dim childelement As IXMLDOMElement
    Dim captionAttr As IXMLDOMAttribute
    Dim labelattr As IXMLDOMAttribute
    Dim captionElement As IXMLDOMElement
    Dim valueattr As IXMLDOMAttribute
    Dim i As Integer
    Dim j As Integer
    
    For Each child In inElement.childNodes
        If child.nodeType = NODE_ELEMENT Then
            Set childelement = child
            Set captionAttr = Nothing
            Set valueattr = Nothing
            
            If childelement.baseName = "textbox" Then
                Set valueattr = childelement.Attributes.getNamedItem("outtext")
                If valueattr Is Nothing Then Set valueattr = childelement.Attributes.getNamedItem("text")
                If (valueattr Is Nothing) Or (valueattr.value = "") Then
                    Set valueattr = childelement.Attributes.getNamedItem("text")
                End If
            ElseIf childelement.baseName = "checkbox" Then
                Set valueattr = childelement.Attributes.getNamedItem("value")
            ElseIf childelement.baseName = "combobox" Then
                Set valueattr = childelement.Attributes.getNamedItem("text")
            ElseIf childelement.baseName = "statusmessage" Then
                Set valueattr = Nothing
            ElseIf childelement.baseName = "grid" Then
                Set valueattr = Nothing
            ElseIf childelement.baseName = "listbox" Then
                Set valueattr = Nothing
            End If
                
            If (Not (valueattr Is Nothing)) Or childelement.baseName = "statusmessage" Or childelement.baseName = "grid" Or childelement.baseName = "listbox" Then
                If Not valueattr Is Nothing Then
                        If childelement.baseName = "textbox" Or childelement.baseName = "combobox" Then
                            Set captionAttr = childelement.Attributes.getNamedItem("caption")
                            If captionAttr Is Nothing Then
                                Set labelattr = childelement.Attributes.getNamedItem("label")
                                If Not (labelattr Is Nothing) Then
                                    If Trim(labelattr.value) <> "" Then
                                        Set captionElement = inElement.selectSingleNode("./label[@name='" & UCase(labelattr.value) & "']")
                                        If Not captionElement Is Nothing Then
                                            Set captionAttr = captionElement.Attributes.getNamedItem("caption")
                                        End If
                                    End If
                                End If
                            ElseIf Trim(captionAttr.value) = "" Then
                                Set labelattr = childelement.Attributes.getNamedItem("label")
                                If Not (labelattr Is Nothing) Then
                                    If Trim(labelattr.value) <> "" Then
                                        Set captionElement = inElement.selectSingleNode("./label[@name='" & UCase(labelattr.value) & "']")
                                        If Not captionElement Is Nothing Then
                                            Set captionAttr = captionElement.Attributes.getNamedItem("caption")
                                        End If
                                    End If
                                End If
                            End If
                            If captionAttr Is Nothing Then
                                Set captionAttr = childelement.Attributes.getNamedItem("name")
                            ElseIf Trim(captionAttr.value) = "" Then
                                Set captionAttr = childelement.Attributes.getNamedItem("name")
                            End If
                        ElseIf childelement.baseName = "checkbox" Then
                            Set captionAttr = childelement.Attributes.getNamedItem("caption")
                            If captionAttr Is Nothing Then
                                Set captionAttr = childelement.Attributes.getNamedItem("name")
                            End If
                        End If
                        
                        If Not (captionAttr Is Nothing) Then
                            If Trim(captionAttr.value) <> "" Then
                                WriteJournal Trim(captionAttr.value) & ": " & valueattr.value, aTrnClass, cTRNNum
                            End If
                        End If
                Else
                    If childelement.baseName = "statusmessage" Then
                        Set captionAttr = childelement.Attributes.getNamedItem("caption")
                        If Not captionAttr Is Nothing Then
                            WriteJournal Trim(captionAttr.value) & ": " & childelement.Text, aTrnClass, cTRNNum
                        End If
                    ElseIf childelement.baseName = "grid" Then
                        Set captionAttr = childelement.Attributes.getNamedItem("caption")
                        If Not (captionAttr Is Nothing) Then
                            If Trim(captionAttr.value) <> "" Then
                                WriteJournal Trim(captionAttr.value), aTrnClass, cTRNNum
                                If childelement.SelectNodes("./row").length > 1 Then
                                    Dim aRow As IXMLDOMElement
                                    Dim aCol As IXMLDOMElement

                                    For i = 0 To childelement.SelectNodes("./row").length - 1
                                       Set aRow = childelement.SelectNodes("./row").item(i)
                                       Dim Row As String
                                       Row = ""
                                        For j = 0 To aRow.SelectNodes("./col").length - 1
                                            Row = Row & aRow.SelectNodes("./col").item(j).Text & "|"
                                        Next
                                        WriteJournal Row, aTrnClass, cTRNNum
                                    Next
                                    Set aRow = Nothing
                                    Set aCol = Nothing
                                End If
                            End If
                        End If
                    ElseIf childelement.baseName = "listbox" Then
                        Set captionAttr = childelement.Attributes.getNamedItem("caption")
                        If Not (captionAttr Is Nothing) Then
                            If Trim(captionAttr.value) <> "" Then
                                WriteJournal Trim(captionAttr.value), aTrnClass, cTRNNum
                                If childelement.SelectNodes("./item").length > 0 Then
                                    Dim aValueAttr As IXMLDOMAttribute
                                    Dim item As IXMLDOMElement
                                    For i = 0 To childelement.SelectNodes("./item").length - 1
                                        Set item = childelement.SelectNodes("./item").item(i)
                                        Set aValueAttr = item.Attributes.getNamedItem("value")
                                        WriteJournal aValueAttr.value, aTrnClass, cTRNNum
                                    Next
                                    Set item = Nothing
                                    Set aValueAttr = Nothing
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next child
    
    Set child = Nothing
    Set childelement = Nothing
    Set captionAttr = Nothing
    Set labelattr = Nothing
    Set captionElement = Nothing
    Set valueattr = Nothing

    Set TrnClassAttr = Nothing
    ProcessForm = ""
End Function

Public Function Parse(inElement As IXMLDOMElement) As String
    Dim childnode As IXMLDOMNode
    Dim childelement As IXMLDOMElement
    
    Dim trncounterattr As IXMLDOMAttribute
    Set trncounterattr = StaticLink.Attributes.getNamedItem("trncounter")
    If Not (trncounterattr Is Nothing) And trncounterattr.value = "before" Then UpdateTrnNum_
    
    If UCase(inElement.baseName) = "TRN" Or UCase(inElement.baseName) = "FORM" Or UCase(inElement.baseName) = "FORMUPDATE" Then
        Parse = ProcessForm(inElement)
    ElseIf UCase(inElement.baseName) = UCase("comarea") Then
        Parse = ProcessComArea(inElement)
    ElseIf UCase(inElement.baseName) = UCase("report") Then
        Parse = ProcessReport(inElement)
    ElseIf UCase(inElement.baseName) = UCase("call") Then
        Dim systemAttr As IXMLDOMAttribute
        Set systemAttr = inElement.Attributes.getNamedItem("system")
        If (Not systemAttr Is Nothing) And systemAttr.value = "IRIS" Then
            Parse = ProcessIRIS(inElement.selectSingleNode("./inputview"))
        End If
        Set systemAttr = Nothing
    ElseIf UCase(inElement.baseName) = UCase("method") Then
        Parse = ProcessComArea(inElement)
    
    Else
        For Each childnode In inElement.childNodes
            If childnode.nodeType = NODE_ELEMENT Then
                Set childelement = childnode
                If UCase(childelement.baseName) = "TRN" Or UCase(childelement.baseName) = "FORM" Then
                    Parse = Parse & ProcessForm(childelement)
                ElseIf childelement.baseName = "comarea" Then
                    Parse = Parse & ProcessComArea(childelement)
                End If
            End If
        Next childnode
    End If
    
    If Not (trncounterattr Is Nothing) And trncounterattr.value = "after" Then UpdateTrnNum_
    
    Set childnode = Nothing
    Set childelement = Nothing
    Set trncounterattr = Nothing
    
End Function

Public Function ProcessComArea(inElement As IXMLDOMElement)
    Dim aTrnClass As String
    Dim TrnClassAttr As IXMLDOMAttribute
    Dim valueattr As IXMLDOMAttribute
    Dim optionalattr As IXMLDOMAttribute
    
    Set TrnClassAttr = StaticLink.Attributes.getNamedItem("trnclass")
    
    If TrnClassAttr Is Nothing Then
        aTrnClass = cTRNCode
    ElseIf TrnClassAttr.value = "" Then
        MsgBox "Λάθος: Δεν περιλαμβάνεται κωδικός συναλλαγής στην δήλωση του ημερολογίου...", vbOKOnly, "Λάθος Δήλωσης(cXmlJournal.ProcessForm)"
        Exit Function
    Else
        aTrnClass = TrnClassAttr.value
    End If
    
     Dim i As Long, journalList As IXMLDOMNodeList, captionAttr As IXMLDOMAttribute
     Set journalList = inElement.SelectNodes(".//*[@caption!='']")
     If Not (journalList Is Nothing) Then
        For i = 0 To journalList.length - 1
           Set captionAttr = journalList.item(i).Attributes.getNamedItem("caption")
           Set valueattr = journalList.item(i).Attributes.getNamedItem("value")
           Set optionalattr = journalList.item(i).Attributes.getNamedItem("optional")
           
           If Not (optionalattr Is Nothing) And Trim(journalList.item(i).Text) = "" Then
           
           Else
                If Not (captionAttr Is Nothing) And Not (valueattr Is Nothing) Then
                    WriteJournal Trim(captionAttr.Text) & ": " & valueattr.Text, aTrnClass, cTRNNum
                End If
                If Not (captionAttr Is Nothing) And (valueattr Is Nothing) Then
                    WriteJournal Trim(captionAttr.Text) & ": " & journalList.item(i).Text, aTrnClass, cTRNNum
                End If
           End If
           
        Next
     End If
     ProcessComArea = ""

     Set captionAttr = Nothing
     Set journalList = Nothing
     Set optionalattr = Nothing
     Set valueattr = Nothing
     Set TrnClassAttr = Nothing

End Function

Public Function ProcessReport(inElement As IXMLDOMElement)
    Dim aTrnClass As String
    Dim TrnClassAttr As IXMLDOMAttribute
    Dim valueattr As IXMLDOMAttribute
    
    Set TrnClassAttr = StaticLink.Attributes.getNamedItem("trnclass")
    
    If TrnClassAttr Is Nothing Then
        aTrnClass = cTRNCode
    ElseIf TrnClassAttr.value = "" Then
        MsgBox "Λάθος: Δεν περιλαμβάνεται κωδικός συναλλαγής στην δήλωση του ημερολογίου...", vbOKOnly, "Λάθος Δήλωσης(cXmlJournal.ProcessForm)"
        Exit Function
    Else
        aTrnClass = TrnClassAttr.value
    End If
    
     Dim i As Long, journalList As IXMLDOMNodeList, captionAttr As IXMLDOMAttribute
     Set journalList = inElement.SelectNodes(".//*[@caption!='']")
     If Not (journalList Is Nothing) Then
        For i = 0 To journalList.length - 1
           Set captionAttr = journalList.item(i).Attributes.getNamedItem("caption")
           Set valueattr = journalList.item(i).Attributes.getNamedItem("value")
           If Not (captionAttr Is Nothing) And Not (valueattr Is Nothing) Then
              WriteJournal Trim(captionAttr.Text) & ": " & valueattr.Text, aTrnClass, cTRNNum
           End If
           If Not (captionAttr Is Nothing) And (valueattr Is Nothing) Then
              WriteJournal Trim(captionAttr.Text) & ": " & journalList.item(i).Text, aTrnClass, cTRNNum
           End If
        Next
     End If
     
     Set TrnClassAttr = Nothing
     Set valueattr = Nothing
     ProcessReport = ""
End Function

Public Function LoadXML(invalue As String)

    Dim workDocument As New MSXML2.DOMDocument30
    Dim elm As IXMLDOMElement, childElm As IXMLDOMElement
    
    workDocument.LoadXML invalue
    
    LoadXML = Parse(workDocument.documentElement)
    SaveJournal
    
    Set elm = Nothing
    Set childElm = Nothing
    Set workDocument = Nothing
                
End Function

Public Sub WriteJournal(Line As String, inTrnCode As String, inTRNNum As Integer)
    Dim astr As String, bstr As String, cmdStr As String, recno As Integer, TrnNum As Integer
    
    On Error GoTo error_handler
    astr = Line
    
    If (Trim(astr) = "") Then Exit Sub

    With GenWorkForm.vJournal
        If (LastTRNCode <> inTrnCode) Or (LastTRNNum <> inTRNNum) Then
            LastTRNCode = inTrnCode
            LastTRNNum = inTRNNum
            .SelStart = Len(GenWorkForm.vJournal.Text): .SelLength = 0: .SelBold = True
            .SelText = vbCrLf & vbCrLf & "Συναλλαγή: " & _
                LastTRNCode & "   A/A: " & CStr(LastTRNNum) & "   Χρήστης: " & cUserName
        End If
        .SelStart = Len(GenWorkForm.vJournal.Text): .SelLength = 0: .SelBold = False
        .SelText = vbCrLf & astr
    End With
    
    GoTo ExitPoint

error_handler:
    Call NBG_LOG_MsgBox("Error :" & error$, True)
ExitPoint:
End Sub

Public Function XML() As String
    If StaticLink Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(StaticLink.XML)
    End If
End Function

Public Function ProcessIRIS(inElement As IXMLDOMElement)
    Dim aTrnClass As String
    Dim TrnClassAttr As IXMLDOMAttribute
    Dim valueattr As IXMLDOMAttribute
    
    Set TrnClassAttr = StaticLink.Attributes.getNamedItem("trnclass")
    
    If TrnClassAttr Is Nothing Then
        aTrnClass = cTRNCode
    ElseIf TrnClassAttr.value = "" Then
        MsgBox "Λάθος: Δεν περιλαμβάνεται κωδικός συναλλαγής στην δήλωση του ημερολογίου...", vbOKOnly, "Λάθος Δήλωσης(cXmlJournal.ProcessForm)"
        Exit Function
    Else
        aTrnClass = TrnClassAttr.value
    End If
    
     Dim i As Long, journalList As IXMLDOMNodeList, captionAttr As IXMLDOMAttribute
     Set journalList = inElement.SelectNodes(".//*[@caption!='']")
     If Not (journalList Is Nothing) Then
        For i = 0 To journalList.length - 1
           Set captionAttr = journalList.item(i).Attributes.getNamedItem("caption")
           Set valueattr = journalList.item(i).Attributes.getNamedItem("value")
           If Not (captionAttr Is Nothing) And Not (valueattr Is Nothing) Then
              WriteJournal Trim(captionAttr.Text) & ": " & valueattr.Text, aTrnClass, cTRNNum
           End If
           If Not (captionAttr Is Nothing) And (valueattr Is Nothing) Then
              WriteJournal Trim(captionAttr.Text) & ": " & journalList.item(i).Text, aTrnClass, cTRNNum
           End If
        Next
     End If
     
     Set TrnClassAttr = Nothing
     Set valueattr = Nothing
     ProcessIRIS = ""
End Function

