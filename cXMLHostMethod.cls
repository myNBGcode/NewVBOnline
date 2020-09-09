VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLHostMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public code As String
Public fakeResult As String
Private messagestruct As MSXML2.DOMDocument30

Private Function processmessage(messagedoc As IXMLDOMElement) As String
    Dim fieldlist As IXMLDOMNodeList
    Dim messagestring
    Set fieldlist = messagedoc.SelectNodes(".//field")
    If fieldlist Is Nothing Then
    Else
        If fieldlist.length = 0 Then
        
        Else
            Dim Node As IXMLDOMNode
            messagestring = ""
            Dim lastseqno As Integer
            Dim aaattr As IXMLDOMAttribute
            Dim descriptionAttr As IXMLDOMAttribute
            
            Dim seqattr As IXMLDOMAttribute
            Dim lengthattr As IXMLDOMAttribute
            Dim leftpadattr As IXMLDOMAttribute
            Dim rightpadattr As IXMLDOMAttribute
            Dim optionalattr As IXMLDOMAttribute
            Dim valueattr As IXMLDOMAttribute
            Dim headerattr As IXMLDOMAttribute
            Dim journalattr As IXMLDOMAttribute
            
            Dim aa As Integer, description As String
            Dim length As Integer, leftpadchar As String, rightpadchar As String, OptionalFlag  As Boolean, value As String
            Dim journalflag As Boolean
            
            lastseqno = 0
            For Each Node In fieldlist
                If Node.nodeType = NODE_ELEMENT Then
                    Set aaattr = Node.Attributes.getNamedItem("aa")
                    Set descriptionAttr = Node.Attributes.getNamedItem("description")
                    Set seqattr = Node.Attributes.getNamedItem("seq")
                    Set lengthattr = Node.Attributes.getNamedItem("length")
                    Set leftpadattr = Node.Attributes.getNamedItem("leftpad")
                    Set rightpadattr = Node.Attributes.getNamedItem("rightpad")
                    Set optionalattr = Node.Attributes.getNamedItem("optional")
                    Set valueattr = Node.Attributes.getNamedItem("value")
                    Set headerattr = Node.Attributes.getNamedItem("header")
                    Set journalattr = Node.Attributes.getNamedItem("writejournal")
                    
                    
                    aa = 0
                    If Not (aaattr Is Nothing) Then
                        On Error Resume Next
                        aa = CInt(aaattr.value)
                    End If
                    description = ""
                    If Not (descriptionAttr Is Nothing) Then
                        description = descriptionAttr.value
                    End If
                                        
                    If seqattr Is Nothing Then GoTo errMissingSeqAttr
                    On Error GoTo errMissingSeqAttr
                    If lastseqno > CInt(seqattr.value) Then GoTo errMissingSeqAttr
                    lastseqno = CInt(seqattr.value)
                    If lengthattr Is Nothing Then GoTo errMissingLengthAttr
                    On Error GoTo errMissingLengthAttr
                    length = CInt(lengthattr.value)
                    On Error GoTo errMissingPadInfo
                    leftpadchar = ""
                    rightpadchar = ""
                    If leftpadattr Is Nothing Then
                        If rightpadattr Is Nothing Then GoTo errMissingPadInfo
                        rightpadchar = rightpadattr.value
                        If rightpadchar = "" Then GoTo errMissingPadInfo
                        leftpadchar = ""
                    Else
                        leftpadchar = leftpadattr.value
                        If leftpadchar = "" Then GoTo errMissingPadInfo
                        rightpadchar = ""
                    End If
                    On Error GoTo errGenProcess
                    OptionalFlag = False
                    If Not (optionalattr Is Nothing) Then
                        If LCase(optionalattr.value) = "yes" Then OptionalFlag = True
                    End If
                    If Not (valueattr Is Nothing) Then
                        value = Trim(valueattr.value)
                    Else
                        value = Trim(Node.Text)
                    End If
                    'todo Additional formating, clear instructions ( remove leading 0, etc )
                    
                    Dim MessagePart As String, JournalValue As String
                    If OptionalFlag = True And Trim(value) = "" Then
                        MessagePart = "": JournalValue = ""
                    Else
                        If leftpadchar <> "" Then MessagePart = Right(String(length, leftpadchar) & value, length)
                        If rightpadchar <> "" Then MessagePart = Left(value & String(length, rightpadchar), length)
                        If Not (headerattr Is Nothing) Then MessagePart = headerattr.value & MessagePart
                        JournalValue = value
                    End If
                    journalflag = True
                    If Not (journalattr Is Nothing) Then
                        If journalattr.value = "false" Then journalflag = False
                    End If
                    If journalflag Then
                        If JournalValue <> "" Then
                            eJournalWriteFld Nothing, aa, description, JournalValue
                        End If
                    End If
                    
                    messagestring = messagestring & MessagePart
                    
                End If
            Next Node
        
        End If
        processmessage = messagestring
    
    End If
    Exit Function
errMissingSeqAttr:
    LogMsgbox "Λάθος Παράμετρος σειράς στην περιγραφή μηνύματος..." & Err.number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
errMissingLengthAttr:
    LogMsgbox "Λάθος Παράμετρος μήκους στην περιγραφή μηνύματος..." & Err.number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
errMissingPadInfo:
    LogMsgbox "Λάθος Παράμετρος συμπλήρωσης στην περιγραφή μηνύματος..." & Err.number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
errGenProcess:
    LogMsgbox "Λάθος στην επεξεργασία της περιγραφής μηνύματος..." & Err.number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
End Function

Private Function processheader(messagedoc As IXMLDOMElement) As String
    Dim headerattr As IXMLDOMAttribute
    Dim keyattr As IXMLDOMAttribute
    Set headerattr = messagedoc.Attributes.getNamedItem("header")
    Set keyattr = messagedoc.Attributes.getNamedItem("key")
    
    Dim headerstring As String
    headerstring = String(4, " ")
    If Not (headerattr Is Nothing) Then headerstring = Left(headerattr.value & headerstring, 4)
    headerstring = Left(headerstring & cHEAD & String(4, " "), 8)
    If keyattr Is Nothing Then
        headerstring = headerstring & " "
    Else
        headerstring = Left(headerstring + keyattr.value + " ", 9)
    End If
    UpdateTrnNum_
    headerstring = headerstring & Right("000" & cTRNNum, 3)
    processheader = headerstring
End Function

'Private Function processlink(messagedoc As IXMLDOMElement) As Boolean
'    processlink = False
'    Dim linkattr As IXMLDOMAttribute
'    Set linkattr = messagedoc.Attributes.getNamedItem("UseSNAPoolLink")
'    If Not (linkattr Is Nothing) Then
'        If linkattr.value = "1" Then processlink = True
'    End If
'End Function
Private Function processmod28(messagedoc As IXMLDOMElement) As Boolean
    processmod28 = False
    Dim modattr As IXMLDOMAttribute
    Set modattr = messagedoc.Attributes.getNamedItem("Module28PoolLink")
    If Not (modattr Is Nothing) Then
        If modattr.value = "1" Then processmod28 = True
    End If
End Function

Private Function PrepareDoc(Data As String) As MSXML2.DOMDocument30
    Dim codedoc As New MSXML2.DOMDocument30
    
    codedoc.LoadXML code
    
    Dim datadoc As New MSXML2.DOMDocument30
    datadoc.LoadXML Data
    
    Dim fieldNode As IXMLDOMElement
    For Each fieldNode In datadoc.SelectNodes("//field")
        Dim nameattr As IXMLDOMAttribute
        Set nameattr = fieldNode.Attributes.getNamedItem("name")
        If nameattr Is Nothing Then
        Else
            Dim codefield As IXMLDOMElement
            Set codefield = codedoc.selectSingleNode("//field[@name='" & nameattr.value & "']")
            If codefield Is Nothing Then
            Else
                codefield.Text = fieldNode.Text
            End If
        End If
    Next fieldNode
    Set PrepareDoc = codedoc
End Function

Public Function DetectResultGroup(tokenizer As CXmlHostMethodResultTokenizer, groupelm As IXMLDOMNode) As Boolean
    Dim detectelm As IXMLDOMElement
    Set detectelm = groupelm.selectSingleNode("./detection")
    If detectelm Is Nothing Then
        DetectResultGroup = True: Exit Function
    Else
        Dim patternelm As IXMLDOMElement
        Set patternelm = detectelm.selectSingleNode("./pattern")
        If patternelm Is Nothing Then
            DetectResultGroup = False: Exit Function
        Else
            Dim fieldelm As IXMLDOMElement
            Dim totallength As Integer, length As Integer, value As String, testvalue As String
            totallength = 0
            For Each fieldelm In patternelm.childNodes
                Dim lengthattr As IXMLDOMAttribute
                Set lengthattr = fieldelm.Attributes.getNamedItem("length")
                Dim valueattr As IXMLDOMAttribute
                Set valueattr = fieldelm.Attributes.getNamedItem("value")

                If lengthattr Is Nothing Then
                    If valueattr Is Nothing Then
                        DetectResultGroup = False: Exit Function
                    Else
                        length = Len(valueattr.value)
                        value = valueattr.value
                    End If
                Else
                    If lengthattr.value = "unbounded" Then
                        length = Len(tokenizer.TestAll)
                    Else
                        length = lengthattr.value
                    End If
                    If valueattr Is Nothing Then
                        value = ""
                    Else
                        value = valueattr.value
                    End If
                End If
                
                totallength = totallength + length
                If value <> "" Then
                    testvalue = tokenizer.TestToken(totallength)
                    If Len(testvalue) <> totallength Then
                        DetectResultGroup = False: Exit Function
                    Else
                        If Right(testvalue, length) <> value Then
                            DetectResultGroup = False: Exit Function
                        Else
                            'do nothing, process next field
                        End If
                    End If
                Else
                    'do nothing, process next field
                End If
            Next fieldelm
        End If
    End If
    DetectResultGroup = True: Exit Function
    
End Function

Public Function ProcessField(tokenizer As CXmlHostMethodResultTokenizer, fieldNode As IXMLDOMNode, resultelement As IXMLDOMElement) As Boolean
    Dim lengthattr As IXMLDOMAttribute, nameattr As IXMLDOMAttribute
    Set lengthattr = fieldNode.Attributes.getNamedItem("length")
    Set nameattr = fieldNode.Attributes.getNamedItem("name")
    If Not (lengthattr Is Nothing) Then
        Dim length As Integer
        Dim value As String
        
        If lengthattr.value = "unbounded" Then
            value = tokenizer.GetAll
        Else
            length = lengthattr.value
            value = tokenizer.GetToken(length)
        End If
        Dim fieldelement As IXMLDOMElement
        Set fieldelement = resultelement.ownerDocument.createElement("field")
        If nameattr Is Nothing Then
        Else
            'fieldelement.Attributes.setNamedItem nameattr
            
            Dim copynameattr As IXMLDOMAttribute
            Set copynameattr = fieldelement.ownerDocument.createAttribute("name")
            copynameattr.value = nameattr.value
            fieldelement.Attributes.setNamedItem copynameattr
            fieldelement.Text = value
        End If
        resultelement.appendChild fieldelement
        Dim valueattr As IXMLDOMAttribute
        'Set valueattr = fieldelement.ownerDocument.createAttribute("value")
        'valueattr.Value = Value
        'fieldelement.Attributes.setNamedItem valueattr
        
'        fieldelement.Text = Value
    Else
    
    End If
    ProcessField = True: Exit Function
End Function

Public Function ProcessGroup(tokenizer As CXmlHostMethodResultTokenizer, GroupNode As IXMLDOMNode, resultelement As IXMLDOMElement) As Boolean
    Dim minattr As IXMLDOMAttribute, maxattr As IXMLDOMAttribute
    Set minattr = GroupNode.Attributes.getNamedItem("min")
    Set maxattr = GroupNode.Attributes.getNamedItem("max")
    
    Dim min As Long, max As Long, Counter As Long, unbounded As Boolean
    If minattr Is Nothing Then
        min = 1
    Else
        min = minattr.value
    End If
    If maxattr Is Nothing Then
        max = min: unbounded = False
    Else
        If maxattr.value = "unbounded" Then
            unbounded = True
        Else
            max = maxattr.value: unbounded = False
        End If
    End If
    
    Counter = 0
    Do
        If DetectResultGroup(tokenizer, GroupNode) Then
            Dim groupelement As IXMLDOMElement
            Set groupelement = resultelement.ownerDocument.createElement("group")
            Dim nameattr As IXMLDOMAttribute
            Set nameattr = GroupNode.Attributes.getNamedItem("name")
            If nameattr Is Nothing Then
            Else
                Dim groupnameattr As IXMLDOMAttribute
                Set groupnameattr = resultelement.ownerDocument.createAttribute("name")
                groupelement.Attributes.setNamedItem groupnameattr
                groupnameattr.value = nameattr.value
            End If
            
            resultelement.appendChild groupelement
            Dim Node As IXMLDOMNode
            For Each Node In GroupNode.childNodes
                If Node.nodeType = NODE_ELEMENT Then
                    If Node.baseName = "group" Then
                        ProcessGroup tokenizer, Node, groupelement
                    ElseIf Node.baseName = "field" Then
                        ProcessField tokenizer, Node, groupelement
                    Else
                    
                    End If
                End If
            Next Node
        
            Counter = Counter + 1
            If Not unbounded Then
                If Counter >= max Then Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    If Counter < min Then
        ProcessGroup = False
    Else
        ProcessGroup = True
    End If
End Function

Public Function ParseComResult(tokenizer As CXmlHostMethodResultTokenizer, resultelement As IXMLDOMElement) As Boolean
    Dim LineStr As String
    Dim i As Integer
    Dim msgdoc As New MSXML2.DOMDocument30
    msgdoc.LoadXML code
    
    Dim msgelm As IXMLDOMElement
    Set msgelm = msgdoc.selectSingleNode(".//lastmessage")
    
    Dim elmchild As IXMLDOMElement
    Dim astr As String
    Dim curline As Integer
    curline = 0
    
    For Each elmchild In msgelm.childNodes
        If elmchild.baseName = "group" Then
            ProcessGroup tokenizer, elmchild, resultelement
        ElseIf elmchild.baseName = "field" Then
            ProcessField tokenizer, elmchild, resultelement
        Else
        
        End If
    Next elmchild
End Function

Public Function LoadXML(astr As String) As String
    Dim resulttree As New MSXML2.DOMDocument30
    resulttree.appendChild resulttree.createElement("lastmessage")
    Dim ErrElm As IXMLDOMElement
    
    UpdateTrnNum_
    Set messagestruct = PrepareDoc(astr)
    If messagestruct Is Nothing Then
        Set ErrElm = resulttree.createElement("error")
        resulttree.documentElement.appendChild ErrElm
        ErrElm.Text = "Απέτυχε η προετοιμασία μηνύματος."
        
    Else
        Dim trncodeattr As IXMLDOMAttribute
        Dim initialelm As IXMLDOMElement
        Set initialelm = messagestruct.documentElement.selectSingleNode("//hostmethod/initialmessage")
        If Not initialelm Is Nothing Then
            Set trncodeattr = initialelm.Attributes.getNamedItem("header")
        End If
        If messagestruct.documentElement Is Nothing Then
            Set ErrElm = resulttree.createElement("error")
            resulttree.documentElement.appendChild ErrElm
            ErrElm.Text = "Απέτυχε η προετοιμασία μηνύματος."
        Else
            
            Dim OutputStr As String, headerstr As String
            Dim module28aflag As Boolean
            OutputStr = processmessage(messagestruct.documentElement.selectSingleNode(".//initialmessage"))
            headerstr = processheader(messagestruct.documentElement.selectSingleNode(".//initialmessage"))
            module28aflag = processmod28(messagestruct.documentElement.selectSingleNode(".//initialmessage"))
            Dim encodegreek As Boolean
            Dim encodeattr As IXMLDOMAttribute
            Set encodeattr = messagestruct.documentElement.Attributes.getNamedItem("encodegreek")
            If encodeattr Is Nothing Then
                encodegreek = False
            Else
                If encodeattr.value = "yes" Then encodegreek = True Else encodegreek = False
            End If
            Dim res As Long, resultList As Collection
            Dim StartTime
            StartTime = Time
                
            cb.send_str = headerstr & OutputStr
            res = SNAPool_Communicate(module28aflag)
            Set resultList = New Collection
            
            Dim k As Integer
            For k = 1 To ReceivedData.count
                If (Trim(ReceivedData(k)) <> "") Then
                    resultList.add ReceivedData(k) & "`"
                End If
            Next k

            Dim EndTime
            EndTime = Time
            If fakeResult <> "" Then
                While resultList.count > 0
                    resultList.Remove 1
                Wend
                resultList.add fakeResult
            End If
            If res = COM_OK Or (Mid(cb.receive_str, 1, 1) = "4") Then
                On Error GoTo 0
                Dim tokenizer As New CXmlHostMethodResultTokenizer
                tokenizer.prepare resultList
                ParseComResult tokenizer, resulttree.documentElement
            Else
                Set ErrElm = resulttree.createElement("comerror")
                resulttree.documentElement.appendChild ErrElm
                ErrElm.Text = "Απέτυχε η επικοινωνία: " & res
            End If
        End If
    End If
    LoadXML = resulttree.XML

End Function

Public Function XML() As String
    If messagestruct Is Nothing Then
        XML = ""
    Else
        If messagestruct.documentElement Is Nothing Then
            XML = ""
        Else
            XML = messagestruct.XML
        End If
    End If
End Function
