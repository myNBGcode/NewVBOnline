VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSWIFTmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private avalidator
Public messagedoc As New MSXML2.DOMDocument30
Public ResultCode As Integer
Public ResultMessage As String
Private Sub CopyAttr(aattr As IXMLDOMAttribute, elm As IXMLDOMElement)
    If aattr Is Nothing Then
    Else
        Dim newAttr As IXMLDOMAttribute
        Set newAttr = elm.ownerDocument.createAttribute(aattr.baseName)
        newAttr.value = aattr.value
        elm.Attributes.setNamedItem newAttr
    End If
End Sub

Public Sub prepare(messagename As String)
    Dim adoc 'As New MSXML2.DOMDocument30
    Set adoc = New MSXML2.DOMDocument30 'CreateObject("Msxml2.DOMDocument.6.0")
    adoc.Load ReadDir & "\onlinemessages\swift.xml"
    
    Dim elm As IXMLDOMElement
    Set elm = adoc.selectSingleNode("//transactions/trn[@name='" & messagename & "']")
    If elm Is Nothing Then
        Err.Raise 1000, "cSWIFTmessage", "Δεν βρέθηκε το μήνυμα: " & messagename
    Else
        messagedoc.LoadXML elm.XML
        For Each elm In messagedoc.SelectNodes("//trn/field")
            Dim aaattr As IXMLDOMAttribute
            Set aaattr = elm.Attributes.getNamedItem("aa")
            If aaattr Is Nothing Then
            
            Else
                Dim fieldelm As IXMLDOMElement
                Set fieldelm = adoc.selectSingleNode("//fields/field[@aa='" & aaattr.value & "']")
                If fieldelm Is Nothing Then
                    Err.Raise 1001, "cSWIFTmessage", "Δεν βρέθηκε το πεδίο " & aaattr.value & " στο μήνυμα " & messagename
                Else
                    Dim descriptionAttr As IXMLDOMAttribute
                    Dim nameattr As IXMLDOMAttribute
                    Dim labelattr As IXMLDOMAttribute
                    Dim validationattr As IXMLDOMAttribute
                    
                    
                    Dim newAttr As IXMLDOMAttribute
                    
                    Set descriptionAttr = fieldelm.Attributes.getNamedItem("description")
                    Set nameattr = fieldelm.Attributes.getNamedItem("name")
                    Set labelattr = fieldelm.Attributes.getNamedItem("label")
                    Set validationattr = fieldelm.Attributes.getNamedItem("validation")
                    
                    
                    
                    CopyAttr descriptionAttr, elm
                    CopyAttr nameattr, elm
                    CopyAttr labelattr, elm
                    CopyAttr validationattr, elm
                    
                    
                    If Not (validationattr Is Nothing) Then
                        Dim aregexpelm As IXMLDOMElement
                        Set aregexpelm = adoc.selectSingleNode("//validationtemplates/template[@name='" & validationattr.value & "']")
                        If aregexpelm Is Nothing Then
                            Err.Raise 1002, "cSWIFTmessage", "Δεν βρέθηκε η μεταφραση για τον τύπο " & validationattr.value & " στο μήνυμα " & messagename
                        Else
                            Dim aregexpattr As IXMLDOMAttribute
                            Dim avalidformatattr As IXMLDOMAttribute
                            Dim aformatattr As IXMLDOMAttribute
                            
                            Set aregexpattr = aregexpelm.Attributes.getNamedItem("regexp")
                            Set avalidformatattr = aregexpelm.Attributes.getNamedItem("validformat")
                            Set aformatattr = aregexpelm.Attributes.getNamedItem("format")
                            If aregexpattr Is Nothing Then
                                Err.Raise 1002, "cSWIFTmessage", "Δεν έχει δηλωθεί η μεταφραση στον τύπο " & validationattr.value & " στο μήνυμα " & messagename
                            Else
                                CopyAttr aregexpattr, elm
                            End If
                            
                            CopyAttr avalidformatattr, elm
                            CopyAttr aformatattr, elm
                        End If
                    End If
                    
                End If
            End If
        Next elm
        
    End If
    
    Set avalidator = CreateObject("MSScriptControl.ScriptControl.1")
    
    'set avalidator =ValidationControl.
    Dim aregcode As String
    aregcode = _
        "function MatchRegExpr( teststring , regexpstring )" & vbCrLf & _
        "    stop" & vbCrLf & _
        "    MatchRegExpr = False" & vbCrLf & _
        "    Dim regEx, Match, Matches" & vbCrLf & _
        "    Set regEx = New RegExp" & vbCrLf & _
        "    regEx.Pattern = regexpstring" & vbCrLf & _
        "    regEx.IgnoreCase = True" & vbCrLf & _
        "    regEx.Global = True" & vbCrLf & _
        "    Set Matches = regEx.Execute(teststring)" & vbCrLf & _
        "    'If Matches.count <> 1 Then Exit Function" & vbCrLf & _
        "    If Matches.count <> 1 Then " & vbCrLf & _
        "    For Each Match In Matches" & vbCrLf & _
        "        If CStr(Match.value) <> teststring Then " & vbCrLf & _
        "        MsgBox Match.FirstIndex & "","" & Match.length & "", "" & CStr(Match.value)" & vbCrLf & _
        "        Exit Function" & vbCrLf & _
        "        Exit For" & vbCrLf & _
        "        end if" & vbCrLf & _
        "    Next" & vbCrLf & _
        "    end if " & vbCrLf & _
        "    MatchRegExpr = True" & vbCrLf & _
        "End Function"
   
    aregcode = _
    "function MatchRegExpr( teststring , regexpstring )" & vbCrLf & _
    "MatchRegExpr = False" & vbCrLf & "Dim regEx, Match, Matches" & vbCrLf & _
    "Set regEx = New RegExp" & vbCrLf & _
    "regEx.Pattern = regexpstring" & vbCrLf & _
    "regEx.IgnoreCase = True" & vbCrLf & _
    "regEx.Global = True" & vbCrLf & _
    "Set Matches = regEx.Execute(teststring)" & vbCrLf & _
    "'If Matches.count <> 1 Then Exit Function" & vbCrLf & _
    "If Matches.count <> 1 Then" & vbCrLf & _
    "   If len(regexpstring) > 1 Then" & vbCrLf & _
    "      If Right(regexpstring, 1) = ""$"" Then" & vbCrLf & _
    "         regexpstring = Left(regexpstring, Len(regexpstring) - 1)" & vbCrLf & "         Set regEx = New RegExp" & vbCrLf & _
    "         regEx.Pattern = regexpstring" & vbCrLf & "         regEx.IgnoreCase = True" & vbCrLf & _
    "         regEx.Global = True" & vbCrLf & "         Set Matches = regEx.Execute(teststring)" & vbCrLf & _
    "         For Each Match In Matches" & vbCrLf & _
    "             If Match.FirstIndex = 0 Then" & vbCrLf & _
    "                MsgBox ""Μή αποδεκτοί χαρακτήρες: "" & Right(teststring, Len(teststring) - Match.length)" & vbCrLf & "                Exit Function" & vbCrLf & _
    "             End If" & vbCrLf & "             Exit Function" & vbCrLf & "         Next" & vbCrLf & _
    "      End If" & vbCrLf & "    End If" & vbCrLf & "    Exit Function" & vbCrLf & _
    "End If" & vbCrLf & "MatchRegExpr = True" & vbCrLf & "End Function"

   avalidator.TimeOut = -1
   
   avalidator.language = "vbscript"
   avalidator.AddCode aregcode

End Sub


Public Property Get value(fieldaa As String)
    Dim aelm As IXMLDOMElement
    
    Set aelm = messagedoc.selectSingleNode("//field[@aa='" + fieldaa + "']")
    
    If aelm Is Nothing Then
        Err.Raise 1002, "cSWIFTmessage", "Δεν υπάρχει το πεδίο " & fieldaa & " στο μήνυμα "
    Else
    
    End If
    
End Property


Public Property Let value(fieldaa As String, aValue)
    
    Dim aelm As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    ResultCode = 0: ResultMessage = ""
    'MsgBox "ok3"
    Set aelm = messagedoc.selectSingleNode("//field[@aa='" + fieldaa + "']")
    
    If aelm Is Nothing Then
        Err.Raise 1002, "cSWIFTmessage", "Δεν υπάρχει το πεδίο " & fieldaa & " στο μήνυμα "
    Else
        Set aattr = aelm.Attributes.getNamedItem("regexp")
        If Not (aattr Is Nothing) Then
            Dim aregexp As String
            aregexp = aattr.value
            Dim res As Boolean
            res = avalidator.Run("MatchRegExpr", CStr(aValue), aregexp)
            If Not res Then
                Dim adescr As String, amessage As String
                
                Set aattr = aelm.Attributes.getNamedItem("description")
                If aattr Is Nothing Then adescr = fieldaa Else adescr = aattr.value
                'amessage = "Λάθος τιμή για το πεδίο " & adescr & " στο μήνυμα SWIFT."
                Set aattr = aelm.Attributes.getNamedItem("validformat")
                If Not (aattr Is Nothing) Then amessage = "Αποδεκτή μορφή τιμής:" & aattr.value
               ' MsgBox amessage, vbCritical, "Ελεγχος Τιμών SWIFT"
                ResultCode = 1: ResultMessage = amessage
                Exit Property
            End If
        End If
        Set aattr = aelm.Attributes.getNamedItem("value")
        If aattr Is Nothing Then
            Set aattr = messagedoc.createAttribute("value")
            aelm.Attributes.setNamedItem aattr
        End If
        
        aattr.value = CStr(aValue)
    End If
    'MsgBox fieldaa & "-" & avalue
End Property


Public Property Get FormatedText() As String
    Dim astr As String
    astr = ""
    Dim elm As IXMLDOMElement
    For Each elm In messagedoc.SelectNodes("//field")
        Dim valueattr As IXMLDOMAttribute
        Dim aaattr As IXMLDOMAttribute
        Dim optionattr As IXMLDOMAttribute
        Dim formatattr As IXMLDOMAttribute
        
        Set valueattr = elm.Attributes.getNamedItem("value")
        Set aaattr = elm.Attributes.getNamedItem("aa")
        Set optionattr = elm.Attributes.getNamedItem("option")
        Set formatattr = elm.Attributes.getNamedItem("format")
        
        If Not aaattr Is Nothing Then
            If Not (optionattr Is Nothing) Then
                If optionattr.value = "M" Then
                    If valueattr Is Nothing Then
                        LogMsgbox "Υποχρεωτικό το πεδίο " & aaattr.value, vbCritical, "Λάθος στη Σύνθεση του Μηνύματος SWIFT"
                        Exit Function
                    Else
                        If Trim(valueattr.value) = "" Then
                            LogMsgbox "Υποχρεωτικό το πεδίο " & aaattr.value, vbCritical, "Λάθος στη Σύνθεση του Μηνύματος SWIFT"
                            Exit Function
                        End If
                    End If
                End If
            End If
            If Not valueattr Is Nothing Then
                If Trim(valueattr.value) <> "" Then
                    If Not (formatattr Is Nothing) Then
                       If Trim(formatattr.value) <> "" Then
                          Dim cformat As New cSwiftFormat
                          cformat.invalue = valueattr.value
                          cformat.InFormat = formatattr.value
                          astr = astr & ":" & aaattr.value & ":" & cformat.FormatedValue & vbCrLf
                       Else
                           astr = astr & ":" & aaattr.value & ":" & valueattr.value & vbCrLf
                       End If
                    Else
                       astr = astr & ":" & aaattr.value & ":" & valueattr.value & vbCrLf
                    End If
                End If
            End If
        End If
    Next elm
    FormatedText = astr
   ' MsgBox astr
End Property

Public Property Get PrintSwiftMessage()
    Dim astr As String
    Dim ReturnArray() As String
    Dim Index As Integer
    Index = -1
    Dim elm As IXMLDOMElement
    For Each elm In messagedoc.SelectNodes("//field")
        astr = ""
        Dim valueattr As IXMLDOMAttribute
        Dim aaattr As IXMLDOMAttribute
        Dim optionattr As IXMLDOMAttribute
        Dim formatattr As IXMLDOMAttribute
        Dim descriptionAttr As IXMLDOMAttribute
        
        Set valueattr = elm.Attributes.getNamedItem("value")
        Set aaattr = elm.Attributes.getNamedItem("aa")
        Set optionattr = elm.Attributes.getNamedItem("option")
        Set formatattr = elm.Attributes.getNamedItem("format")
        Set descriptionAttr = elm.Attributes.getNamedItem("description")
        
        If Not aaattr Is Nothing Then
            If Not valueattr Is Nothing And Not descriptionAttr Is Nothing Then
                If Trim(valueattr.value) <> "" Then
                    If Not (formatattr Is Nothing) Then
                       If Trim(formatattr.value) <> "" Then
                          Dim cformat As New cSwiftFormat
                          cformat.invalue = valueattr.value
                          cformat.InFormat = formatattr.value
                          'astr = ":" & aaattr.value & ":" & cformat.FormatedValue
                          astr = descriptionAttr.value & ":" & aaattr.value & ":" & cformat.FormatedValue
                       Else
                           'astr = ":" & aaattr.value & ":" & valueattr.value
                           astr = descriptionAttr.value & ":" & aaattr.value & ":" & valueattr.value
                       End If
                    Else
                       astr = descriptionAttr.value & ":" & aaattr.value & ":" & valueattr.value
                    End If
                End If
            End If
        End If
        Dim aPos As Integer
        Dim astart As Integer
        Dim i As Integer
        aPos = 0: astart = 1
        For i = astart To Len(astr)
            Index = Index + 1
            aPos = InStr(astart, astr, vbCrLf)
            ReDim Preserve ReturnArray(Index)
            If aPos = 0 Then
               ReturnArray(Index) = Mid(astr, astart)
               Exit For
            Else
               ReturnArray(Index) = Mid(astr, astart, aPos - 1)
               astr = Mid(astr, aPos + 2)
            End If
        Next
    Next elm
    PrintSwiftMessage = ReturnArray
End Property

Public Property Get FormatedValue(fieldaa As String)
    Dim aelm As IXMLDOMElement
    Dim valueattr As IXMLDOMAttribute
    Dim formatattr As IXMLDOMAttribute
    
    Dim astr As String
    Set aelm = messagedoc.selectSingleNode("//field[@aa='" + fieldaa + "']")
    
    If aelm Is Nothing Then
        Err.Raise 1002, "cSWIFTmessage", "Δεν υπάρχει το πεδίο " & fieldaa & " στο μήνυμα "
    Else
       Set valueattr = aelm.Attributes.getNamedItem("value")
       Set formatattr = aelm.Attributes.getNamedItem("format")
       
       If Not (valueattr Is Nothing) Then
          If Not (formatattr Is Nothing) Then
             Dim cformat As New cSwiftFormat
             cformat.invalue = valueattr.value
             cformat.InFormat = formatattr.value
             FormatedValue = cformat.FormatedValue
          Else
             FormatedValue = valueattr.value
          End If
       Else
          FormatedValue = ""
       End If
    End If
End Property





