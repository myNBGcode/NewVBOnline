VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlComArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Container As Buffers
Private onLineComResult As cSNAResult

Public content As IXMLDOMElement
Public workDocument As MSXML2.DOMDocument30
Dim trncounter As Boolean
Dim c4EyesHandler As c4EyesComareaHandler

Public MessageID As String

'definition
'<comarea name="test" id="comareaid" filename="comareafilename">
'   <module name="ZADBX43" trncall="ACBTO" inputname="" outputname="">
'</comarea>
'
'call
'<input>
'   <structurename1>
'       <structurename2>
'           <fieldname>fieldvalue</fieldname>
'      </structurename2>
'   </structurename1>
'</input>
'<output name="test">
'</output>

Public Property Get ComResult() As cSNAResult
    Set ComResult = onLineComResult
End Property

Public Function HandleResp(resultdoc As MSXML2.DOMDocument30, Optional Title As String) As Boolean
    HandleResp = False
    
    If IsMissing(Title) Then Title = "Πρόβλημα Εφαρμογής..."
    
    Dim respnode As IXMLDOMElement
    Dim rcnode As IXMLDOMElement
    Dim msgnode As IXMLDOMElement
    Dim txtnode As IXMLDOMElement
    Dim pgmnode As IXMLDOMElement
    
    Set respnode = GetXmlNode(resultdoc.documentElement, "//RESP", "RESP", , "Πρόβλημα στο HandleResp...")
    If respnode Is Nothing Then Exit Function
    Set rcnode = GetXmlNode(resultdoc.documentElement, "//RC", "RC", "RESP", "Πρόβλημα στο HandleResp...")
    If rcnode Is Nothing Then Exit Function
    Set msgnode = GetXmlNode(resultdoc.documentElement, "//RC_MSG", "RC_MSG", "RESP", "Πρόβλημα στο HandleResp...")
    If msgnode Is Nothing Then Exit Function
    Set txtnode = GetXmlNode(resultdoc.documentElement, "//RC_TXT", "RC_TXT", "RESP", "Πρόβλημα στο HandleResp...")
    If txtnode Is Nothing Then Exit Function
    Set pgmnode = GetXmlNode(resultdoc.documentElement, "//RC_PGM", "RC_PGM", "RESP", "Πρόβλημα στο HandleResp...")
    If pgmnode Is Nothing Then Exit Function
    
    If Trim(rcnode.Text) <> "0" Then
        Dim msg As String
        msg = "Λάθος:" & rcnode.Text & " (" & msgnode.Text & " στο " & pgmnode.Text & ") " & txtnode.Text
        If ActiveL2TrnHandler Is Nothing Then
            eJournalWriteAll Nothing, CStr(msg)
        Else
            eJournalWriteAll ActiveL2TrnHandler.activeform, CStr(msg)
        End If
        NBG_MsgBox msg, True, Title
        Exit Function
    End If
    
    HandleResp = True

End Function

Private Function UpdateFields(inViewRoot, inDocRoot As IXMLDOMElement, ByVal aPos As Long) As Boolean
UpdateFields = False
Dim elm As IXMLDOMElement, selm As IXMLDOMElement, scount As Integer
Dim i As Long
    If inViewRoot Is Nothing Then Exit Function
    If inDocRoot Is Nothing Then UpdateFields = True: Exit Function
    If inDocRoot.childNodes.length > 0 Then
        If inDocRoot.firstChild.nodeTypeString <> "text" Then
        For Each elm In inDocRoot.childNodes
            'Ελεγχος αν υπάρχουν περισσότερα απο ενα buffer parts με το ιδιο ονομα
            scount = 1:
            Set selm = elm.previousSibling
            Do While Not (selm Is Nothing)
                If selm.baseName = elm.baseName Then
                    scount = scount + 1
                    Set selm = selm.previousSibling 'εδω
                Else
                    Exit Do
                End If
            Loop
            Dim apart
            Set apart = inViewRoot.ByName(elm.baseName, aPos)
            If apart Is Nothing Then
                Exit Function
            Else
                Dim apartElement As IXMLDOMElement
                Set apartElement = apart.ownerBuffer.xmlNode(elm.baseName, aPos)
                If Not UpdateFields(apart, elm, scount) Then Exit Function
            End If
        Next elm
        Else
            If inDocRoot.Text <> "" Then inViewRoot.Value = inDocRoot.Text
        End If
    Else
        If inDocRoot.Text <> "" Then inViewRoot.Value = inDocRoot.Text
    End If
    
    UpdateFields = True
End Function

Private Sub UpdateAttributes(source As IXMLDOMElement, destination As IXMLDOMElement)
  Dim i As Integer
  Dim newAttr As IXMLDOMAttribute
  For i = 0 To source.Attributes.length - 1
     If destination.Attributes.getNamedItem(source.Attributes(i).baseName) Is Nothing Then
        Set newAttr = destination.ownerDocument.createAttribute(source.Attributes(i).baseName)
        newAttr.Value = source.Attributes(i).Text
        destination.Attributes.setNamedItem newAttr
     End If
  Next
End Sub

Private Function ParseOnLineOutput(acomarea As Buffer, aresult As cSNAResult) As String
    
    If aresult.ErrCode <> 0 Then
        If SenseCode <> "" Or SenseCodeMessage <> "" Then
            acomarea.v2Value("RESP/RC") = 999
            acomarea.v2Value("RESP/RC_MSG") = SenseCode
            acomarea.v2Value("RESP/RC_TXT") = SenseCodeMessage
        Else
            acomarea.v2Value("RESP/RC") = aresult.ErrCode
            acomarea.v2Value("RESP/RC_TXT") = aresult.ErrMessage
        End If
    End If
    
    Dim doc
    Set doc = acomarea.GetXMLView
    doc.documentElement.setAttribute "_journalID", aresult.MessageID
    ParseOnLineOutput = doc.XML

End Function

Public Function ParseCall(callElm As IXMLDOMElement) As String
    Dim comareaid As String
    Dim idattr As IXMLDOMAttribute
    If Not (content Is Nothing) Then
        Set idattr = content.Attributes.getNamedItem("id")
        If Not (idattr Is Nothing) Then comareaid = idattr.Value
    End If
    If Not (callElm Is Nothing) Then
        Set idattr = callElm.Attributes.getNamedItem("id")
        If Not (idattr Is Nothing) Then comareaid = idattr.Value
    End If
    ParseCall = ParseCallWithID(callElm, comareaid, "")
End Function

Public Function ParseCallWithID(callElm As IXMLDOMElement, CallID As String, ODataName As String) As String
    Dim filename As String
    Dim methodname As String
    Dim trncall As String
    Dim InputName As String
    Dim OutputName As String
    Dim filenameattr As IXMLDOMAttribute
    Dim trncounterattr As IXMLDOMAttribute
    Dim methodnode As IXMLDOMElement
    Dim methodnameattr As IXMLDOMAttribute
    Dim trncallattr As IXMLDOMAttribute
    Dim inputnameattr As IXMLDOMAttribute
    Dim outputnameattr As IXMLDOMAttribute
    Dim snapoollinkattr As IXMLDOMAttribute
    Dim snapoollinkvalue As String
    snapoollinkvalue = ""
    
    If Not (content Is Nothing) Then
        Set filenameattr = content.Attributes.getNamedItem("filename")
        Set trncounterattr = content.Attributes.getNamedItem("trncounter")
        If Not (filenameattr Is Nothing) Then filename = filenameattr.Value
        If Not (trncounterattr Is Nothing) Then
            If trncounterattr.Value = "none" Then
                trncounter = False
            End If
        End If
    End If
    If Not (callElm Is Nothing) Then
        Set filenameattr = callElm.Attributes.getNamedItem("filename")
        Set trncounterattr = callElm.Attributes.getNamedItem("trncounter")
        If Not (filenameattr Is Nothing) Then filename = filenameattr.Value
        If Not (trncounterattr Is Nothing) Then
            If trncounterattr.Value = "none" Then
                trncounter = False
            End If
        End If
    End If
    If Not (content Is Nothing) Then
        Set methodnode = content.selectSingleNode("./method")
        If methodnode Is Nothing Then
        Else
            Set methodnameattr = methodnode.Attributes.getNamedItem("name")
            Set trncallattr = methodnode.Attributes.getNamedItem("trncall")
            Set inputnameattr = methodnode.Attributes.getNamedItem("inputname")
            Set outputnameattr = methodnode.Attributes.getNamedItem("outputname")
            Set snapoollinkattr = methodnode.Attributes.getNamedItem("snapoollink")
            If Not (methodnameattr Is Nothing) Then methodname = methodnameattr.Value
            If Not (trncallattr Is Nothing) Then trncall = trncallattr.Value
            If Not (inputnameattr Is Nothing) Then InputName = inputnameattr.Value
            If Not (outputnameattr Is Nothing) Then OutputName = outputnameattr.Value
            If Not (snapoollinkattr Is Nothing) Then snapoollinkvalue = snapoollinkattr.Value
        End If
    End If
    If Not (callElm Is Nothing) Then
        Set methodnode = callElm.selectSingleNode("./method")
        If methodnode Is Nothing Then
        Else
            Set methodnameattr = methodnode.Attributes.getNamedItem("name")
            Set trncallattr = methodnode.Attributes.getNamedItem("trncall")
            Set inputnameattr = methodnode.Attributes.getNamedItem("inputname")
            Set outputnameattr = methodnode.Attributes.getNamedItem("outputname")
            Set snapoollinkattr = methodnode.Attributes.getNamedItem("snapoollink")
            If Not (methodnameattr Is Nothing) Then methodname = methodnameattr.Value
            If Not (trncallattr Is Nothing) Then trncall = trncallattr.Value
            If Not (inputnameattr Is Nothing) Then InputName = inputnameattr.Value
            If Not (outputnameattr Is Nothing) Then OutputName = outputnameattr.Value
            If Not (snapoollinkattr Is Nothing) Then snapoollinkvalue = snapoollinkattr.Value
            callElm.removeChild methodnode
        End If
    End If
    
    If BuildComArea(Container, CallID, filename) Then
        If Not (callElm Is Nothing) Then
            Container.ByName(CallID).ClearData
            If Not UpdateFields(Container.ByName(CallID), callElm, 1) Then Exit Function
        End If
        If (ODataName <> "") Then
            If Not (Container.ByName(CallID).ByName(ODataName, 1) Is Nothing) Then
                Container.ByName(CallID).ByName(ODataName, 1).ClearData
            End If
        End If
        
        SenseCode = "": SenseCodeMessage = ""
        If methodname <> "" Then
            If trncallattr Is Nothing Then
                Set onLineComResult = ComAreaCom_(methodname, Container.ByName(CallID), InputName, OutputName, , , , , trncounter)
            Else
                Set onLineComResult = ComAreaCom_(methodname, Container.ByName(CallID), InputName, OutputName, , trncallattr.Value, , , trncounter)
            End If
            MessageID = onLineComResult.MessageID
            ParseCallWithID = ParseOnLineOutput(Container.ByName(CallID), onLineComResult)
        Else
            'αν χρησιμοποιειται μονο απο 9060 δεν παιζει επρεπε να διαβιβαζει
            'methodname = "S1041"
            'Set onLineComResult = ComAreaCom_(methodname, Container.ByName(CallID), "IDATA", "ODATA", , , , , trncounter)
            'ParseCallWithID = ParseOnLineOutput(Container.ByName(CallID), onLineComResult)
            ParseCallWithID = Container.ByName(CallID).GetXMLView.XML
        End If
    Else
        ParseCallWithID = ""
    End If
End Function

Public Function BuilderParseArea(callElm As IXMLDOMElement) As String
Dim comareaid As String
Dim filename As String
    Dim idattr As IXMLDOMAttribute
    Dim filenameattr As IXMLDOMAttribute
    Set idattr = content.Attributes.getNamedItem("id")
    Set filenameattr = content.Attributes.getNamedItem("filename")
    If Not (idattr Is Nothing) Then comareaid = idattr.Value
    If Not (filenameattr Is Nothing) Then filename = filenameattr.Value
    If BuildComArea(Container, comareaid, filename) Then
        BuilderParseArea = Container.ByName(comareaid).xmlDocV2.XML
    Else
        BuilderParseArea = ""
    End If
End Function

Public Property Get BufferByName(inName As String)
   Set BufferByName = GenWorkForm.AppBuffers.ByName(inName)
End Property

Public Function LoadXML(inDocument As String) As String
    Set onLineComResult = New cSNAResult
    onLineComResult.ErrCode = 0
    
    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim astr As String
    trncounter = True
    
    Set c4EyesHandler = New c4EyesComareaHandler
    astr = c4EyesHandler.LoadXML(Me)
    
    If astr <> "" Then
        LoadXML = Replace(astr, Chr(7), "")
    Else
        LoadXML = "<ERROR></ERROR>"
    End If
    UpdatexmlEnvironment "MESSAGEID", MessageID
    
    
    workDocument.LoadXML LoadXML
    If Not (workDocument.parseError Is Nothing) Then
       
        If workDocument.parseError.errorCode <> 0 Then
            workDocument.appendChild workDocument.createElement("ERROR")
            Dim positem As IXMLDOMElement
            Dim messageitem As IXMLDOMElement
            Dim dataitem As IXMLDOMElement
        
            Set positem = workDocument.createElement("FilePos")
            Set messageitem = workDocument.createElement("Message")
            Set dataitem = workDocument.createElement("Data")
            workDocument.documentElement.appendChild positem
            workDocument.documentElement.appendChild messageitem
            workDocument.documentElement.appendChild dataitem
        
            positem.Text = workDocument.parseError.filepos
            messageitem.Text = workDocument.parseError.Reason
            dataitem.Text = LoadXML
        
            SaveXmlFile "ComAreaErrorDocument.xml", workDocument
        Else
            Dim odata As IXMLDOMElement
            Set odata = workDocument.selectSingleNode("//ODATA")
            If Not odata Is Nothing Then
                Dim Message As IXMLDOMElement
                Set Message = odata.selectSingleNode("WARNINGS/WARN_MESSG")
                If Not Message Is Nothing Then
                    Dim warning As String
                    warning = Trim(Message.Text)
                    If warning <> "" Then
                        MsgBox warning, vbInformation, "Ειδοποίηση"
                    End If
                End If
            End If
        End If
    End If
    
End Function

Public Function XML() As String
    If workDocument Is Nothing Then
        XML = ""
    Else
        XML = workDocument.XML
    End If
End Function

Public Property Get Buffers() As Buffers
    Set Buffers = Container
End Property

Public Property Get Buffer() As Buffer

Dim comareaid As String
Dim idattr As IXMLDOMAttribute
    
    Set idattr = content.Attributes.getNamedItem("id")
    If Not (idattr Is Nothing) Then comareaid = idattr.Value
    If comareaid <> "" Then
        Set Buffer = Container.ByName(comareaid)
    Else
        Set Buffer = Nothing
    End If
End Property

Private Sub Class_Initialize()
    Set Container = GenWorkForm.AppBuffers
End Sub

Private Sub Class_Terminate()
    Set Container = Nothing
    Set content = Nothing
    Set workDocument = Nothing
    Set c4EyesHandler = Nothing
End Sub
