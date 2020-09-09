VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLDocumentJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Title As String
Private mergedDoc 'As MSXML2.DOMDocument30
Private mergedxml As String

Public exitformflag As Boolean

Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement

Private DocumentStack As New Collection

Private prevchiefname, prevmanagername As String

Public Property Get DebugMergedXml() As String
    DebugMergedXml = MergedDocXml
End Property

Private Property Get MergedDocXml() As String
    MergedDocXml = ""
    If mergedDoc Is Nothing Then
    Else
        If mergedDoc.documentElement Is Nothing Then
        Else
            MergedDocXml = mergedDoc.documentElement.XML
        End If
    End If
End Property

Public Function chkEqualPart(elm As IXMLDOMElement, docname As String) As Boolean
Dim selectAttr As IXMLDOMAttribute, localstr As String, cmpstr As String
    cmpstr = "": cmpstr = ""
    Set selectAttr = elm.getAttributeNode("select")
    If Not (selectAttr Is Nothing) Then
        If Not (mergedDoc.selectSingleNode(selectAttr.Text) Is Nothing) Then
            localstr = mergedDoc.selectSingleNode(selectAttr.Text).XML
        End If
    Else
        localstr = MergedDocXml
    End If
    If TypeOf Manager.XmlObjectList.item(docname) Is MSXML2.DOMDocument30 Then
        If Not (Manager.XmlObjectList.item(docname).documentElement Is Nothing) Then
            cmpstr = Manager.XmlObjectList.item(docname).documentElement.XML
        End If
    Else
        cmpstr = Manager.XmlObjectList.item(docname).XML
    End If

    chkEqualPart = (cmpstr = localstr)
End Function

Public Function chkSelectPart(elm As IXMLDOMElement, statement As String) As Boolean
    Dim xpath As String
    xpath = elm.getAttribute(statement)
    If xpath <> "" Then
        chkSelectPart = (mergedDoc.SelectNodes(xpath).length > 0)
    Else
        chkSelectPart = False
    End If
End Function


Public Function ParseJob(ByRef rootElm As IXMLDOMElement) As Boolean
    Dim elm As IXMLDOMNode, nameattr As IXMLDOMAttribute, addressAttr As IXMLDOMAttribute, selectAttr As IXMLDOMAttribute, childElm As IXMLDOMElement
    Dim childnode As IXMLDOMNode
    Dim cmpattr As IXMLDOMAttribute
    Dim InputName As String, OutputName As String
    Dim functionnameattr As IXMLDOMAttribute
    Dim functionName As String, aXmlObject
    Dim tmpDocument As MSXML2.DOMDocument30
    Dim localstr As String, cmpstr As String
    Dim currStatement As String
    
    ParseJob = False
    If mergedDoc Is Nothing Then Set mergedDoc = CreateObject("Msxml2.DOMDocument.6.0") 'New MSXML2.DOMDocument30
    On Error GoTo processingError
    For Each elm In rootElm.childNodes
        If elm.nodeType = NODE_ELEMENT Then
            If elm.baseName = "pause" Then
                localstr = ""
            ElseIf elm.baseName = "clear" Then
                Set mergedDoc = Nothing
                Set mergedDoc = CreateObject("Msxml2.DOMDocument.6.0") 'New MSXML2.DOMDocument30:
                mergedxml = "":
            ElseIf elm.baseName = "job" Then
                Set nameattr = elm.Attributes.getNamedItem("name")
                If nameattr Is Nothing Then
'                    For Each childElm In elm.childNodes
'                        Set aXmlObject = Manager.ParseElement(childElm)
'                        If aXmlObject Is Nothing Then
'                            mergedxml = mergedxml & childElm.Xml:
'                        Else
'                            mergedxml = mergedxml & aXmlObject.Xml:
'                        End If
'                        mergedDoc.LoadXml "<" & Title & ">" & mergedxml & "</" & Title & ">"
'                    Next childElm
                Else
                    Set functionnameattr = elm.Attributes.getNamedItem("name")
                    If functionnameattr Is Nothing Then
                        functionName = ""
                    Else
                        functionName = functionnameattr.value
                    End If
                    
                    Set aXmlObject = Manager.XmlObjectList.item(functionName)
                    If Not (mergedDoc Is Nothing) Then
                        If Not (mergedDoc.documentElement Is Nothing) Then
                            mergedDoc.LoadXML aXmlObject.LoadXML(MergedDocXml)
                        Else
                            mergedDoc.LoadXML aXmlObject.LoadXML(mergedxml)
                        End If
                    Else
                        mergedDoc.LoadXML aXmlObject.LoadXML(mergedxml)
                    End If
                    'mergedDoc.LoadXML aXmlObject.LoadXML(mergedDoc.XML)
                    mergedxml = MergedDocXml
                End If
            ElseIf elm.baseName = "input" Then
                Set nameattr = elm.Attributes.getNamedItem("name")
                If nameattr Is Nothing Then
                    For Each childElm In elm.childNodes
                        Set aXmlObject = Manager.ParseElement(childElm)
                        If aXmlObject Is Nothing Then
                            mergedxml = mergedxml & childElm.XML:
                        Else
                            mergedxml = mergedxml & aXmlObject.XML:
                        End If
                        mergedDoc.LoadXML "<" & Title & ">" & mergedxml & "</" & Title & ">"
                    Next childElm
                Else
                    InputName = nameattr.Text
                    If Not Manager.XmlObjectList.item(InputName) Is Nothing Then
                        mergedxml = mergedxml & Manager.XmlObjectList.item(InputName).XML
                        mergedDoc.LoadXML "<" & Title & ">" & mergedxml & "</" & Title & ">"
                    Else
                        If Not GenWorkForm.L2AppDocs.item(InputName) Is Nothing Then
                            Dim bdoc As New MSXML2.DOMDocument30
                            Set bdoc = GenWorkForm.L2AppDocs.item(InputName)
                            mergedxml = mergedxml & bdoc.XML
                            mergedDoc.LoadXML "<" & Title & ">" & mergedxml & "</" & Title & ">"
                            Set bdoc = Nothing
                        End If
                    End If
                End If
            ElseIf elm.baseName = "save" Then
                Set nameattr = elm.Attributes.getNamedItem("filename")
                If nameattr Is Nothing Then
                    Set nameattr = elm.Attributes.getNamedItem("name")
                    If nameattr Is Nothing Then
                    Else
                        Dim scopeattr As IXMLDOMAttribute
                        Set scopeattr = elm.Attributes.getNamedItem("scope")
                        If scopeattr Is Nothing Then
                            Set tmpDocument = New MSXML2.DOMDocument30
                            tmpDocument.LoadXML MergedDocXml
                            Manager.XmlObjectList.add tmpDocument, nameattr.Text
                            Manager.nameList.add nameattr.Text, nameattr.Text
                            Manager.typeList.add "document", nameattr.Text
                        ElseIf UCase(scopeattr.Text) = "GLOBAL" Then
                            Set tmpDocument = New MSXML2.DOMDocument30
                            tmpDocument.LoadXML MergedDocXml
                            GenWorkForm.L2AppDocs.addunique tmpDocument, nameattr.Text
                        End If
                    End If
                Else
                    If Len(nameattr.Text) > 3 Then
                        If UCase(Left(nameattr.Text, 3)) = "C:\" Then
                            mergedDoc.Save Replace(UCase(nameattr.Text), "C:", NetworkHomeDir, 1, 1)
                        End If
                    End If
                    'If cDebug = 1 Then mergedDoc.save nameattr.Text
                End If
            ElseIf elm.baseName = "load" Then
                Set nameattr = elm.Attributes.getNamedItem("filename")
                If nameattr Is Nothing Then
                    Set nameattr = elm.Attributes.getNamedItem("name")
                    If nameattr Is Nothing Then
                    Else
                        mergedDoc.LoadXML Manager.XmlObjectList.item(nameattr.Text).XML
                    End If
                ElseIf UCase(nameattr.Text) = "USER" Then
                    Dim iboxtitle As String, titleattr As IXMLDOMAttribute, filename As String
                    iboxtitle = "≈ÈÛ·„˘„ﬁ ”ÙÔÈ˜Âﬂ˘Ì"
                    Set titleattr = elm.Attributes.getNamedItem("title")
                    If titleattr Is Nothing Then
                    Else
                        iboxtitle = titleattr.Text
                    End If
                    filename = ""
                    filename = InputBox("œÌÔÏ· ·Ò˜ÂﬂÔı:", iboxtitle)
                    If filename <> "" Then mergedDoc.Load filename
                    
                    Set titleattr = Nothing
                Else
                    mergedDoc.Load nameattr.Text
                End If
            ElseIf elm.baseName = "function" Then
                Set nameattr = elm.Attributes.getNamedItem("name")
                If nameattr Is Nothing Then
                    Set addressAttr = elm.Attributes.getNamedItem("address")
                    If addressAttr Is Nothing Then
                        For Each childnode In elm.childNodes
                            If childnode.nodeType = NODE_ELEMENT Then
                                Set aXmlObject = Manager.ParseElement(childnode)
                                If aXmlObject Is Nothing Then
                                Else
                                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                                        If Not Manager Is Nothing Then
                                            Manager.ShowMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ....."
                                        End If
                                    End If
                                    mergedDoc.LoadXML aXmlObject.LoadXML(MergedDocXml)
                                    mergedxml = MergedDocXml
                                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                                        If Not aXmlObject.ComResult Is Nothing Then
                                            aXmlObject.ComResult.UpdateXmlDocumentManager Manager
                                        End If
                                    End If
                                End If
                            End If
                        Next childnode
                    Else
                        For Each childElm In mergedDoc.SelectNodes(addressAttr.Text)
                            Set aXmlObject = Manager.ParseElement(childElm)
                            If aXmlObject Is Nothing Then
                            Else
                                If TypeOf aXmlObject Is shine.cXmlComArea Then
                                    If Not Manager Is Nothing Then
                                        Manager.ShowMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ....."
                                    End If
                                End If
                                mergedDoc.LoadXML aXmlObject.LoadXML(MergedDocXml)
                                mergedxml = MergedDocXml
                                If TypeOf aXmlObject Is shine.cXmlComArea Then
                                    If Not aXmlObject.ComResult Is Nothing Then
                                        aXmlObject.ComResult.UpdateXmlDocumentManager Manager
                                    End If
                                End If
                            End If
                        Next childElm
                    End If
                Else
                    Set functionnameattr = elm.Attributes.getNamedItem("name")
                    If functionnameattr Is Nothing Then
                        functionName = ""
                    Else
                        functionName = functionnameattr.value
                    End If
                    Set aXmlObject = Manager.XmlObjectList.item(functionName)
                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                        If Not Manager Is Nothing Then
                            Manager.ShowMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ....."
                        End If
                    End If
                    mergedDoc.LoadXML aXmlObject.LoadXML(MergedDocXml)
                    mergedxml = MergedDocXml
                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                        If Not aXmlObject.ComResult Is Nothing Then
                            aXmlObject.ComResult.UpdateXmlDocumentManager Manager
                        End If
                    End If
                End If
            ElseIf elm.baseName = "output" Then
                Dim outputattr As IXMLDOMAttribute
                Set outputattr = elm.Attributes.getNamedItem("name")
                If outputattr Is Nothing Then
                    OutputName = ""
                Else
                    OutputName = outputattr.value
                End If
                Set aXmlObject = Manager.XmlObjectList.item(OutputName)
                If aXmlObject Is Nothing Then
                    Dim adoc As New MSXML2.DOMDocument30
                    adoc.LoadXML MergedDocXml
                    GenWorkForm.L2AppDocs.Update adoc, OutputName
                    Set adoc = Nothing
                Else
                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                        If Not Manager Is Nothing Then
                            Manager.ShowMessage "ƒ…¡¬…¬¡”« ”‘œ…◊≈…ŸÕ....."
                        End If
                    End If
                    Set selectAttr = elm.Attributes.getNamedItem("select")
                    If selectAttr Is Nothing Then
                        aXmlObject.LoadXML MergedDocXml
                    Else
                        aXmlObject.LoadXML mergedDoc.selectSingleNode(selectAttr.Text).XML
                    End If
                    If TypeOf aXmlObject Is shine.cXmlDocument Then
                    'If Manager.typeList(OutputName) = "document" Then
                        Set childElm = aXmlObject.documentElement
                        If childElm Is Nothing Then
                        Else
                            aXmlObject.LoadXML childElm.XML
                        End If
                    End If
                    If TypeOf aXmlObject Is shine.cXmlComArea Then
                        If Not aXmlObject.ComResult Is Nothing Then
                            aXmlObject.ComResult.UpdateXmlDocumentManager Manager
                        End If
                    End If
                End If
                Set outputattr = Nothing
            ElseIf elm.baseName = "while" Then
                Dim res As Boolean
                Do
                    If Not (elm.Attributes.getNamedItem("equal") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("equal")
                        If chkEqualPart(elm, cmpattr.value) Then
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        Else
                            Exit Do
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("notequal") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("notequal")
                        If Not chkEqualPart(elm, cmpattr.value) Then
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        Else
                            Exit Do
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("select") Is Nothing) Then
                        If chkSelectPart(elm, "select") Then
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        Else
                            Exit Do
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("notselect") Is Nothing) Then
                        If Not chkSelectPart(elm, "notselect") Then
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        Else
                            Exit Do
                        End If
                    End If
                Loop
            ElseIf elm.baseName = "else" Then
            ElseIf elm.baseName = "elseif" Then
            ElseIf elm.baseName = "if" Then
                Dim astatement As String, ifsucceed As Boolean
                ifsucceed = False
                Do While Not ifsucceed
                    If Not (elm.Attributes.getNamedItem("equal") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("equal")
                        Set selectAttr = elm.Attributes.getNamedItem("select")
                        If Not (selectAttr Is Nothing) Then
                            If mergedDoc.selectSingleNode(selectAttr.Text) Is Nothing Then
                                localstr = ""
                            Else
                                localstr = mergedDoc.selectSingleNode(selectAttr.Text).XML
                            End If
                        Else
                            localstr = MergedDocXml
                        End If
                        If TypeOf Manager.XmlObjectList.item(cmpattr.value) Is MSXML2.DOMDocument30 Then
                            If Manager.XmlObjectList.item(cmpattr.value).documentElement Is Nothing Then
                                cmpstr = ""
                            Else
                                cmpstr = Manager.XmlObjectList.item(cmpattr.value).documentElement.XML
                            End If
                        Else
                            cmpstr = Manager.XmlObjectList.item(cmpattr.value).XML
                        End If
                        If localstr = cmpstr Then
                            ifsucceed = True
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("notequal") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("notequal")
                        Set selectAttr = elm.Attributes.getNamedItem("select")
                        If Not (selectAttr Is Nothing) Then
                            If mergedDoc.selectSingleNode(selectAttr.Text) Is Nothing Then
                                localstr = ""
                            Else
                                localstr = mergedDoc.selectSingleNode(selectAttr.Text).XML
                            End If
                        Else
                            localstr = MergedDocXml
                        End If
                        If TypeOf Manager.XmlObjectList.item(cmpattr.value) Is MSXML2.DOMDocument30 Then
                            If Manager.XmlObjectList.item(cmpattr.value).documentElement Is Nothing Then
                                cmpstr = ""
                            Else
                                cmpstr = Manager.XmlObjectList.item(cmpattr.value).documentElement.XML
                            End If
                        Else
                            cmpstr = Manager.XmlObjectList.item(cmpattr.value).XML
                        End If
                        If localstr <> cmpstr Then
                            ifsucceed = True
                            res = ParseJob(elm)
                            If Not res Then Exit Function
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("select") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("select")
                        astatement = cmpattr.nodeValue
                        If astatement <> "" Then
                            If mergedDoc.SelectNodes(astatement).length > 0 Then
                                ifsucceed = True
                                res = ParseJob(elm)
                                If Not res Then Exit Function
                            End If
                        End If
                    ElseIf Not (elm.Attributes.getNamedItem("notselect") Is Nothing) Then
                        Set cmpattr = elm.Attributes.getNamedItem("notselect")
                        astatement = cmpattr.nodeValue
                        If astatement <> "" Then
                            If mergedDoc.SelectNodes(astatement).length = 0 Then
                                ifsucceed = True
                                res = ParseJob(elm)
                                If Not res Then Exit Function
                            End If
                        End If
                    End If
                    
                    If ifsucceed Then Exit Do
                    
                    Dim elseElm As IXMLDOMElement
                    Set elseElm = Nothing
                    If Not elm.nextSibling Is Nothing Then
                        If elm.nextSibling.nodeType = NODE_ELEMENT Then
                            Set elseElm = elm.nextSibling
                        End If
                    End If
                        
                    If elseElm Is Nothing Then
                        ifsucceed = True
                    Else
                        If elseElm.baseName = "else" Then
                            ifsucceed = True
                            res = ParseJob(elseElm)
                            If Not res Then Exit Function
                        ElseIf elseElm.baseName = "elseif" Then
                            Set elm = elseElm
                        Else
                            ifsucceed = True
                        End If
                    End If
                    
                    Set elseElm = Nothing
                Loop
            ElseIf elm.baseName = "push" Then
                Dim stackdoc As MSXML2.DOMDocument
                Set stackdoc = New MSXML2.DOMDocument30
                stackdoc.LoadXML MergedDocXml
                DocumentStack.add stackdoc
                Set stackdoc = Nothing
            ElseIf elm.baseName = "pop" Then
                If DocumentStack.count > 0 Then
                    mergedDoc.LoadXML DocumentStack(DocumentStack.count).XML
                    DocumentStack.Remove DocumentStack.count
                    mergedxml = MergedDocXml
                End If
            ElseIf elm.baseName = "exitjob" Then
                Exit Function
            ElseIf elm.baseName = "exitform" Then
                exitformflag = True
                Exit Function
                'Unload Manager.owner.activeform
            End If
        End If
    Next elm
    
    Set elm = Nothing
    Set nameattr = Nothing
    Set addressAttr = Nothing
    Set selectAttr = Nothing
    Set childElm = Nothing
    Set childnode = Nothing
    Set cmpattr = Nothing
    Set functionnameattr = Nothing
    Set aXmlObject = Nothing
    Set tmpDocument = Nothing
    
    ParseJob = True: Exit Function
    
processingError:
    Dim aMsg As String
    aMsg = "¡›Ùı˜Â Á ÂÍÙ›ÎÂÛÁ ÙÁÚ ‰È·‰ÈÍ·Ûﬂ·Ú: " & Title & vbCrLf
    If elm Is Nothing Then
    Else
        aMsg = aMsg & elm.XML
    End If
    'MsgBox aMsg, vbCritical, "À‹ËÔÚ ≈ÂÓÂÒ„·Ûﬂ·Ú..."
    LogMsgbox aMsg, vbCritical, "À‹ËÔÚ ≈ÂÓÂÒ„·Ûﬂ·Ú...", Err
End Function


Public Function XML() As String
    '<JOB TITLE="xxxxx">
    '   <INPUT TITLE="xxxx" />.....
    '   <FUNCTION TITLE="xxxx" />.....
    '   <OUTPUT TITLE="xxxx" />.....
    '</JOB>
    '
    Title = content.getAttribute("name")
    Set mergedDoc = CreateObject("Msxml2.DOMDocument.6.0") 'New MSXML2.DOMDocument30:
    mergedxml = ""
    Dim res As Boolean
    
     BackUpKeys
     
    res = ParseJob(content)
    
    If ExistsChangeKeysAttribute = True Then
    Else
        UpdateChiefKey (prevchiefname)
        UpdateManagerKey (prevmanagername)
    End If
   
    If Not exitformflag Then
    XML = MergedDocXml
    Else
    Set mergedDoc = Nothing
    Unload Manager.owner.activeform
    End If
    'End If
End Function


Public Function LoadXML(Document As String) As String
    
    'Dim attrchange As IXMLDOMAttribute
    'Dim changekeys As Boolean
    'changekeys = False
    'Set attrchange = content.Attributes.getNamedItem("changekeys")
    'If Not (attrchange Is Nothing) Then changekeys = True
    
    BackUpKeys
    
    If content.Attributes.getNamedItem("name") Is Nothing Then
    Else
        Title = content.Attributes.getNamedItem("name").Text
    End If
    'If mergedDoc Is Nothing Then
    Set mergedDoc = CreateObject("Msxml2.DOMDocument.6.0") 'New MSXML2.DOMDocument30
    mergedDoc.LoadXML Document: mergedxml = Document
    Dim res As Boolean
    res = ParseJob(content)
    
    If ExistsChangeKeysAttribute = True Then
    Else
        UpdateChiefKey (prevchiefname)
        UpdateManagerKey (prevmanagername)
    End If
    
    If Not exitformflag Then
        LoadXML = MergedDocXml
    Else
        Set mergedDoc = Nothing

        Unload Manager.owner.activeform
    End If
    
    'End If
End Function

Private Sub BackUpKeys()
    If prevchiefname = "" Then
        prevchiefname = cCHIEFUserName
        prevmanagername = cMANAGERUserName
   End If
End Sub
Private Function ExistsChangeKeysAttribute() As Boolean
      Dim attrchange As IXMLDOMAttribute
      Set attrchange = content.Attributes.getNamedItem("changekeys")
      If attrchange Is Nothing Then ExistsChangeKeysAttribute = False: Exit Function
      ExistsChangeKeysAttribute = True
     
End Function
