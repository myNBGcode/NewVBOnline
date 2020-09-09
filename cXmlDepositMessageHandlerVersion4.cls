VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlDepositMessageHandlerVersion4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public workDocument As MSXML2.DOMDocument30
Public Manager As cXMLDocumentManager

Public ComArea As cXmlComArea
Public content As IXMLDOMElement

Dim msglist As New Collection
Dim authlist As New Collection
Dim displaylist As New Collection
Dim doc As New MSXML2.DOMDocument30

Const TellerKey As String = "T"
Const ChiefKey As String = "C"
Const ManagerKey As String = "M"

Private Function PrepareDisplayDocument()
    
       Dim aItem As IXMLDOMNode
        'clear doc
        Dim Node As IXMLDOMNode
        For Each Node In doc.documentElement.childNodes
            doc.documentElement.removeChild Node
        Next
        
        'clear msglist
        For Each aItem In msglist
           doc.documentElement.appendChild aItem
        Next
        While msglist.Count > 0
            displaylist.add (msglist.Item(1))
            msglist.Remove (1)
        Wend
        
        Set Node = Nothing
        Set aItem = Nothing

End Function

Private Function DisplayMessages() As String
    Dim res As String
    
    If msglist.Count > 0 Then
         PrepareDisplayDocument
        res = L2ShowDepositMessages(doc.documentElement)
        If res <> "" Then
            ShowError res
            DisplayMessages = res
            Exit Function
        End If
    End If
End Function
Private Function ProcessMessages(workDocument As MSXML2.DOMDocument30, Key As String)
    Dim messagesList As IXMLDOMNodeList
    
    Dim res As String
    
    On Error GoTo processingError
    
    Set messagesList = workDocument.SelectNodes("//ODATA/STRMSG[MSG_CODE/text()!='' and MSG_KEYS/text()= '" + Key + "']")
    If (messagesList.length = 0) Then
        ProcessMessages = ""
        Exit Function
    End If
    
    Dim msg As IXMLDOMElement
    For Each msg In messagesList
        msglist.add msg
    Next
    Set msg = Nothing
    
    res = DisplayMessages
    If res <> "" Then
       ProcessMessages = res
       Exit Function
    End If
    
    
'    Select Case UCase(Key)
'        Case TellerKey:
'
'        Case ChiefKey:
'            res = L2ChiefKey(Nothing)
'        Case ManagerKey:
'            res = L2ManagerKey(Nothing)
'    End Select
    
    Dim authdoc As New MSXML2.DOMDocument60
    authdoc.LoadXML "<input><DoNotAcceptCurrentUserAsAuthUser>true</DoNotAcceptCurrentUserAsAuthUser></input>"

    If UCase(Key) = ChiefKey Then
        res = L2ChiefKey(authdoc.documentElement)
    ElseIf UCase(Key) = ManagerKey Then
        res = L2ManagerKey(authdoc.documentElement)
    End If

    Set authdoc = Nothing
    
    If res <> "" Then
        Dim resultdocument As New MSXML2.DOMDocument30
        resultdocument.LoadXML res
        If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            Load XMLMessageForm
            Set XMLMessageForm.MessageDocument = resultdocument
            XMLMessageForm.Show vbModal
            ProcessMessages = res
            Exit Function
        Else
            If UCase(Key) = ChiefKey Then
                authlist.add UCase(cCHIEFUserName), Key
            ElseIf UCase(Key) = ManagerKey Then
                authlist.add UCase(cMANAGERUserName), Key
            End If
        End If
        
        Set resultdocument = Nothing
    End If
    
    Exit Function
    
processingError:
    Dim aMsg As String
    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας: " & "ProcessMessages" & vbCrLf
    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
    ProcessMessages = "<MESSAGE><ERROR><LINE>" & error & "</LINE></ERROR></MESSAGE>"
    
End Function

Private Function ClearAuthList()
    If Not (authlist Is Nothing) Then
        While authlist.Count > 0
            authlist.Remove (1)
        Wend
    End If
End Function
Public Function LoadXML(inDocument As String) As String
        
    Dim stringerror As String
    
    Dim astr As String, comareaid As String, Key As String, res As String
    Dim elm As IXMLDOMElement, aItem As IXMLDOMNode
    Dim msg As IXMLDOMElement
    Dim area As IXMLDOMAttribute
    Dim error As IXMLDOMElement

    Dim msg_text As IXMLDOMNode
    Dim MSG_KEYS  As IXMLDOMElement
    
    On Error GoTo processingError
     
     If Trim(inDocument) = "" Then 'ERROR
        stringerror = "Λάθος παραμετρος inDocument κενή.. "
        Err.Raise 1300, "LoadXml", stringerror
        Exit Function
    End If
    
    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    
    If (doc.documentElement Is Nothing) Then doc.appendChild doc.createElement("root")
    
    If Not (content Is Nothing) Then
        Set area = content.Attributes.getNamedItem("com")
        If Not (area Is Nothing) Then
          Set ComArea = Manager.XmlObjectList.Item(area.Text)
        End If
    End If
    If Not (ComArea Is Nothing) Then
        comareaid = ComArea.content.Attributes.getNamedItem("id").Text
    End If
    
    ClearAuthList
    
    Dim rc As Integer
    Dim rc_txt As String
    rc = CInt(Trim(workDocument.selectSingleNode("//RESP/RC").Text))
    rc_txt = workDocument.selectSingleNode("//RESP/RC_TXT").Text
    
    If rc > 0 Then 'ERROR
        LoadXML = inDocument
    ElseIf rc <= 0 Then
        res = ProcessMessages(workDocument, "T")
        If res <> "" Then
           LoadXML = res
           Exit Function
        End If
        res = ProcessMessages(workDocument, "C")
        If res <> "" Then
           LoadXML = res
           Exit Function
        End If
        res = ProcessMessages(workDocument, "M")
        If res <> "" Then
           LoadXML = res
           Exit Function
        End If
        
        If rc = 0 Then 'δειξε οτι περισσεψε δεν θα ξαναδιαβιβασει
            DisplayMessages
            LoadXML = inDocument 'ΤΕΛΟΣ
        ElseIf rc = -1 Then
         If Not (ComArea Is Nothing) Then
            comareaid = ComArea.content.Attributes.getNamedItem("id").Text
            Set elm = UpdateAreaDocument(inDocument)
            If Not (elm Is Nothing) Then
                inDocument = ComArea.ParseCallWithID(elm, comareaid, "ODATA")
                ComArea.ComResult.UpdateXmlDocumentManager Me.Manager
                LoadXML = LoadXML(inDocument)
            Else
                LoadXML = inDocument
            End If
         End If
        End If
    
    End If
    
    Set elm = Nothing
    Set aItem = Nothing
    Set msg = Nothing
    Set area = Nothing
    Set error = Nothing
    Set msg_text = Nothing

    Exit Function
     
processingError:
    Dim aMsg As String
    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας:  LoadXml " & vbCrLf
    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
    LoadXML = "<MESSAGE><ERROR><LINE>" & stringerror & "</LINE></ERROR></MESSAGE>"
    
End Function

Private Function UpdateAreaDocument(inDocument As String) As IXMLDOMElement
    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim copynode As IXMLDOMNode
    Dim replacenode As IXMLDOMNode
    Dim error As String
    Dim Key As String
    On Error GoTo UpdateError
    
    'update messages
    For Each copynode In workDocument.SelectNodes("//ODATA//STRMSG[MSG_CODE/text()!='']")
        Set replacenode = workDocument.selectSingleNode("//IDATA//STRMSG[MSG_CODE='']")
        If Not (replacenode Is Nothing) Then
           replacenode.selectSingleNode("MSG_CODE").Text = copynode.selectSingleNode("MSG_CODE").Text
           
           If Not (replacenode.selectSingleNode("MSG_STAT") Is Nothing) Then
            replacenode.selectSingleNode("MSG_STAT").Text = copynode.selectSingleNode("MSG_STAT").Text
           End If
           replacenode.selectSingleNode("MSG_KEYS").Text = copynode.selectSingleNode("MSG_KEYS").Text
           
           If Not (replacenode.selectSingleNode("MSG_TEXT") Is Nothing) Then
            replacenode.selectSingleNode("MSG_TEXT").Text = copynode.selectSingleNode("MSG_TEXT").Text
           End If
           Key = copynode.selectSingleNode("MSG_KEYS").Text
           If Not (replacenode.selectSingleNode("MSG_AUTH_USER") Is Nothing) Then
                If (authlist.Count >= 0) Then
                    If (Key = "C" Or Key = "M") Then
                        replacenode.selectSingleNode("MSG_AUTH_USER").Text = authlist.Item(Key)
                    End If
                End If
           End If
        ElseIf replacenode Is Nothing Then 'Error more messages than we can handle
            error = "Λάθος :ODATA returned more messages than IDATA can handle"
            Err.Raise 1300, "UpdateAreaDocument", error
            Exit Function
        End If
    Next
    
    'update timespamp
    If Not (workDocument.selectSingleNode("//ODATA//LOCK_TIMESTAMP") Is Nothing) Then
        workDocument.selectSingleNode("//IDATA//LOCK_TIMESTAMP").Text = workDocument.selectSingleNode("//ODATA//LOCK_TIMESTAMP").Text
    End If
    
    Set UpdateAreaDocument = workDocument.documentElement
    
    Set copynode = Nothing
    Set replacenode = Nothing
    
    Exit Function
UpdateError:
    Dim aMsg As String
    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας: " & "UpdateAreaDocument" & vbCrLf
    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
    Set UpdateAreaDocument = Nothing
End Function
Private Function ShowError(error As String) As String
    Dim resultdocument As New MSXML2.DOMDocument30
    resultdocument.LoadXML error
    If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
        Load XMLMessageForm
        Set XMLMessageForm.MessageDocument = resultdocument
        XMLMessageForm.Show vbModal
        ShowError = error
        Set resultdocument = Nothing
        Exit Function
    End If
    Set resultdocument = Nothing
End Function
Private Sub Class_Terminate()

    Set workDocument = Nothing
    Set Manager = Nothing
    
    Set ComArea = Nothing
    Set content = Nothing
    
    Set msglist = Nothing
    Set displaylist = Nothing
    Set authlist = Nothing
    Set doc = Nothing

End Sub
