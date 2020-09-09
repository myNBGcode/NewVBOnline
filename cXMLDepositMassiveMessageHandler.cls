VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLDepositMassiveMessageHandler"
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
Dim displaylist As New Collection
Dim doc As New MSXML2.DOMDocument30

Const TellerKey As String = "T"
Const ChiefKey As String = "C"
Const ManagerKey As String = "M"
Const MsgShow As String = "2"
Const MsgHide As String = "0"


Public Function LoadXML(inDocument As String) As String
    Dim stringerror As String
    
    Dim astr As String, comareaid As String, key As String, res As String
    Dim elm As IXMLDOMElement, aItem As IXMLDOMNode
    Dim msg As IXMLDOMElement
    Dim area As IXMLDOMAttribute
    Dim error As IXMLDOMElement

    Dim msg_text As IXMLDOMNode
    
    On Error GoTo processingError
     
     If Trim(inDocument) = "" Then 'ERROR
        stringerror = "Λάθος παραμετρος inDocument κενή.. "
        Err.Raise 1300, "LoadXml", stringerror
        Exit Function
    End If
    
    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    
    If (doc.documentElement Is Nothing) Then doc.appendChild doc.createElement("root")
    
    If content Is Nothing Then
    Else
        Set area = content.Attributes.getNamedItem("com")
        If Not (area Is Nothing) Then
          Set ComArea = Manager.XmlObjectList.item(area.Text)
        End If
    End If
    If Not (ComArea Is Nothing) Then
      comareaid = ComArea.content.Attributes.getNamedItem("id").Text
    
    End If
    
    Dim rc As Integer
    Dim rc_txt As String
    rc = CInt(Trim(workDocument.selectSingleNode("//RESP/RC").Text))
    rc_txt = workDocument.selectSingleNode("//RESP/RC_TXT").Text
    
    If rc > 0 Then 'ERROR
        LoadXML = inDocument
        Exit Function
    ElseIf rc <= 0 Then
    
    
            For Each msg In workDocument.SelectNodes("//ODATA/DB_MSG[MSG_CODE/text()!='']")

               res = processmessage(msg)
               If res <> "" Then
                    LoadXML = res
                    Exit Function
               ElseIf res = "" Then
               End If
            Next
            For Each msg In workDocument.SelectNodes("//ODATA/CR_MSG[MSG_CODE/text()!='']")


               res = processmessage(msg)
               If res <> "" Then
                    LoadXML = res
                    Exit Function
               ElseIf res = "" Then
               End If
            Next
            If rc = 0 Then 'δειξε οτι περισσεψε δεν θα ξαναδιαβιβασει
                DisplayMessages
                LoadXML = inDocument 'ΤΕΛΟΣ
                Exit Function
'            ElseIf rc = -1 Then
'             If Not (ComArea Is Nothing) Then
'
'                   comareaid = ComArea.content.Attributes.getNamedItem("id").Text
'                   Set elm = UpdateAreaDocument(inDocument)
'                   'ComArea.Buffer.ByName("ODATA").ClearData
'                   'ComArea.Buffer.ByName("ODATA").ByName("CDATA", 1).ByName("CNV_DATA", 1).ClearData
'                   If Not (elm Is Nothing) Then
'                        inDocument = ComArea.ParseCallWithID(elm, comareaid, "ODATA")
'                        ComArea.ComResult.UpdateXmlDocumentManager Me.Manager
'                        LoadXML = LoadXML(inDocument)
'                   Else
'                        LoadXML = inDocument
'                   End If
'                   Exit Function
'                End If
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
    'MsgBox aMsg, vbCritical, "Λάθος Επεξεργασίας..."
    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
    LoadXML = "<MESSAGE><ERROR><LINE>" & stringerror & "</LINE></ERROR></MESSAGE>"
    
End Function

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
        While msglist.count > 0
            displaylist.add (msglist.item(1))
            msglist.Remove (1)
        Wend
        
        Set Node = Nothing
        Set aItem = Nothing

End Function
Private Function DisplayMessages() As String
    Dim res As String
    
    If msglist.count > 0 Then
         PrepareDisplayDocument
        res = L2ShowDepositMassiveMessages(doc.documentElement)
        If res <> "" Then
            ShowError res
            DisplayMessages = res
            Exit Function
        End If
    End If
End Function
Private Function processmessage(msg As IXMLDOMElement)
    Dim error As String
    Dim key As String, res As String
   
    On Error GoTo processingError
    If Trim(msg.selectSingleNode("MSG_CODE").Text) = "" Then 'ERROR
        error = "Λάθος στα ODATA: RC=-1 και MSG_CODE κενό.. "
        Err.Raise 1300, "processmessage", error
        Exit Function
    End If
    'HandleMessage msg
    Dim MSG_STAT As IXMLDOMElement
    Dim MSG_KEYS  As IXMLDOMElement
        
    Set MSG_STAT = msg.selectSingleNode("MSG_STAT")
    Set MSG_KEYS = msg.selectSingleNode("MSG_KEYS")
    key = MSG_KEYS.Text
        
    msglist.add msg

    Select Case Trim(MSG_STAT.Text)
        Case MsgHide: '0
        
        Case MsgShow: '2
            res = DisplayMessages
            If res <> "" Then
                processmessage = res
                Exit Function
           End If
    End Select

   Select Case UCase(key)
        Case TellerKey:  'teller
           
        Case ChiefKey:  'chief teller
            res = L2ChiefKey(Nothing)
            
        Case ManagerKey:  'manager
            res = L2ManagerKey(Nothing)
        End Select
    If res <> "" Then
        Dim resultdocument As New MSXML2.DOMDocument30
        resultdocument.LoadXML res
        If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
            Load XMLMessageForm
            Set XMLMessageForm.MessageDocument = resultdocument
            XMLMessageForm.Show vbModal
            processmessage = res
            Exit Function
        End If
        Set resultdocument = Nothing
    End If
    
    
    Exit Function
    
processingError:
    Dim aMsg As String
    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας: " & "processmessage" & vbCrLf
    'MsgBox aMsg, vbCritical, "Λάθος Επεξεργασίας..."
    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
    processmessage = "<MESSAGE><ERROR><LINE>" & error & "</LINE></ERROR></MESSAGE>"
End Function


'Private Function UpdateAreaDocument(inDocument As String) As IXMLDOMElement
'    Set workDocument = New MSXML2.DOMDocument30
'    workDocument.LoadXML inDocument
'    Dim copynode As IXMLDOMNode
'    Dim replacenode As IXMLDOMNode
'    Dim error As String
'
'    On Error GoTo UpdateError
'    'update messages
'    For Each copynode In workDocument.SelectNodes("//ODATA//STRMSG[MSG_CODE/text()!='']")
'        Set replacenode = workDocument.selectSingleNode("//IDATA//STRMSG[MSG_CODE='']")
'        If Not (replacenode Is Nothing) Then
'           replacenode.selectSingleNode("MSG_CODE").Text = copynode.selectSingleNode("MSG_CODE").Text
'           replacenode.selectSingleNode("MSG_KEYS").Text = copynode.selectSingleNode("MSG_KEYS").Text
'           'workDocument.selectSingleNode("//IDATA").replaceChild copynode, replacenode
'        ElseIf replacenode Is Nothing Then 'Error more messages than we can handle
'            error = "Λάθος :ODATA returned more messages than IDATA can handle"
'            Err.Raise 1300, "UpdateAreaDocument", error
'            Exit Function
'        End If
'    Next
'
'    'update timespamp
'    If Not (workDocument.selectSingleNode("//ODATA//LOCK_TIMESTAMP") Is Nothing) Then
'        workDocument.selectSingleNode("//IDATA//LOCK_TIMESTAMP").Text = workDocument.selectSingleNode("//ODATA//LOCK_TIMESTAMP").Text
'    End If
'
'    Set UpdateAreaDocument = workDocument.documentElement
'
'    Set copynode = Nothing
'    Set replacenode = Nothing
'
'    Exit Function
'UpdateError:
'    Dim aMsg As String
'    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας: " & "UpdateAreaDocument" & vbCrLf
'    'MsgBox aMsg, vbCritical, "Λάθος Επεξεργασίας..."
'    LogMsgbox aMsg, vbCritical, "Λάθος Επεξεργασίας...", Err
'    Set UpdateAreaDocument = Nothing ' "<MESSAGE><ERROR><LINE>" & error & "</LINE></ERROR></MESSAGE>"
'End Function

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
Set doc = Nothing

End Sub

