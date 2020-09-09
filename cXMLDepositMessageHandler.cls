VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLDepositMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public workDocument As MSXML2.DOMDocument30
Public Manager As cXMLDocumentManager

Private Comarea As cXmlComArea
Public content As IXMLDOMElement

Dim msglist As New Collection
Dim doc As New MSXML2.DOMDocument30

Public Function LoadXml(inDocument As String) As String

    'Dim response As New MSXML2.DOMDocument30
    Dim astr As String, comareaid As String, Key As String, res As String
    Dim elm As IXMLDOMElement, aItem As IXMLDOMNode
    Dim msg As IXMLDOMElement
    Dim area As IXMLDOMAttribute
    Dim error As IXMLDOMElement

    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXml inDocument
    
    If Trim(workDocument.selectSingleNode("//RESP/RC").Text) <> "0" Then
         LoadXml = inDocument
        'LoadXml = "<MESSAGE><ERROR>" + workDocument.selectSingleNode("//RESP/RC_TXT").Text + "</MESSAGE></ERROR>"
        Exit Function
    End If
    
   If (doc.documentElement Is Nothing) Then doc.appendChild doc.createElement("root")
    
  
    
    Set msg = workDocument.selectSingleNode("//ODATA/STRMSG")
    If Trim(msg.selectSingleNode("MSG_CODE").Text) <> "" Then
       
        'HandleMessage msg
        Dim MSG_STAT As IXMLDOMElement
        Dim MSG_KEYS  As IXMLDOMElement
        
        Set MSG_STAT = msg.selectSingleNode("MSG_STAT")
        Set MSG_KEYS = msg.selectSingleNode("MSG_KEYS")
        Key = MSG_KEYS.Text
        
        
'        Select Case UCase(Key)
'        Case "3": 'teller
'
'        Case "7": 'chief teller
'            res = L2ChiefKey(Nothing)
'
'        Case "B": 'manager
'            res = L2ManagerKey(Nothing)
'        End Select
'        If res <> "" Then
''            ShowError res
''            LoadXml = res
''            Exit Function
'            Dim resultdocument As New MSXML2.DOMDocument30
'            resultdocument.LoadXml res
'            If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
'                Load XMLMessageForm
'                Set XMLMessageForm.MessageDocument = resultdocument
'                XMLMessageForm.Show vbModal
'                LoadXml = res
'                Exit Function
'            End If
'        End If
        
       
        If Trim(MSG_STAT.Text) = "2" Then
            
        
            If msglist.count > 0 Then
                msglist.add msg
                
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
                   msglist.Remove (1)
                Wend
                res = L2ShowDepositMessages(doc.documentElement)
                If res <> "" Then
                    ShowError res
                    LoadXml = res
                    Exit Function
                End If

            Else
                res = L2ShowDepositMessages(msg)
                If res <> "" Then
                   ShowError res
                   LoadXml = res
                   Exit Function
                End If
            End If
            
            
        ElseIf Trim(MSG_STAT.Text) = "0" Then
            msglist.add msg
        End If
        
'****Εγκριση
        
        Select Case UCase(Key)
        Case "3": 'teller
           
        Case "7": 'chief teller
            res = L2ChiefKey(Nothing)
            
        Case "B": 'manager
            res = L2ManagerKey(Nothing)
        End Select
        If res <> "" Then
'            ShowError res
'            LoadXml = res
'            Exit Function
            Dim resultdocument As New MSXML2.DOMDocument30
            resultdocument.LoadXml res
            If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
                Load XMLMessageForm
                Set XMLMessageForm.MessageDocument = resultdocument
                XMLMessageForm.Show vbModal
                LoadXml = res
                Exit Function
            End If
        End If
        
'****

        Set area = content.Attributes.getNamedItem("com")
        If Not (area Is Nothing) Then
            Set Comarea = Manager.XmlObjectList.item(area.Text)
        
            If Not (Comarea Is Nothing) Then
               
               comareaid = Comarea.content.Attributes.getNamedItem("id").Text
               Set elm = UpdateAreaDocument(inDocument)
               inDocument = Comarea.DepositSendArea(elm, comareaid)
               'astr = "<MESSAGE>" + inDocument + "</MESSAGE>"
               
               LoadXml = LoadXml(inDocument)
               Exit Function
            End If
    End If
    
    Else
        LoadXml = inDocument
        Exit Function
    End If
    
    

End Function


Private Function UpdateAreaDocument(inDocument As String) As IXMLDOMElement
    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXml inDocument
    
    Dim copynode As IXMLDOMNode
    Set copynode = workDocument.selectSingleNode("//ODATA/STRMSG")
    Dim replacenode As IXMLDOMNode
    Dim code As String
    For Each replacenode In workDocument.SelectNodes("//IDATA/STRMSG")
        code = replacenode.selectSingleNode("MSG_CODE").Text
        If code = "" Then
       
            Exit For
        End If
        
    Next
       
    If Not (replacenode Is Nothing) Then workDocument.selectSingleNode("//IDATA").replaceChild copynode, replacenode
    Set UpdateAreaDocument = workDocument.documentElement
End Function

Private Function ShowError(error As String) As String
    Dim resultdocument As New MSXML2.DOMDocument30
    resultdocument.LoadXml error
    If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
        Load XMLMessageForm
        Set XMLMessageForm.MessageDocument = resultdocument
        XMLMessageForm.Show vbModal
        ShowError = error
        Exit Function
    End If
End Function
