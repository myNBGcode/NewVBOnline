VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlKAAMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public owner As cXMLDocumentManager

Public content As IXMLDOMElement
Public workDocument As MSXML2.DOMDocument30
'definition
'<kaamessage name="test" id="messagename" >
'   <webdocumentmethod name="APILink" namespace="http://www.nbg.gr/online/KAAServer/" virtualdirectory="KAAServer_WebLink"></method>
'</kaamessage>
'
'call
'<input>
'   <structurename1>
'       <structurename2>
'           <fieldname>fieldvalue</fieldname> ή εναλλακτικά
'           <fieldname value="fieldvalue"></fieldname>
'      </structurename2>
'   </structurename1>
'</input>
'<function name="test">
'</function>

Public Function ParseCall(callElm As IXMLDOMElement) As String
Dim name As String
Dim messageid As String
Dim methodname As String

    Dim methodelm As IXMLDOMElement
    Dim nameattr As IXMLDOMAttribute
    Dim messageidattr As IXMLDOMAttribute
    Dim methodnameattr As IXMLDOMAttribute
    
    Set methodelm = content.selectSingleNode("./webdocumentmethod")
    If Not methodelm Is Nothing Then
        Set methodnameattr = methodelm.Attributes.getNamedItem("name")
        If Not methodnameattr Is Nothing Then methodname = methodnameattr.value
        
        Set nameattr = content.Attributes.getNamedItem("name")
        If Not (nameattr Is Nothing) Then name = nameattr.value
        Set messageidattr = content.Attributes.getNamedItem("messagename")
        If Not (messageidattr Is Nothing) Then messageid = messageidattr.value
        
        Dim Message As New cKAAMessage
        Message.Prepare messageid
        
        If Not Message.UpdateValues(Message.Document.documentElement, callElm) Then Exit Function
        Dim WebMethod As cXMLWebMethod
        Set WebMethod = owner.namedxmlObject(methodname)
        WriteJournalBeforeSend Message.Document.documentElement
        ParseCall = WebMethod.LoadXML(Message.Document.XML)
        WriteJournalAfterReceive ParseCall
    Else
        ParseCall = ""
    End If
End Function

Public Sub WriteJournalBeforeSend(elm As IXMLDOMElement)
    Dim list As IXMLDOMNodeList
    Set list = elm.SelectNodes("//request[@journaltitle!='']|//request/field[@journaltitle!='' and @value!='' ]|//columns/column[@journaltitle!='' and @value!='' ]")
    If list.length > 0 Then
        UpdateTrnNum_
        Dim item As IXMLDOMElement
        For Each item In list
            Dim atitle As String
            Dim avalue As String
            
            atitle = item.getAttribute("journaltitle")
            If IsNull(item.getAttribute("value")) Then
                eJournalWriteAll Nothing, atitle
            Else
                avalue = item.getAttribute("value")
                eJournalWriteFld Nothing, 0, atitle & ":", avalue
            End If
        Next item
    End If
End Sub

Public Sub WriteJournalAfterReceive(data As String)
    Dim anewdoc As New MSXML2.DOMDocument30
    anewdoc.LoadXML data
    Dim list As IXMLDOMNodeList
    Set list = anewdoc.SelectNodes("//response[@journaltitle!='']|//response//field[@journaltitle!='' and @value!='' ]")
    If list.length > 0 Then
        Dim item As IXMLDOMElement
        For Each item In list
            Dim atitle As String
            Dim avalue As String
            
            atitle = item.getAttribute("journaltitle")
            If IsNull(item.getAttribute("value")) Then
                eJournalWriteAll Nothing, atitle
            Else
                avalue = item.getAttribute("value")
                eJournalWriteFld Nothing, 0, atitle & ":", avalue
            End If
        Next item
    End If
End Sub

Public Function LoadXML(inDocument As String) As String

    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim astr As String
    astr = ParseCall(workDocument.documentElement)
    If astr <> "" Then
        LoadXML = astr
    Else
        LoadXML = "<ERROR></ERROR>"
    End If
    workDocument.LoadXML LoadXML
End Function

Public Function XML() As String
    If workDocument Is Nothing Then
        XML = ""
    Else
        XML = workDocument.XML
    End If
End Function




