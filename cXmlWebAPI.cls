VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlWebAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public ApiUrl As String
Public WrapperUrl As String

Private Function Execute(URL As String, request As String) As String
    Dim response As String
    Dim xmlhttp As New MSXML2.ServerXMLHTTP60
    xmlhttp.open "POST", URL, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setTimeouts 180000, 180000, 180000, 180000
    xmlhttp.SEND (request)
    response = xmlhttp.responseText
    
    Execute = response
End Function


Private Function ExecuteDirect(action As String, request As String) As String
    Dim URL As String
    URL = ApiUrl & action
    ExecuteDirect = Execute(URL, request)
End Function

Private Function ExecuteWrapped(request As String) As String
    Dim URL As String
    URL = WrapperUrl
    ExecuteWrapped = Execute(URL, request)
End Function

Public Function LoadXML(inDocument As String) As String
    Dim aview As New MSXML2.DOMDocument60
    aview.LoadXML inDocument

    Dim bodynode As IXMLDOMElement, urlnode As IXMLDOMElement, actionnode As IXMLDOMElement
    Dim request As String, response As String
    Dim ResponseDoc As MSXML2.DOMDocument60

    Set bodynode = aview.selectSingleNode("//Body")
    Set actionnode = aview.selectSingleNode("//Action")

    If WrapperUrl <> "" Then
        Set urlnode = aview.createElement("Url")
        urlnode.Text = ApiUrl
        aview.documentElement.insertBefore urlnode, bodynode
        
        request = JSON2XML.FromXML(aview, True)
        response = ExecuteWrapped(request)
    End If

    If response = "" And (Not actionnode Is Nothing) Then
        request = JSON2XML.FromXML2(bodynode, True)
        response = ExecuteDirect(actionnode.Text, request)
    End If

    If response = "" Then
        LoadXML = "<MESSAGE><ERROR><LINE>дем епестяеьам дедолема</LINE></ERROR></MESSAGE>"
    ElseIf Left(response, 1) = "<" Then
        Dim starttag As String, endtag As String, startpos As Integer, endpos As Integer, errormsg As String
        starttag = "<title>"
        endtag = "</title>"
        startpos = InStr(response, starttag)
        endpos = InStr(response, endtag)
        If startpos > 0 Then
            errormsg = Mid(response, startpos + Len(starttag), endpos - (startpos + Len(starttag)))
        Else
            errormsg = response
            errormsg = Replace(errormsg, "</", "")
            errormsg = Replace(errormsg, "<", "")
            errormsg = Replace(errormsg, "/>", "")
            errormsg = Replace(errormsg, ">", "")
        End If
        LoadXML = "<MESSAGE><ERROR><LINE>" & errormsg & "</LINE></ERROR></MESSAGE>"
    ElseIf Left(response, 1) <> "{" Then
        LoadXML = "<MESSAGE><ERROR><LINE>" & response & "</LINE></ERROR></MESSAGE>"
    Else
        Set ResponseDoc = JSON2XML.ToXML(response, "json")
        LoadXML = ResponseDoc.XML
    End If
    
End Function

Public Function XML() As String

End Function

