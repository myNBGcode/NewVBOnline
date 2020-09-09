VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c4EyesComareaHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True




Public Function LoadXML(ComArea As cXmlComArea) As String
    Dim workDocument As MSXML2.DOMDocument60
    Set workDocument = New MSXML2.DOMDocument60
    workDocument.LoadXML ComArea.XML
    Dim astr As String
    astr = ComArea.ParseCall(workDocument.documentElement)
    
    If astr <> "" Then
        astr = Replace(astr, Chr(7), "")
    Else
        astr = "<ERROR></ERROR>"
    End If
    workDocument.LoadXML astr
    If Not (workDocument.parseError Is Nothing) Then
        LoadXML = astr
    End If
    
    Dim rc As String, rc_pgm As String, rc_text As String
    
    If Not (workDocument.selectSingleNode("//RESP/RC") Is Nothing) Then
        rc = workDocument.selectSingleNode("//RESP/RC").Text
    End If
    If Not (workDocument.selectSingleNode("//RESP/RC_PGM ") Is Nothing) Then
        rc_pgm = workDocument.selectSingleNode("//RESP/RC_PGM ").Text
    End If
    If Not (workDocument.selectSingleNode("//RESP/RC_TXT ") Is Nothing) Then
        rc_text = workDocument.selectSingleNode("//RESP/RC_TXT ").Text
    End If
    
    If Is4Eyes(rc, rc_pgm) Then
       Dim res As String
       res = L24EyesKey(rc_text)
       If res <> "" Then
           Dim resultdocument As New MSXML2.DOMDocument30
           resultdocument.LoadXML res
           If (resultdocument.documentElement.SelectNodes("//MESSAGE/ERROR").length > 0) Then
               Load XMLMessageForm
               Set XMLMessageForm.MessageDocument = resultdocument
               XMLMessageForm.Show vbModal
               Set resultdocument = Nothing
               Set workDocument = Nothing
               LoadXML = astr
               Exit Function
           Else
               UpdateAuthData resultdocument.selectSingleNode("//MESSAGE/AUTHUSER").Text, resultdocument.selectSingleNode("//MESSAGE/KEY").Text, ComArea
               Set resultdocument = Nothing
               Set workDocument = Nothing
               
               LoadXML = LoadXML(ComArea)
               Exit Function
           End If
           
           Set resultdocument = Nothing
           Set workDocument = Nothing
       End If
    
    Else
        LoadXML = astr
        Set workDocument = Nothing
        Exit Function
    End If
    LoadXML = ""
        
End Function

Private Sub UpdateAuthData(AuthUser As String, authKey, ComArea As cXmlComArea) 'As cXmlComArea
        Dim elm As IXMLDOMElement
        Dim root As IXMLDOMElement
        Set root = ComArea.workDocument.selectSingleNode("//")
        
        
    If Not (ComArea.workDocument Is Nothing) Then
        If (ComArea.workDocument.selectSingleNode("//NT_HEADER") Is Nothing) Then
            Set elm = root.ownerDocument.createElement("NT_HEADER")
            root.appendChild elm
            
        End If
        If (ComArea.workDocument.selectSingleNode("//NT_HEADER/AUTHORISATION") Is Nothing) Then
            Set root = ComArea.workDocument.selectSingleNode("//NT_HEADER")
            Set elm = root.ownerDocument.createElement("AUTHORISATION")
            root.appendChild elm
            
        End If
        If (ComArea.workDocument.selectSingleNode("//NT_HEADER/AUTHORISATION/AUTH_USER") Is Nothing) Then
          Set root = ComArea.workDocument.selectSingleNode("//NT_HEADER/AUTHORISATION")
            Set elm = root.ownerDocument.createElement("AUTH_USER")
            root.appendChild elm
        End If
        ComArea.workDocument.selectSingleNode("//NT_HEADER/AUTHORISATION/AUTH_USER").Text = AuthUser
    End If
        
End Sub
