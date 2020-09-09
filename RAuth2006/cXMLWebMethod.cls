VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLWebMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public content As IXMLDOMElement
Public owner As cXMLWebLink
Public methodname As String
Public namespace As String
Public isDocument As Boolean

Private lPNameArray()
Private lPValueArray()
Private xmlhttp As MSXML2.ServerXMLHTTP30

Public soapDoc As MSXML2.DOMDocument60
Public soapResponse As MSXML2.DOMDocument30
Public Rowsets As MSXML2.IXMLDOMNodeList
Public Rowset As MSXML2.IXMLDOMElement
Public failed As Boolean
Public error As String

Public Function GetParamIndex(ParamName) As Integer
    Dim i As Integer
    GetParamIndex = -1
    For i = LBound(lPNameArray) To UBound(lPNameArray)
        If UCase(lPNameArray(i)) = UCase(ParamName) Then
            GetParamIndex = i: Exit Function
        End If
    Next i
End Function

Public Property Let PNameArray(value)
    lPNameArray = value
End Property

Public Property Get PNameArray()
    PNameArray = lPNameArray
End Property

Public Property Let PValueArray(value)
    lPValueArray = value
End Property

Public Property Get PValueArray()
    PValueArray = lPValueArray
End Property

Public Function AddParam(ParamName As String)
    If GetParamIndex(ParamName) <> -1 Then
        ReDim Preserve lPNameArray(UBound(lPNameArray) + 1)
        lPNameArray(UBound(lPNameArray)) = ParamName
    End If
End Function

Public Function AddParamValue(ParamName As String, ParamValue)
    If GetParamIndex(ParamName) = -1 Then
        ReDim Preserve lPNameArray(UBound(lPNameArray) + 1)
        lPNameArray(UBound(lPNameArray)) = ParamName
    End If
    ReDim Preserve lPValueArray(UBound(lPValueArray) + 1)
    lPValueArray(UBound(lPValueArray)) = ParamValue
End Function

Public Property Get VirtualDirectory() As String
    VirtualDirectory = ""
    If owner Is Nothing Then
        'MsgBox "owner is nothing"
        Exit Property
    End If
    VirtualDirectory = owner.VirtualDirectory
End Property

Public Sub PrepareMethod(Optional aPValueArray)
    Dim sXML As String, i As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
    Dim mNamespace As IXMLDOMAttribute
    
    If Not IsMissing(aPValueArray) Then
        lPValueArray = aPValueArray
    End If
    
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXml sXML
    
    Set mElm = soapDoc.createElement(methodname)
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = namespace
    mElm.setAttributeNode mNamespace
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    If isDocument Then
    Else
        For i = LBound(lPNameArray) To UBound(lPNameArray)
            Set aelm = soapDoc.createElement(lPNameArray(i))
            Set mNamespace = soapDoc.createAttribute("xmlns")
            mNamespace.value = namespace
            aelm.setAttributeNode mNamespace
            aelm.Text = lPValueArray(i)
            mElm.appendChild aelm
        Next i
    End If
End Sub

Public Sub PrepareMethod_usingXml(Paramdoc As IXMLDOMElement)
    Dim sXML As String, i As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
    Dim mNamespace As IXMLDOMAttribute
    
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXml sXML
    
    Set mElm = soapDoc.createElement(methodname)
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = namespace
    mElm.setAttributeNode mNamespace
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    Dim Node As IXMLDOMNode
    For Each Node In Paramdoc.childNodes
        If Node.nodeType = NODE_ELEMENT Then
            Set aelm = soapDoc.createElement(Node.baseName)
            Set mNamespace = soapDoc.createAttribute("xmlns")
            mNamespace.value = namespace
            aelm.setAttributeNode mNamespace
            aelm.Text = Node.Text
            mElm.appendChild aelm
        End If
    Next Node
End Sub

Public Sub PrepareDocumentMethod(Optional inDocument)
    Dim sXML As String, i As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
    Dim mNamespace As IXMLDOMAttribute
    
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument60
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXml sXML
    
    Set anode = soapDoc.selectSingleNode("//*[local-name()='Body' and namespace-uri()='http://schemas.xmlsoap.org/soap/envelope/']")
    If isDocument Then
        If Not IsMissing(inDocument) Then
            If Not inDocument Is Nothing Then
                anode.appendChild inDocument.documentElement
            End If
        End If
    Else
    End If
    
    Set mElm = Nothing
    Set aelm = Nothing
    Set anode = Nothing
    Set mNamespace = Nothing
End Sub

Public Function RowsetCount() As Long
    RowsetCount = Rowsets.length
End Function

Public Function SelectRowset(Index As Long) As MSXML2.IXMLDOMElement
    Set Rowset = Nothing
    If Rowsets.length < Index Then Exit Function
    Set Rowset = Rowsets.Item(Index - 1)
    Set SelectRowset = Rowset
End Function

Public Function Cols() As Long
    Cols = 0
    If Rowset Is Nothing Then Exit Function
    If Rowset.childNodes.length = 0 Then Exit Function
    Cols = Rowset.childNodes.Item(0).childNodes.length
End Function

Public Function Rows() As Long
    Rows = 0
    If Rowset Is Nothing Then Exit Function
    Rows = Rowset.childNodes.length
End Function

Public Function TextMatrix(Row, col) As String
    TextMatrix = ""
    If Rowset Is Nothing Then Exit Function
    If Rowset.childNodes.length < Row Then Exit Function
    If Rowset.childNodes.Item(Row - 1).childNodes.length < col Then Exit Function
    TextMatrix = Rowset.childNodes.Item(Row - 1).childNodes.Item(col - 1).Text
End Function

Public Function RowField(Row, fieldname) As String
    Dim afld
    RowField = ""
    If Rowset Is Nothing Then Exit Function
    If Rowset.childNodes.length < Row Then Exit Function
    Set afld = Rowset.childNodes.Item(Row - 1).selectSingleNode(fieldname)
    If afld Is Nothing Then Exit Function
    RowField = afld.Text
End Function

Public Function ExecMethod() As cXMLWebMethod
    error = "": failed = True: Dim Line As Integer
    On Error GoTo GenericException
    Line = 1
    Set ExecMethod = Me
    Line = 2
    Set xmlhttp = New MSXML2.ServerXMLHTTP
    Line = 3
    Set soapResponse = Nothing
    Line = 4
    Set Rowsets = Nothing
    Line = 5
    Set Rowset = Nothing
    Line = 6
    xmlhttp.setTimeouts 0, 300000, 300000, 0
    Line = 7
    
    'xmlhttp.open "POST", "http://W000010026/WebService1/Service1.asmx", False
    'xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    'xmlhttp.setRequestHeader "SOAPAction", "TestMethod2"
    'xmlhttp.SEND soapDoc.XML
    
    xmlhttp.Open "POST", VirtualDirectory, False
    Line = 8
    xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    Line = 9
    If isDocument Then
        Line = 10
        xmlhttp.setRequestHeader "SOAPAction", methodname 'namespace &
    Else
        Line = 11
        xmlhttp.setRequestHeader "SOAPAction", methodname 'namespace & methodname
    End If
    
    Line = 12
    SaveXmlFile "Request.xml", soapDoc
    'soapDoc.save "C:\request.xml"
    'soapDoc.save NetworkHomeDir() & "\" & "request.xml"
    Line = 13
    On Error GoTo SoapSendFailed
    
    Line = 14
    xmlhttp.send soapDoc.XML
    Line = 15
    On Error GoTo GenericException
    Line = 16
    Set soapResponse = Nothing
    Line = 17
    Set soapResponse = New MSXML2.DOMDocument30
    Line = 18
    Set soapResponse = xmlhttp.responseXML
    
    
'- <SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/">
'- <SOAP-ENV:Body>
'- <SOAP-ENV:Fault>
'  <faultcode>SOAP-ENV:Server</faultcode>
'  <faultstring>Runtime errors.</faultstring>
'  </SOAP-ENV:Fault>
'  </SOAP-ENV:Body>
'  </SOAP-ENV:Envelope>
  
    Line = 19
    Dim aErr As IXMLDOMElement
    Line = 20
    Dim aMsg As String
    Line = 21
    Dim soapFaultDocument As MSXML2.DOMDocument30
    Line = 22
    If soapResponse.XML = "" Then
        Line = 23
        Set soapFaultDocument = New MSXML2.DOMDocument30
        Line = 24
        aMsg = "<MESSAGE><ERROR><LINE>ΔΕΝ ΥΠΑΡΧΕΙ ΑΠΑΝΤΗΣΗ ΑΠΟ ΤΟ ΣΥΣΤΗΜΑ... Status: " & xmlhttp.Status & " " & xmlhttp.StatusText & "</LINE><EXCEPTIONTYPE></EXCEPTIONTYPE></ERROR></MESSAGE>"
        Line = 25
        soapFaultDocument.LoadXml aMsg
        Line = 26
        soapFaultDocument.documentElement.selectSingleNode("//EXCEPTIONTYPE").Text = "ΤΟ soapResponse.xml είναι κενό..."
        Line = 27
        Load XMLMessageForm
        Line = 28
        Set XMLMessageForm.MessageDocument = soapFaultDocument
        Line = 29
        'Open "c:\xmlhttpfailure.html" For Output As #1
        Open NetworkHomeDir & "\xmlhttpfailure.html" For Output As #1
        Line = 30
        Print #1, xmlhttp.responseText
        Line = 31
        Close #1
        Line = 32
        XMLMessageForm.Show vbModal
        Line = 33
        soapResponse.LoadXml soapFaultDocument.XML
        Line = 34
        Set soapFaultDocument = Nothing
        Line = 35
        Exit Function
    End If
    Line = 36
    Set aErr = soapResponse.documentElement.selectSingleNode("//" & soapResponse.documentElement.prefix & ":Fault")
    Line = 37
    If Not (aErr Is Nothing) Then
        Line = 38
        'soapResponse.save "C:\response.xml"
        SaveXmlFile "Response.xml", soapResponse
        Line = 39
        If Not aErr.childNodes Is Nothing Then
            Line = 40
            If aErr.childNodes.length > 0 Then
                Line = 41
                Set soapFaultDocument = New MSXML2.DOMDocument30
                Line = 42
                aMsg = "<MESSAGE><ERROR><LINE>" & aErr.selectSingleNode("faultcode").Text & "</LINE><EXCEPTIONTYPE></EXCEPTIONTYPE></ERROR></MESSAGE>"
                Line = 43
                soapFaultDocument.LoadXml aMsg
                Line = 44
                soapFaultDocument.documentElement.selectSingleNode("//EXCEPTIONTYPE").Text = aErr.selectSingleNode("faultstring").Text
                Line = 45
                Load XMLMessageForm
                Line = 46
                Set XMLMessageForm.MessageDocument = soapFaultDocument
                Line = 47
                XMLMessageForm.Show vbModal
                Line = 48
                Set soapFaultDocument = Nothing
                Line = 49
                Exit Function
            End If
        End If
    Else
        Line = 50
        Dim errorDocument As New MSXML2.DOMDocument30
        Line = 51
        If (soapResponse.documentElement.selectNodes("//MESSAGE/ERROR").length > 0) Then
            Line = 52
            errorDocument.LoadXml soapResponse.documentElement.selectSingleNode("//MESSAGE").XML
            Line = 53
            Load XMLMessageForm
            Line = 54
            Set XMLMessageForm.MessageDocument = errorDocument
            Line = 55
            XMLMessageForm.Show vbModal
            Line = 56
            Set errorDocument = Nothing
            Line = 57
            'soapResponse.save "C:\response.xml"
            SaveXmlFile "Response.xml", soapResponse
            Line = 58
            Exit Function
        End If
        Line = 59
        If (soapResponse.documentElement.selectNodes("//MESSAGE/WARNING").length > 0) Then
            Line = 60
            errorDocument.LoadXml soapResponse.documentElement.selectSingleNode("//MESSAGE").XML
            Line = 61
            Load XMLMessageForm
            Line = 62
            Set XMLMessageForm.MessageDocument = errorDocument
            Line = 63
            XMLMessageForm.Show vbModal
            Line = 64
            Set errorDocument = Nothing
        End If
        Line = 65
        If (soapResponse.documentElement.selectNodes("//RESULT/RETURNCODE").length > 0) Then
            If (soapResponse.documentElement.selectSingleNode("//RESULT/RETURNCODE").Text <> "1") Then
                Line = 66
                    If (soapResponse.documentElement.selectNodes("//RESULT/ERROR/MESSAGE").length > 0) Then
                    errorDocument.LoadXml soapResponse.documentElement.selectSingleNode("//RESULT/ERROR/MESSAGE").XML
                    Line = 67
                    Load XMLMessageForm
                    Line = 68
                    Set XMLMessageForm.MessageDocument = errorDocument
                    Line = 69
                    XMLMessageForm.Show vbModal
                    Line = 70
                    Set errorDocument = Nothing
                End If
            End If
        End If
        
        Set errorDocument = Nothing
    End If
    
    Set aErr = Nothing
    Set soapFaultDocument = Nothing
    
    On Error Resume Next
    Line = 71
    'soapResponse.save "C:\response.xml"
    SaveXmlFile "Response.xml", soapResponse
    Line = 72
    failed = False: Exit Function
    
Dim genErrorDocument As New MSXML2.DOMDocument30

SoapSendFailed:
    genErrorDocument.LoadXml "<MESSAGE><ERROR><LINE>ΑΠΕΤΥΧΕ Η ΕΠΙΚΟΙΝΩΝΙΑ " & " " & Err.Number & " " & Err.Description & " </LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    GoTo ExceptionExitPos
GenericException:
    Dim errstring As String
    errstring = "ΑΠΕΤΥΧΕ Η ΣΥΝΑΛΛΑΓΗ " & " " & Err.source & " Line: " & Line & " " & Err.Number & " " & Err.Description
    If Line = 7 Then errstring = errstring & " VirtualDirectory: " & VirtualDirectory
    genErrorDocument.LoadXml "<MESSAGE><ERROR><LINE>" & errstring & " </LINE><EXCEPTIONTYPE/></ERROR></MESSAGE>"
    GoTo ExceptionExitPos
ExceptionExitPos:
    Load XMLMessageForm
    Set XMLMessageForm.MessageDocument = genErrorDocument
    XMLMessageForm.Show vbModal
    Set genErrorDocument = Nothing
End Function

Private Sub SetHeaderAttribute(Node As IXMLDOMElement, attributename As String, attributevalue As String)
    Dim aattribute As IXMLDOMAttribute
    Set aattribute = Node.Attributes.getNamedItem(attributename)
    If aattribute Is Nothing Then
        Set aattribute = Node.ownerDocument.createAttribute(attributename)
        Node.Attributes.setNamedItem aattribute
    End If
    If attributename = "username" Then
        aattribute.Text = attributevalue
    ElseIf aattribute.Text = "" Then
        aattribute.Text = attributevalue
    End If
End Sub

Private Function SetHeaderElement(Node As IXMLDOMElement, namespace As String, elementname As String, elementvalue As String) As IXMLDOMElement
    Dim element As IXMLDOMElement, namespaceattribute As IXMLDOMAttribute
    Set element = Node.selectSingleNode("//*[local-name()='" & elementname & "' and  namespace-uri()='" & namespace & "']")
    If element Is Nothing Then
        Set element = Node.ownerDocument.createElement(elementname)
        Node.appendChild element
        Set namespaceattribute = Node.ownerDocument.createAttribute("xmlns")
        namespaceattribute.value = namespace
        element.setAttributeNode namespaceattribute
        element.Text = elementvalue
    End If
    If element.Text = "" Then element.Text = elementvalue
    Set SetHeaderElement = element
End Function

Public Function LoadXml(value As String) As String
    Dim aview As New MSXML2.DOMDocument60
    aview.LoadXml value
    
    Dim anode As IXMLDOMElement, cnode As IXMLDOMElement, dnode As IXMLDOMElement
    
    Set anode = aview.selectSingleNode("//*[local-name()='MESSAGE' and  namespace-uri()='http://www.nbg.gr/online/msg']/*[local-name()='MESSAGEHEADER' and  namespace-uri()='http://www.nbg.gr/online/msg' ]")
    If Not (anode Is Nothing) Then
        SetHeaderAttribute anode, "applnamespace", namespace
        If anode.Attributes.getNamedItem("operation") Is Nothing Then
        End If
        SetHeaderAttribute anode, "branch", cBRANCH
        SetHeaderAttribute anode, "computername", MachineName
        SetHeaderAttribute anode, "username", cUserName
        SetHeaderAttribute anode, "workenvironment", WorkEnvironment
    End If
    
    Set anode = aview.selectSingleNode("//*[local-name()='MESSAGE' and  namespace-uri()='http://www.nbg.gr/online/obj']/*[local-name()='MESSAGEHEADER' and  namespace-uri()='http://www.nbg.gr/online/obj']")
    If Not (anode Is Nothing) Then
        Set cnode = SetHeaderElement(anode, "http://www.nbg.gr/online/obj", "CLIENT_DATA", "")
        If Not (cnode Is Nothing) Then
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "USERNAME", cUserName
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "COMPUTERNAME", MachineName
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "SYSTEM", "RAUTH"
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "BRANCH", cBRANCH
        End If
        Set dnode = SetHeaderElement(anode, "http://www.nbg.gr/online/obj", "AUTHORISATION", "")
        If Not (dnode Is Nothing) Then
            SetHeaderElement dnode, "http://www.nbg.gr/online/obj", "AUTHUSER", ""
            SetHeaderElement dnode, "http://www.nbg.gr/online/obj", "CODTX", ""
        End If
    End If
    
    On Error Resume Next
    Screen.ActiveForm.Status.Panels(1).Text = "Διαβίβαση. Παρακαλώ Περιμένετε...."
    Screen.ActiveForm.StatusBar.SimpleText = "Διαβίβαση. Παρακαλώ Περιμένετε...."
    On Error GoTo 0
    If isDocument Then
        owner.ExecDocumentMethod methodname, aview
    Else
        owner.ExecMethod_usingXml methodname, aview.documentElement
    End If
    
    Set aview = Nothing
    Set anode = Nothing
    Set cnode = Nothing
    Set dnode = Nothing

    On Error Resume Next
        Screen.ActiveForm.Status.Panels(1).Text = ""
        Screen.ActiveForm.StatusBar.SimpleText = ""
    On Error GoTo 0
    
    If soapResponse Is Nothing Then
        LoadXml = ""
    Else
        LoadXml = soapResponse.XML
        LoadXml = Replace(LoadXml, "xsi:nil=""true""", "")
    End If
End Function

Public Function LoadXmlNoTrnUpdate(value As String) As String
    Dim aview As New MSXML2.DOMDocument60
    aview.LoadXml value
    
    Dim anode As IXMLDOMElement, cnode As IXMLDOMElement, dnode As IXMLDOMElement
        
    Set anode = aview.selectSingleNode("//*[local-name()='MESSAGE' and  namespace-uri()='http://www.nbg.gr/online/msg']/*[local-name()='MESSAGEHEADER' and  namespace-uri()='http://www.nbg.gr/online/msg' ]")
    If Not (anode Is Nothing) Then
        SetHeaderAttribute anode, "applnamespace", namespace
        If anode.Attributes.getNamedItem("operation") Is Nothing Then
        End If
        SetHeaderAttribute anode, "branch", cBRANCH
        SetHeaderAttribute anode, "computername", MachineName
        SetHeaderAttribute anode, "username", cUserName
        SetHeaderAttribute anode, "workenvironment", WorkEnvironment
    End If
    
    Set anode = aview.selectSingleNode("//*[local-name()='MESSAGE' and  namespace-uri()='http://www.nbg.gr/online/obj']/*[local-name()='MESSAGEHEADER' and  namespace-uri()='http://www.nbg.gr/online/obj']")
    If Not (anode Is Nothing) Then
        Set cnode = SetHeaderElement(anode, "http://www.nbg.gr/online/obj", "CLIENT_DATA", "")
        If Not (cnode Is Nothing) Then
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "USERNAME", cUserName
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "COMPUTERNAME", MachineName
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "SYSTEM", "RAUTH"
            SetHeaderElement cnode, "http://www.nbg.gr/online/obj", "BRANCH", cBRANCH
        End If
        Set dnode = SetHeaderElement(anode, "http://www.nbg.gr/online/obj", "AUTHORISATION", "")
        If Not (dnode Is Nothing) Then
            SetHeaderElement dnode, "http://www.nbg.gr/online/obj", "AUTHUSER", ""
            SetHeaderElement dnode, "http://www.nbg.gr/online/obj", "CODTX", ""
        End If
    End If
    
    On Error Resume Next
    Screen.ActiveForm.Status.Panels(1).Text = "Διαβίβαση. Παρακαλώ Περιμένετε...."
    Screen.ActiveForm.StatusBar.SimpleText = "Διαβίβαση. Παρακαλώ Περιμένετε...."
    On Error GoTo 0
    If isDocument Then
        owner.ExecDocumentMethod methodname, aview
    Else
        owner.ExecMethod_usingXml methodname, aview.documentElement
    End If
    
    Set aview = Nothing
    Set anode = Nothing
    Set cnode = Nothing
    Set dnode = Nothing

    On Error Resume Next
        Screen.ActiveForm.Status.Panels(1).Text = ""
        Screen.ActiveForm.StatusBar.SimpleText = ""
    On Error GoTo 0
    
    If soapResponse Is Nothing Then
        LoadXmlNoTrnUpdate = ""
    Else
        LoadXmlNoTrnUpdate = soapResponse.XML
        LoadXmlNoTrnUpdate = Replace(LoadXmlNoTrnUpdate, "xsi:nil=""true""", "")
    End If
End Function
Public Function XML() As String
    If soapResponse Is Nothing Then
        XML = ""
    Else
        XML = soapResponse.XML
        XML = Replace(XML, "xsi:nil=""true""", "")
    End If
End Function

Public Sub PrepareDocumentMethod_Old(Optional inDocument)
    Dim sXML As String, i As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
    Dim mNamespace As IXMLDOMAttribute

    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"

    soapDoc.LoadXml sXML

    'Set anode = soapDoc.selectSingleNode("//*[local-name()='Body' and namespace-uri()='http://schemas.xmlsoap.org/soap/envelope/']")
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    If isDocument Then
        If Not IsMissing(inDocument) Then
            If Not inDocument Is Nothing Then
                anode.appendChild inDocument.documentElement
            End If
        End If
    Else
    End If
End Sub


