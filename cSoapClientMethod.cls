VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSoapClientMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public owner As CSoapClient
Public methodname As String
Private lPNameArray()
Private lPValueArray()
Private xmlhttp As MSXML2.ServerXMLHTTP30

Public soapDoc As MSXML2.DOMDocument30
Public soapResponse As MSXML2.DOMDocument30
Public Rowsets As MSXML2.IXMLDOMNodeList
Public Rowset As MSXML2.IXMLDOMElement
Public error As String

Public Function GetParamIndex(ParamName) As Integer
    Dim I As Integer
    GetParamIndex = -1
    For I = LBound(lPNameArray) To UBound(lPNameArray)
        If UCase(lPNameArray(I)) = UCase(ParamName) Then
            GetParamIndex = I: Exit Function
        End If
    Next I
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
    If owner Is Nothing Then Exit Property
    VirtualDirectory = owner.VirtualDirectory
End Property

Public Sub PrepareMethod(Optional aPValueArray)
    Dim sXML As String, I As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
    
    If Not IsMissing(aPValueArray) Then
        lPValueArray = aPValueArray
    End If
    
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXML sXML
    
    Set mElm = soapDoc.createElement(methodname)
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    For I = LBound(lPNameArray) To UBound(lPNameArray)
        Set aelm = soapDoc.createElement(lPNameArray(I))
        aelm.Text = lPValueArray(I)
        mElm.appendChild aelm
    Next I
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

Public Function ExecMethod() As cSoapClientMethod
    error = ""
    Set xmlhttp = New MSXML2.ServerXMLHTTP
    Set soapResponse = Nothing
    Set Rowsets = Nothing
    Set Rowset = Nothing
    xmlhttp.open "POST", VirtualDirectory, False
    xmlhttp.setRequestHeader "Content-Type", "text/xml"
    xmlhttp.setRequestHeader "SOAPAction", VirtualDirectory & "/" & methodname
    SaveXmlFile "Request.xml", soapDoc
    
    eJournalWrite "cSoapClientMethod.ExecMethod Send : " & format(Time, "HH:MM:SS")
    xmlhttp.SEND soapDoc.xml
    eJournalWrite "cSoapClientMethod.ExecMethod Receive : " & format(Time, "HH:MM:SS")
    Set soapResponse = Nothing
    Set soapResponse = New MSXML2.DOMDocument30
    Set soapResponse = xmlhttp.responseXML
    
    
'- <SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/">
'- <SOAP-ENV:Body>
'- <SOAP-ENV:Fault>
'  <faultcode>SOAP-ENV:Server</faultcode>
'  <faultstring>Runtime errors.</faultstring>
'  </SOAP-ENV:Fault>
'  </SOAP-ENV:Body>
'  </SOAP-ENV:Envelope>
  
    Set ExecMethod = Me
    
    Dim aErr As IXMLDOMElement
    Set aErr = soapResponse.documentElement.selectSingleNode("//SOAP-ENV:Fault")
    If Not (aErr Is Nothing) Then
        If Not aErr.childNodes Is Nothing Then
            If aErr.childNodes.length > 0 Then
                Dim aMsg As String
                Dim aErrLine As IXMLDOMElement
                
                For Each aErrLine In aErr.childNodes
                    aMsg = aMsg & "( " & aErrLine.nodename & " ):" & aErrLine.Text & vbCrLf
                Next
            
                LogMsgbox aMsg, vbCritical, "Εφαρμογή On Line"
                error = aMsg
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Exit Function
            End If
        End If
    End If
    
    Set Rowsets = soapResponse.documentElement.SelectNodes("//rowset")
    If Rowsets.length > 0 Then Set Rowset = Rowsets(0)
    SaveXmlFile "Response.xml", soapResponse
End Function

Public Sub PrepareSQLXmlMethod(Optional aPValueArray)
    Dim sXML As String, I As Integer, mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode
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
    soapDoc.LoadXML sXML
    
    Set mElm = soapDoc.createElement(methodname) ' "ExecuteSP") 'methodname
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/SQLXml/"
    mElm.setAttributeNode mNamespace
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    
    For I = LBound(lPNameArray) To UBound(lPNameArray)
        Set aelm = soapDoc.createElement(lPNameArray(I))
        Set mNamespace = soapDoc.createAttribute("xmlns")
        mNamespace.value = "http://www.nbg.gr/online/SQLXml/"
        aelm.setAttributeNode mNamespace
        aelm.Text = lPValueArray(I)
        mElm.appendChild aelm
    Next I
End Sub

Public Function ExecSQLXmlMethod() As cSoapClientMethod
    error = ""
    Set xmlhttp = New MSXML2.ServerXMLHTTP
    Set soapResponse = Nothing
    Set Rowsets = Nothing
    Set Rowset = Nothing
    xmlhttp.setTimeouts 0, 300000, 300000, 0
    xmlhttp.open "POST", VirtualDirectory, False
    xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    xmlhttp.setRequestHeader "SOAPAction", "ExecuteSP" 'methodname
    SaveXmlFile "Request.xml", soapDoc
    eJournalWrite "cSoapClientMethod.ExecSQLXmlMethod Send : " & format(Time, "HH:MM:SS")
    xmlhttp.SEND soapDoc.xml
    eJournalWrite "cSoapClientMethod.ExecSQLXmlMethod Receive : " & format(Time, "HH:MM:SS")
    Set soapResponse = Nothing
    Set soapResponse = New MSXML2.DOMDocument30
    Set soapResponse = xmlhttp.responseXML
    
    
'- <SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/">
'- <SOAP-ENV:Body>
'- <SOAP-ENV:Fault>
'  <faultcode>SOAP-ENV:Server</faultcode>
'  <faultstring>Runtime errors.</faultstring>
'  </SOAP-ENV:Fault>
'  </SOAP-ENV:Body>
'  </SOAP-ENV:Envelope>
  
    Set ExecSQLXmlMethod = Me
    
    Dim aErr As IXMLDOMElement
    Set aErr = soapResponse.documentElement.selectSingleNode("//soap:Fault")
    If Not (aErr Is Nothing) Then
        If Not aErr.childNodes Is Nothing Then
            If aErr.childNodes.length > 0 Then
                Dim aMsg As String
                Dim aErrLine As IXMLDOMElement
                
                For Each aErrLine In aErr.childNodes
                    aMsg = aMsg & "( " & aErrLine.nodename & " ):" & aErrLine.Text & vbCrLf
                Next
            
                LogMsgbox aMsg, vbCritical, "Εφαρμογή On Line"
                error = aMsg
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Exit Function
            End If
        End If
    End If
    
    Set Rowsets = soapResponse.documentElement.SelectNodes("//rowset")
    If Rowsets.length > 0 Then Set Rowset = Rowsets(0)
    SaveXmlFile "Response.xml", soapResponse
End Function


