VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADORecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public TName As String
Public Recordset As New ADODB.Recordset
Attribute Recordset.VB_VarHelpID = -1
Public RecordsetOld As New ADODB.Recordset
Public RecordsetEntry As New RecordsetEntry
Public TVirtualDirectoryName As String
Public TDBName As String
Public Function fields() As ADODB.fields
     Set fields = Me.Recordset.fields
End Function
Public Property Get DBName() As String
      DBName = Me.TDBName
End Property
Public Property Let DBName(value As String)
      Me.TDBName = value
End Property
Public Property Get name() As String
       name = Me.TName
End Property
Public Property Let name(value As String)
       Me.TName = value
End Property
Public Property Get VirtualDirectoryName() As String
       VirtualDirectoryName = Me.TVirtualDirectoryName
End Property
Public Property Let VirtualDirectoryName(value As String)
       Me.TVirtualDirectoryName = value
End Property
Public Property Get RecordCount() As Integer
     RecordCount = Me.Recordset.RecordCount
End Property
Public Property Get Bof() As Boolean
     Bof = Me.Recordset.Bof
End Property
Public Property Get Eof() As Boolean
     Eof = Me.Recordset.Eof
End Property
Public Sub MoveFirst()
     Me.Recordset.MoveFirst
End Sub
Public Sub MoveNext()
     Me.Recordset.MoveNext
End Sub
Public Sub MoveLast()
     Me.Recordset.MoveLast
End Sub
Public Function Open_(inCmd, Optional inCursorType, Optional inLockType) As ADODB.Recordset
    Set Open_ = Me.ExecuteQUERY(inCmd, inCursorType, inLockType)
End Function
Public Sub AddNew()
    Me.Recordset.AddNew
End Sub
Public Sub Update()
    Me.ExecuteUPDATE
End Sub
Public Sub Close_()
    Me.Recordset.Close
End Sub
Public Sub cancelupdate()
   Me.Recordset.cancelupdate
End Sub

Public Function ExecuteQUERY(query, Optional inCursorType, Optional inLockType) As ADODB.Recordset
    
    Dim sXML As String, I As Integer
    Dim mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode As MSXML2.IXMLDOMElement
    Dim mNamespace As IXMLDOMAttribute
    Dim xmlhttp
    Dim responsedoc As MSXML2.DOMDocument30
    Dim soapResponse As MSXML2.DOMDocument30
    Dim soapDoc As MSXML2.DOMDocument30
    If IsMissing(inCursorType) Then
       inCursorType = adOpenStatic
    End If
    If IsMissing(inLockType) Then
       inLockType = adLockReadOnly
    End If
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXML sXML
    
    Set mElm = soapDoc.createElement("ExecuteREQUEST")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    mElm.setAttributeNode mNamespace
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    
    Set aelm = soapDoc.createElement("query")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = query
    anode.appendChild aelm
    
    Set aelm = soapDoc.createElement("dbname")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = Me.DBName
    anode.appendChild aelm
    
    Set aelm = soapDoc.createElement("CursorType")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = CStr(inCursorType)
    anode.appendChild aelm
    
    Set aelm = soapDoc.createElement("LockType")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = CStr(inLockType)
    anode.appendChild aelm
    
    
    
    Set xmlhttp = New MSXML2.ServerXMLHTTP
    Set soapResponse = Nothing
    xmlhttp.setTimeouts 0, 300000, 300000, 0
    xmlhttp.open "POST", Me.VirtualDirectoryName, False
    xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    xmlhttp.setRequestHeader "SOAPAction", "ExecuteREQUEST"
    
        
    SaveXmlFile "request.xml", soapDoc
    
    eJournalWrite "cADORecordset.ExecuteQUERY Send : " & format(Time, "HH:MM:SS")
    xmlhttp.SEND soapDoc.xml
    eJournalWrite "cADORecordset.ExecuteQUERY Receive : " & format(Time, "HH:MM:SS")
    Set soapResponse = New MSXML2.DOMDocument30
    Set soapResponse = xmlhttp.responseXML
    
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
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Exit Function
            End If
        End If
    End If
        
    SaveXmlFile "Response.xml", soapResponse
    
    Set responsedoc = New MSXML2.DOMDocument30
    responsedoc.LoadXML soapResponse.documentElement.selectSingleNode("//ExecuteREQUESTResult").Text
      
    On Error Resume Next
    Me.Recordset.Close
    Me.RecordsetOld.Close
    Me.Recordset.open responsedoc, , ADODB.PersistFormatEnum.adPersistXML, adLockOptimistic
    Me.RecordsetOld.open responsedoc, , ADODB.PersistFormatEnum.adPersistXML, adLockOptimistic
    
    Set Me.RecordsetEntry.rs = Me.Recordset
    
    Set ExecuteQUERY = Me.Recordset
End Function
Public Sub ExecuteUPDATE()
    Dim sXML As String, I As Integer
    Dim mElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, anode As MSXML2.IXMLDOMElement
    Dim mNamespace As IXMLDOMAttribute
    Dim xmlhttp
    Dim responsedoc As MSXML2.DOMDocument30
    Dim soapResponse As MSXML2.DOMDocument30
    Dim soapDoc As MSXML2.DOMDocument30
    
    Set soapDoc = Nothing
    Set soapDoc = New MSXML2.DOMDocument30
    sXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>" & _
            "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" >" & _
            " <SOAP-ENV:Body />" & _
            "</SOAP-ENV:Envelope>"
    soapDoc.LoadXML sXML
    
    Set mElm = soapDoc.createElement("ExecuteUPDATE")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    mElm.setAttributeNode mNamespace
    Set anode = soapDoc.selectSingleNode("//SOAP-ENV:Body")
    anode.appendChild mElm
    
    Dim oldxml As New MSXML2.DOMDocument30
    Me.RecordsetOld.Save oldxml, adPersistXML
    
    Set aelm = soapDoc.createElement("rsOLDString")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = oldxml.xml
    anode.appendChild aelm
    
    Dim newxml As New MSXML2.DOMDocument30
    Me.Recordset.Save newxml, adPersistXML
    Set aelm = soapDoc.createElement("rsNEWString")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = newxml.xml
    anode.appendChild aelm
    
    Set aelm = soapDoc.createElement("dbName")
    Set mNamespace = soapDoc.createAttribute("xmlns")
    mNamespace.value = "http://www.nbg.gr/online/Shine.L2.ADODispatcher/"
    aelm.setAttributeNode mNamespace
    aelm.Text = Me.DBName
    anode.appendChild aelm
    
    Set xmlhttp = New MSXML2.ServerXMLHTTP
    Set soapResponse = Nothing
    xmlhttp.setTimeouts 0, 300000, 300000, 0
    xmlhttp.open "POST", Me.VirtualDirectoryName, False
    xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    xmlhttp.setRequestHeader "SOAPAction", "ExecuteUPDATE"
    
    
    SaveXmlFile "Request.xml", soapDoc
    eJournalWrite "cADORecordset.ExecuteUPDATE Send : " & format(Time, "HH:MM:SS")
    xmlhttp.SEND soapDoc.xml
    eJournalWrite "cADORecordset.ExecuteUPDATE Receive : " & format(Time, "HH:MM:SS")
    
    Set soapResponse = New MSXML2.DOMDocument30
    Set soapResponse = xmlhttp.responseXML
    SaveXmlFile "Response.xml", soapResponse
    
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
                SaveXmlFile "SoapResponseError.xml", soapResponse
                Exit Sub
            End If
        End If
    End If
    
    
End Sub




