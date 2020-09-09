VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScriptHostMethodBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private method As cXMLHostMethod
Private fakeResult As String

Public initial As cHostMessage
Public last As cHostMessage


Public Property Let hostresult(aValue As String)
    fakeResult = aValue
    method.fakeResult = aValue
End Property

Public Function buildmethod(source As String, paramnames(), paramvalues()) As Boolean
    Dim descriptiondoc As New MSXML2.DOMDocument60
    'Set descriptiondoc = CreateObject("Msxml2.DOMDocument.6.0")
    
    descriptiondoc.Load ReadDir & "\onlinemessages\" & source
    Dim generatornode As IXMLDOMElement
    Set generatornode = descriptiondoc.selectSingleNode("//generator")
    If generatornode Is Nothing Then
        buildmethod = False: Exit Function
    Else
        
        Dim generator As New MSXML2.DOMDocument60
'        Set generator = CreateObject("Msxml2.DOMDocument.6.0")
        generator.LoadXML generatornode.firstChild.XML
        
        Dim param As IXMLDOMNode
        Dim i As Long
        For i = LBound(paramnames) To UBound(paramnames)
            Dim paramdestination As IXMLDOMElement
            Set paramdestination = generator.selectSingleNode("//*[local-name() ='param' and @name= '" & paramnames(i) & "' and namespace-uri()='http://www.w3.org/1999/XSL/Transform']")
            If paramdestination Is Nothing Then
            Else
                paramdestination.Text = paramvalues(i)
            End If
        Next i
        
        Dim functioncode As String
        descriptiondoc.setProperty "AllowXsltScript", True
        generator.setProperty "AllowXsltScript", True
        functioncode = descriptiondoc.transformNode(generator)
        Set method = New cXMLHostMethod
        method.code = functioncode
        
        Set initial = New cHostMessage
        initial.prepare "initialmessage"
        Set last = New cHostMessage
        initial.prepare "lastmessage"
        
        buildmethod = True: Exit Function
    End If
End Function

Public Function Run() As Boolean
    Dim Result As String
    Result = method.LoadXML(initial.workDocument.XML)
    If Result = "" Then
        Run = False: Exit Function
    Else
        Set last = New cHostMessage
        last.prepare "lastmessage"
        last.workDocument.LoadXML Result
        Run = True
    End If
End Function

