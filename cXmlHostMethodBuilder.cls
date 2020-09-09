VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlHostMethodBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement


Private method As cXMLHostMethod

Private Function buildmethod() As Boolean
    If content Is Nothing Then
        buildmethod = False: Exit Function
    Else
        Dim source As IXMLDOMAttribute
        Set source = content.Attributes.getNamedItem("source")
        If source Is Nothing Then
            buildmethod = False: Exit Function
        Else
            Dim descriptiondoc As New MSXML2.DOMDocument60
            'Set descriptiondoc = CreateObject("Msxml2.DOMDocument.6.0")
            
            descriptiondoc.Load ReadDir & "\onlinemessages\" & source.value
            Dim generatornode As IXMLDOMElement
            Set generatornode = descriptiondoc.selectSingleNode("//generator")
            If generatornode Is Nothing Then
                buildmethod = False: Exit Function
            Else
                
                Dim generator As New MSXML2.DOMDocument60
'                Set generator = CreateObject("Msxml2.DOMDocument.6.0")
                generator.LoadXML generatornode.firstChild.XML
                
                Dim param As IXMLDOMNode
                For Each param In content.SelectNodes("./param")
                    Dim nameattr As IXMLDOMAttribute
                    Set nameattr = param.Attributes.getNamedItem("name")
                    If nameattr Is Nothing Then
                        
                    Else
                        Dim valueattr As IXMLDOMAttribute
                        Set valueattr = param.Attributes.getNamedItem("value")
                        If valueattr Is Nothing Then
                        Else
                            Dim paramdestination As IXMLDOMElement
                            Set paramdestination = generator.selectSingleNode("//*[local-name() ='param' and @name= '" & nameattr.value & "' and namespace-uri()='http://www.w3.org/1999/XSL/Transform']")
                            If paramdestination Is Nothing Then
                            Else
                                paramdestination.Text = valueattr.Text
                            End If
                        End If
                    End If
                Next param
                Dim functioncode As String
                descriptiondoc.setProperty "AllowXsltScript", True
                generator.setProperty "AllowXsltScript", True
                functioncode = descriptiondoc.transformNode(generator)
                Set method = New cXMLHostMethod
                method.code = functioncode
                
                buildmethod = True: Exit Function
            End If
            
        End If
    End If
End Function

Public Function LoadXML(Data As String) As String
    If method Is Nothing Then
        If Not buildmethod Then
            LoadXML = "": Exit Function
        End If
    End If
    
    LoadXML = method.LoadXML(Data)
    
End Function

Public Function XML() As String
    If method Is Nothing Then
        XML = ""
    Else
        XML = method.XML
    End If
End Function
