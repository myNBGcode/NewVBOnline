VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Manager As cXMLDocumentManager


'Public Asserts As New Collection
'Private message As String
Public content As IXMLDOMElement
Private operation As String
Private Rules As Collection

Public Function LoadXml(value As String) As String

    Dim validationdoc As New MSXML2.DOMDocument30
    validationdoc.LoadXml value
    If validationdoc.documentElement Is Nothing Then Exit Function
    'validationDoc.save ("c:\xmls\validation.xml")
    Dim doc As IXMLDOMElement, attr As IXMLDOMAttribute
    
    Set doc = validationdoc.documentElement
   Set Rules = New Collection
   LoadXml = Validate(doc)
    'content exei to validation
    'validationdoc exei to input
    
End Function

Function Validate(xmlInput As IXMLDOMElement) As String
  Dim attr As IXMLDOMAttribute
  Dim xmlrule As IXMLDOMElement
  Dim xmlassert As IXMLDOMElement
  Dim xmlreport As IXMLDOMElement
  
  operation = "and"
  
  For Each attr In content.Attributes
    Select Case UCase(attr.baseName)
        Case "OPERATION"
            operation = attr.nodeValue
    End Select
  Next attr
  
   Dim cRule As cXMLValidationRule
    Dim cAssert As cXMLValidationAssert
    Dim cReport As cXMLValidationAssert
  For Each xmlrule In content.SelectNodes("rule")
    Set cRule = New cXMLValidationRule
    cRule.Context = xmlrule.Attributes.getNamedItem("context").nodeValue
    If cRule.Context = "" Then
        cRule.Context = "//"
    End If
    For Each xmlassert In xmlrule.SelectNodes("assert")
        Set cAssert = New cXMLValidationAssert
        cAssert.Message = xmlassert.Attributes.getNamedItem("message").nodeValue
        cAssert.test = xmlassert.Attributes.getNamedItem("test").nodeValue
        cAssert.vType = "ASSERT"
        cAssert.PassTest = True
        cRule.Asserts.add cAssert
    Next xmlassert
    For Each xmlreport In xmlrule.SelectNodes("report")
        Set cReport = New cXMLValidationAssert
        cReport.Message = xmlreport.Attributes.getNamedItem("message").nodeValue
        cReport.test = xmlreport.Attributes.getNamedItem("test").nodeValue
        cReport.vType = "REPORT"
        cReport.PassTest = True
        cRule.Asserts.add cReport
    Next xmlreport
    
    Rules.add cRule
  Next xmlrule
  
  Dim rule As cXMLValidationRule
  Dim AssertFailed As cXMLValidationAssert
  Dim pass As Boolean
 pass = False
  For Each rule In Rules
    
    Set AssertFailed = ValidateRule(rule, xmlInput)
    If Rules.count > 1 Then
        If UCase(operation) = "AND" Then
            If Not AssertFailed Is Nothing Then
                pass = True
                
            ElseIf AssertFailed Is Nothing Then
                Validate = "<statusmessage>" + "OK" + "</statusmessage>"
                Exit Function
            End If
        ElseIf UCase(operation) = "OR" Then
            If Not AssertFailed Is Nothing Then
                 Validate = "<statusmessage>" + AssertFailed.Message + "</statusmessage>"
                Exit Function
            End If
        End If
    Else
        If Not AssertFailed Is Nothing Then
                Validate = "<statusmessage>" + AssertFailed.Message + "</statusmessage>"
                Exit Function
        End If
    End If
    
  Next rule
  If pass = True Then
     Validate = "<statusmessage>" + AssertFailed.Message + "</statusmessage>"
     Exit Function
  End If
 
  Validate = ""
End Function

Function ValidateRule(rule As cXMLValidationRule, xmlInput As IXMLDOMElement) As cXMLValidationAssert
    Dim succeed As Boolean
    Dim AssertRes As Integer
    AssertRes = 0
    Dim inputContext As IXMLDOMElement
    Set inputContext = xmlInput.SelectNodes(rule.Context).item(0)
    
    Dim list As IXMLDOMNodeList
    succeed = True

    Dim Assert As cXMLValidationAssert
    For Each Assert In rule.Asserts
        AssertRes = AssertTest(inputContext, Assert.test)
        If AssertRes = 0 Then
            Assert.PassTest = False
            succeed = False
            Set ValidateRule = Assert
            Exit For
        ElseIf AssertRes = 1 Then
            Assert.PassTest = True
        End If
        'Set list = inputContext.SelectNodes(Assert.test)
        'If list.length = 0 Then
        '    Assert.PassTest = False
        '    succeed = False
        'ElseIf list.length > 0 Then
        '    Assert.PassTest = True
        'End If
    Next Assert
    
'    If succeed = True Then ' ολα ισχυουν αρα validationError
'        For Each Assert In Rule.Asserts
'            If Assert.PassTest = True Then
'                Set ValidateRule = Assert
'                Exit For
'            End If
'        Next Assert
'
'    End If
    
End Function

Function AssertTest(Context As IXMLDOMElement, xpath As String) As Integer
    Dim xsltdoc As New MSXML2.DOMDocument30
    Dim astr As String
    xpath = Replace(xpath, "<", "&lt;")
    xpath = Replace(xpath, ">", "&gt;")
    astr = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" xmlns:fo=""http://www.w3.org/1999/XSL/Format"">  " & _
        " <xsl:template match=""/"">" & _
        " <xsl:choose>" & _
        " <xsl:when test="" " + xpath + """>" & _
        "<XPATH>1</XPATH>" & _
        " </xsl:when>" & _
        "<xsl:otherwise>" + "<XPATH>0</XPATH>" + "</xsl:otherwise>" + "</xsl:choose>" & _
        " </xsl:template> </xsl:stylesheet>"
        
    xsltdoc.LoadXml astr
    'xsltdoc.save "c:\xslt.xml"
  Dim Contextdoc As New MSXML2.DOMDocument30
  Contextdoc.LoadXml Context.XML
  'Contextdoc.save "c:\source.xml"
 Dim LoadXml As String
 LoadXml = Contextdoc.transformNode(xsltdoc)
 
 Dim res As New MSXML2.DOMDocument30
 res.LoadXml LoadXml
 
 AssertTest = CInt(res.selectSingleNode("/XPATH").Text)

   
End Function
