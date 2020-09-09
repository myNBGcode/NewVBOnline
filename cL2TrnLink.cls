VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cL2TrnLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement
'<linkl2trn name="xxxx">
'   <trn name="xxxx">
'   <formupdate name="xxxx">
'       <....>
'   </formupdate>
'</link2trn>

Private Function LinkToL2TRN(Node As IXMLDOMElement) As String
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    Dim TrnCode As String, inDoc As MSXML2.DOMDocument30, outDoc As MSXML2.DOMDocument30
    Dim trnHandler As L2TrnHandler
    
    Set elm = content.selectSingleNode("trn")
    If Not elm Is Nothing Then
        Set attr = elm.getAttributeNode("name")
        If Not (attr Is Nothing) Then TrnCode = attr.Text
    End If
    Set elm = content.selectSingleNode("formupdate")
    If Not elm Is Nothing Then
        Set inDoc = New MSXML2.DOMDocument30
        inDoc.LoadXml elm.XML
    End If


    If Not Node Is Nothing Then
        For Each elm In Node.childNodes
            If UCase(elm.baseName) = UCase("trn") Then
                Set attr = elm.getAttributeNode("name")
                If Not (attr Is Nothing) Then TrnCode = attr.Text
            ElseIf UCase(elm.baseName) = UCase("formupdate") Then
                Set inDoc = New MSXML2.DOMDocument30
                inDoc.LoadXml elm.XML
            End If
        Next elm
    End If
    
    If TrnCode <> "" Then Set trnHandler = New L2TrnHandler
    If Not (trnHandler Is Nothing) Then
        If Not inDoc Is Nothing Then
            trnHandler.addFormUpdate inDoc, inDoc.documentElement.getAttribute("name")
        End If
        
        trnHandler.ExecuteForm TrnCode
        
        If trnHandler.Result Is Nothing Then
            LinkToL2TRN = ""
        Else
            LinkToL2TRN = trnHandler.Result.XML
        End If
        trnHandler.CleanUp
    End If
End Function


Public Function LoadXml(inDocument As String) As String

    Dim workDocument As New MSXML2.DOMDocument30
    workDocument.LoadXml inDocument
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    
    Set elm = workDocument.documentElement
    If Not (elm Is Nothing) Then
        If UCase(elm.baseName) = "LINKL2TRN" Then
            'Set content = elm
            LoadXml = LinkToL2TRN(elm)
            Exit Function
        End If
    End If
    If Not content Is Nothing Then
        If UCase(content.baseName) = "LINKL2TRN" Then
            LoadXml = LinkToL2TRN(elm)
        End If
    End If
    
    
    
End Function

Public Function XML() As String
    If content Is Nothing Then
        XML = ""
    Else
        XML = LoadXml(content.XML)
    End If
End Function

