VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cL2Totals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<totals termid="">
'<entry name="dbtotals1" currency="070" dbvalue="1000" crvalue="222" />
'<entry name="dbtotals1" currency="070" dbvalue="1000" crvalue="222" />
'</totals>
Option Explicit
Public content As IXMLDOMElement
Public termid As String
Public Manager As cXMLDocumentManager
Public StaticLink As IXMLDOMElement

Public Function ParseTotals(inElement As IXMLDOMElement) As String
    On Error GoTo GenError
    Dim entryList As IXMLDOMNodeList
    Dim nameattr As IXMLDOMAttribute, curattr As IXMLDOMAttribute, dbattr As IXMLDOMAttribute, crattr As IXMLDOMAttribute
    Dim i As Integer
    Set entryList = inElement.SelectNodes("//entry")
    For i = 0 To entryList.length - 1
       Set nameattr = entryList(i).Attributes.getNamedItem("name")
       If Not (nameattr Is Nothing) Then
          If nameattr.Text <> "" Then
          Set curattr = entryList(i).Attributes.getNamedItem("currency")
          Set dbattr = entryList(i).Attributes.getNamedItem("dbvalue")
          Set crattr = entryList(i).Attributes.getNamedItem("crvalue")
          If curattr Is Nothing Then
             If Not (dbattr Is Nothing) Then
                TRNFrm.AddDBTotal nameattr.Text, dbattr.Text
             End If
             If Not (crattr Is Nothing) Then
                TRNFrm.AddCRTotal nameattr.Text, crattr.Text
             End If
          Else
             If Not (dbattr Is Nothing) Then
                TRNFrm.AddCurDBTotal nameattr.Text, curattr.Text, dbattr.Text
             End If
              If Not (crattr Is Nothing) Then
                TRNFrm.AddCurCRTotal nameattr.Text, curattr.Text, crattr.Text
             End If
          End If
          End If
       Else
       
       End If
    Next
    ParseTotals = GetAllTermTotals
    Exit Function
GenError:
    ParseTotals = "<totals><errors>" & Err.description & "</errors></totals>"
End Function
Public Function GetAllTermTotals() As String

End Function

Public Function LoadXML(inDocument As String) As String
Dim workDocument As New MSXML2.DOMDocument30
Dim termattr As IXMLDOMAttribute
   workDocument.LoadXML inDocument
   If Not (workDocument.documentElement Is Nothing) Then
        Set termattr = workDocument.documentElement.Attributes.getNamedItem("termid")
        If Not (termattr Is Nothing) Then
            termid = termattr.Text
        End If
   End If
   LoadXML = ParseTotals(workDocument.documentElement)
End Function

Public Function XML() As String
    If StaticLink Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(StaticLink.XML)
    End If
End Function
