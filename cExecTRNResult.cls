VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExecTrnResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ResultLength As Long
Public ResultData As String
Public MessageID As String
Public ExceptionData As String
Public ExecutionTime As Long

Private Function TranslateLong(avalue As String) As Long
    TranslateLong = 0
    If avalue = "" Then
        TranslateLong = 0
    Else
        If IsNumeric(avalue) Then
            TranslateLong = CLng(avalue)
        Else
            TranslateLong = 0
        End If
    End If
End Function

Public Sub TranslateNode(data As MSXML2.IXMLDOMNode)
    Dim MessageIDnode As MSXML2.IXMLDOMNode
    Dim ResponseDatanode As MSXML2.IXMLDOMNode
    Dim ResponseLengthnode As MSXML2.IXMLDOMNode
    Dim Exceptionnode As MSXML2.IXMLDOMNode
    
    Dim node As MSXML2.IXMLDOMNode
    For Each node In data.childNodes
        If node.baseName = "MessageID" Then Set MessageIDnode = node
        If node.baseName = "ResponseData" Then Set ResponseDatanode = node
        If node.baseName = "ResponseLength" Then Set ResponseLengthnode = node
        If node.baseName = "Exception" Then Set Exceptionnode = node
    Next node
    
    'Set MessageIDnode = data.selectSingleNode(".//MessageID")
    'Set ResponseDatanode = data.selectSingleNode(".//ResponseData")
    'Set ResponseLengthnode = data.selectSingleNode(".//ResponseLength")
    'Set Exceptionnode = data.selectSingleNode(".//Exception")

    If Not MessageIDnode Is Nothing Then MessageID = MessageIDnode.Text
    If Not ResponseDatanode Is Nothing Then ResultData = ResponseDatanode.Text
    If Not ResponseLengthnode Is Nothing Then ResultLength = TranslateLong(ResponseLengthnode.Text)
    If Not Exceptionnode Is Nothing Then ExceptionData = Exceptionnode.Text
    If Not MessageIDnode Is Nothing Then MessageID = MessageIDnode.Text
    
'      <ExecTRNB64CompressedWithLogV2Result>
'        <MessageID>string</MessageID>
'        <ResponseData>string</ResponseData>
'        <ResponseLength>int</ResponseLength>
'        <Exception>string</Exception>
'      </ExecTRNB64CompressedWithLogV2Result>

End Sub

