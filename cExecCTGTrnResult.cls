VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExecCTGTrnResult"
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

Private Function TranslateLong(aValue As String) As Long
    TranslateLong = 0
    If aValue = "" Then
        TranslateLong = 0
    Else
        If IsNumeric(aValue) Then
            TranslateLong = CLng(aValue)
        Else
            TranslateLong = 0
        End If
    End If
End Function

Public Sub TranslateNode(Data As MSXML2.IXMLDOMNode)
    Dim MessageIDnode As MSXML2.IXMLDOMNode
    Dim ResponseDatanode As MSXML2.IXMLDOMNode
    Dim ResponseLengthnode As MSXML2.IXMLDOMNode
    Dim Exceptionnode As MSXML2.IXMLDOMNode
    
    Dim Node As MSXML2.IXMLDOMNode
    For Each Node In Data.childNodes
        If Node.baseName = "MessageID" Then Set MessageIDnode = Node
        If Node.baseName = "ResponseData" Then Set ResponseDatanode = Node
        If Node.baseName = "ResponseLength" Then Set ResponseLengthnode = Node
        If Node.baseName = "Exception" Then Set Exceptionnode = Node
    Next Node
    
    If Not MessageIDnode Is Nothing Then MessageID = MessageIDnode.Text
    If Not ResponseDatanode Is Nothing Then ResultData = ResponseDatanode.Text
    If Not ResponseLengthnode Is Nothing Then ResultLength = TranslateLong(ResponseLengthnode.Text)
    If Not Exceptionnode Is Nothing Then ExceptionData = Exceptionnode.Text
    If Not MessageIDnode Is Nothing Then MessageID = MessageIDnode.Text
End Sub


