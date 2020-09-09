VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msgmember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public namespace As String
Public name As String, memberclass As String, membertype As String
Public min As Integer, max As Integer, unbounded As Boolean
Public element As MSXML2.IXMLDOMElement
Public workDocument As MSXML2.DOMDocument30
Public xpath As String
'Public workelement As MSXML2.IXMLDOMElement
Public memberstable As Collection
Public initialworkstructure As String
Public rootmember As msgmember
Public constructor
Public StrictPath As Boolean

Private Sub Class_Initialize()
    name = ""
    memberclass = ""
    min = 1
    max = 1
    Set element = Nothing
    Set workDocument = Nothing
    xpath = ""
    'Set workelement = Nothing
    initialworkstructure = ""
    Set rootmember = Nothing
End Sub

Public Sub buildworkelement(ownermember As msgmember)
    Dim workElement As MSXML2.IXMLDOMElement
    If ownermember Is Nothing Then
        Dim workdoc As MSXML2.DOMDocument30
        Set workdoc = New MSXML2.DOMDocument30
        Set workElement = workdoc.createElement(name)
        workdoc.appendChild workElement
        Set workDocument = workdoc
    Else
        Set workDocument = ownermember.workDocument
        Set workElement = workDocument.createElement(name)
        ownermember.workDocument.selectSingleNode(ownermember.xpath).appendChild workElement
    End If
End Sub

Public Sub SaveWorkDocument(path As String)
    If workDocument Is Nothing Then
    Else
        workDocument.save path
    End If
    
End Sub

Public Property Get xml()
    xml = constructor.xml(Me)
End Property

Public Property Let xml(value)
    constructor.xml(Me) = value
End Property

Public Property Get workElement() As MSXML2.IXMLDOMElement
    Set workElement = workDocument.selectSingleNode(xpath)
End Property

Public Function bottomupscan(Node As MSXML2.IXMLDOMNode) As msgmember
    If Node.parentNode.baseName <> "" Then
        Dim parent As msgmember
        Set parent = bottomupscan(Node.parentNode)
        Set bottomupscan = parent.memberstable(Node.baseName)
    Else
        Set bottomupscan = rootmember
    End If
End Function

Public Property Get find(axpath As String) As msgmember
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = workDocument.selectSingleNode(axpath)
    If Node Is Nothing Then
        Set find = Nothing
    Else
        Set find = bottomupscan(Node)
        If find Is Nothing Then
        Else
            Set find = constructor.clonemember(find)
            If Left(axpath, 1) = "." Then
                If Len(axpath) = 1 Then
                    find.xpath = xpath
                Else
                    find.xpath = xpath & Right(axpath, Len(axpath) - 1)
                End If
            Else
                find.xpath = axpath
            End If
        End If
    End If
End Property

Private Function GetValueNode(axpath As String) As MSXML2.IXMLDOMNode
    Dim Node As MSXML2.IXMLDOMNode
    If workDocument.selectSingleNode(xpath) Is Nothing Then Exit Function
    Set Node = workDocument.selectSingleNode(xpath).selectSingleNode(axpath)
    If Node Is Nothing Then
        If StrictPath Then
        LogMsgbox "Λάθος στον εντοπισμό τμήματος msgmember: " & rootmember.name & ".." & name & " path: " & axpath, vbCritical, "Λάθος"
        End If
        Set GetValueNode = Nothing
    Else
        'If Node.childNodes.Length > 0 Then
        '    MsgBox "Λάθος τύπος τμήματος msgmember: " & rootmember.name & ".." & name & " path: " & axpath, vbCritical, "Λάθος"
        '    Set GetValueNode = Nothing
        'Else
            Set GetValueNode = Node
        'End If
    End If
    
End Function

Public Property Get value(axpath As String)
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = GetValueNode(axpath)
    If Node Is Nothing Then
        value = ""
    Else
        value = Node.Text
    End If
End Property

Public Property Let value(axpath As String, invalue)
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = GetValueNode(axpath)
    If Node Is Nothing Then
    Else
        Node.Text = invalue
    End If
End Property

Public Property Get intvalue(axpath As String) As Long
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = GetValueNode(axpath)
    If Node Is Nothing Then
        intvalue = 0
    Else
        intvalue = CLng("0" & Trim(Node.Text))
    End If
End Property

Public Property Let intvalue(axpath As String, invalue As Long)
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = GetValueNode(axpath)
    If Node Is Nothing Then
    Else
        Node.Text = CStr(invalue)
    End If
End Property


