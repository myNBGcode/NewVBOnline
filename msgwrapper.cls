VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msgwrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public namespace As String
Public name As String
Public element As MSXML2.IXMLDOMElement
Public workDocument As MSXML2.DOMDocument30
Public workElement As MSXML2.IXMLDOMElement
Public xpath As String
Public constructor

Private Sub Class_Initialize()
    name = "": namespace = ""
    Set element = Nothing
    Set workDocument = Nothing
    Set constructor = Nothing
    xpath = ""
End Sub

'Public Sub SaveWorkDocument(path As String)
'    If workDocument Is Nothing Then
'    Else
'        workDocument.Save path
'    End If
'
'End Sub
'
'Public Property Get XML()
'    XML = constructor.XML(Me)
'End Property
'


Public Property Let xml(value)
    constructor.xml(Me) = value
End Property
'
'Public Property Get workelement() As MSXML2.IXMLDOMElement
'    Set workelement = workDocument.selectSingleNode(xpath)
'End Property
'
'Public Function bottomupscan(Node As MSXML2.IXMLDOMNode) As msgmember
'    If Node.parentNode.baseName <> "" Then
'        Dim parent As msgmember
'        Set parent = bottomupscan(Node.parentNode)
'        Set bottomupscan = parent.memberstable(Node.baseName)
'    Else
'        Set bottomupscan = rootmember
'    End If
'End Function
'
Public Property Get find(axpath As String) As msgwrapper
    Dim Node As MSXML2.IXMLDOMNode
    If workElement Is Nothing Then
        If xpath = "/" Then
            Set workElement = workDocument.documentElement
        Else
            Set workElement = workDocument.selectSingleNode(xpath)
        End If
    End If
    Set Node = workElement.selectSingleNode(axpath)
    If Node Is Nothing Then
        Set find = Nothing
    Else
        Set find = constructor.clonewrapper(Me)
        Set find.workElement = Node
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
End Property

Public Function ExtendNode(xml, Optional includeroot) As Boolean
    Dim newdoc As MSXML2.DOMDocument30
    
    ExtendNode = False
    If IsMissing(includeroot) Then includeroot = False
    
    On Error GoTo invalidworkelement
    If workElement Is Nothing Then
        If xpath = "/" Then
            Set workElement = workDocument.documentElement
        Else
            Set workElement = workDocument.selectSingleNode(xpath)
        End If
    End If
    If workElement Is Nothing Then
        If xpath = "/" Then
            If includeroot Then
                workDocument.LoadXml xml
                Set workElement = workDocument.documentElement
            Else
                GoTo invalidworkelement
            End If
        Else
            GoTo invalidworkelement
        End If
    Else
    On Error GoTo invaliddocumentelement
    Set newdoc = New MSXML2.DOMDocument30
    newdoc.LoadXml xml
    If newdoc.parseError.errorCode <> 0 Then GoTo invalidxmlstring
    
    Dim elm As IXMLDOMElement
    If Not includeroot Then
        For Each elm In newdoc.documentElement.childNodes
            newdoc.documentElement.removeChild elm
            workElement.appendChild elm
        Next elm
    Else
        Set elm = newdoc.documentElement
        newdoc.removeChild elm
        workElement.appendChild elm
    End If
    End If
    ExtendNode = True
    Exit Function
    
invalidworkelement:
    LogMsgbox "Ανύπαρκτο workelement για το node: " & xpath & " Err: " & Err.Number & "," & Err.description, vbCritical, "Λάθος"
    Exit Function
invalidxmlstring:
    LogMsgbox "Λάθος κατα τη δημιουργία προσωρινού εγγράφου: " & xpath, vbCritical, "Λάθος"
    Xml_ParseError newdoc.parseError
    Exit Function
invaliddocumentelement:
    LogMsgbox "Λάθος κατα την επέκταση κειμένου στο node: " & xpath & " Err: " & Err.Number & "," & Err.description, vbCritical, "Λάθος"
    Exit Function
End Function

Public Function AddInstance(instancecount As Long) As Boolean
    AddInstance = False
    If workElement Is Nothing Then
    Else
        If workElement.parentNode Is Nothing Then
        
        Else
            Dim i As Long
            For i = 1 To instancecount
                workElement.parentNode.appendChild workElement.cloneNode(True)
            Next i
            AddInstance = True
        End If
    End If
End Function

Public Function Remove() As Boolean
    Remove = False
    If workElement Is Nothing Then
    Else
        If workElement.parentNode Is Nothing Then
        Else
            workElement.parentNode.removeChild workElement
            Remove = True
        End If
    End If
End Function
Private Function GetValueNode(axpath As String) As MSXML2.IXMLDOMNode
    Dim Node As MSXML2.IXMLDOMNode
    If workElement.selectSingleNode(xpath) Is Nothing Then Exit Function
    Set Node = workDocument.selectSingleNode(xpath).selectSingleNode(axpath)
    If Node Is Nothing Then
        LogMsgbox "Λάθος στον εντοπισμό τμήματος msgmember: " & name & " path: " & axpath, vbCritical, "Λάθος"
        Set GetValueNode = Nothing
    Else
        Set GetValueNode = Node
    End If
End Function

Public Property Get childlength(axpath As String) As Long
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = GetValueNode(axpath)
    If Node Is Nothing Then
        childlength = 0
    Else
        childlength = Node.childNodes.length
    End If
End Property

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



