VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msgmemberwsconstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public initialpath As String
Public messagesdoc, messagedetails
Public mdlNamespace As String
Public tpNamespace As String

Private Sub Class_Initialize()
    initialpath = ReadDir & "\Messages"
    Set messagesdoc = CreateObject("Msxml2.DOMDocument.6.0")
    messagesdoc.async = False
    messagesdoc.validateOnParse = False
    messagesdoc.Load initialpath & "\" & "messages.xml"
    Set messagedetails = Nothing
    mdlNamespace = "http://www.nbg.gr/online/mdl"
    tpNamespace = "http://www.nbg.gr/online/type"
    
End Sub

Public Function build(namespace As String, element_ As MSXML2.IXMLDOMElement, ownermember As msgmember) As msgmember
    Set build = New msgmember
    With build
        .namespace = namespace
        Set .constructor = Me
        Set .element = element_
        .memberclass = element_.baseName
        Dim attr As MSXML2.IXMLDOMAttribute
        Set attr = element_.Attributes.getNamedItem("name")
        If (Not (attr Is Nothing)) Then .name = attr.value
        Set attr = element_.Attributes.getNamedItem("type")
        If (Not (attr Is Nothing)) Then .membertype = attr.value Else .membertype = .name
        Set attr = element_.Attributes.getNamedItem("min")
        If (Not (attr Is Nothing)) Then .min = attr.value
        Set attr = element_.Attributes.getNamedItem("max")
        If (Not (attr Is Nothing)) Then .max = attr.value
        Set attr = element_.Attributes.getNamedItem("unbounded")
        If (Not (attr Is Nothing)) Then .unbounded = attr.value
        If (ownermember Is Nothing) Then
            Set .rootmember = build
            .xpath = "//MESSAGEDATA"
        Else
            Set .rootmember = ownermember
            .xpath = ownermember.xpath & "/" & .name
        End If
        buildworkelement build, ownermember
        buildmembers build
    End With
End Function

Public Sub buildworkelement(member As msgmember, ownermember As msgmember)
    Dim workElement As MSXML2.IXMLDOMElement
    If ownermember Is Nothing Then
        Dim workdoc As MSXML2.DOMDocument30
        Set workdoc = New MSXML2.DOMDocument30
        Set workElement = workdoc.createElement("MESSAGEDATA")
        workdoc.appendChild workElement
        Set member.workDocument = workdoc
    Else
        Set member.workDocument = ownermember.workDocument
        Set workElement = member.workDocument.createElement(member.name)
        ownermember.workDocument.selectSingleNode(ownermember.xpath).appendChild workElement
    End If
End Sub

Public Function clonemember(inmember As msgmember) As msgmember
    Set clonemember = New msgmember
    With clonemember
        .name = inmember.name
        .namespace = inmember.namespace
        .memberclass = inmember.memberclass
        .membertype = inmember.membertype
        .min = inmember.min
        .max = inmember.max
        .unbounded = inmember.unbounded
        Set .element = inmember.element
        Set .workDocument = inmember.workDocument
        .xpath = inmember.xpath
        'Set .workelement = inmember.workelement
        Set .constructor = Me
        .initialworkstructure = inmember.initialworkstructure
        Set .rootmember = inmember.rootmember
        If Not (inmember.memberstable Is Nothing) Then
            Set .memberstable = inmember.memberstable
        End If
    End With
End Function

Public Function description(member As msgmember) As MSXML2.IXMLDOMNode
    Dim Node As MSXML2.IXMLDOMNode
    Dim apath As String
    Set Node = Nothing
    If member.memberclass = "has" Then
        apath = "//*[local-name() = 'types' and namespace-uri() = '" & tpNamespace & "' and @tns='" & member.namespace & "']//*[local-name() = 'type' and namespace-uri() = '" & tpNamespace & "' and @name='" + member.membertype + "']"
        Set Node = member.element.ownerDocument.selectSingleNode(apath)
    End If
    If member.memberclass = "part" Then
        apath = "//*[local-name() = 'types' and namespace-uri() = '" & tpNamespace & "' and @tns='" & member.namespace & "']//*[local-name() = 'type' and namespace-uri() = '" & tpNamespace & "' and @name='" + member.membertype + "']"
        Set Node = member.element.ownerDocument.selectSingleNode(apath)
    End If
    If member.memberclass = "message" Then
        apath = "//*[local-name() = 'messages' and namespace-uri() = '" & mdlNamespace & "' and @tns='" & member.namespace & "']//*[local-name() = 'message' and namespace-uri() = '" & mdlNamespace & "' and @name='" + member.membertype + "']"
        Set Node = member.element.selectSingleNode(apath)
    End If
    Set description = Node
End Function

Public Function buildmembers(member As msgmember) As Boolean
    buildmembers = False
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = description(member)
    If Not (Node Is Nothing) Then
        Dim cnode As MSXML2.IXMLDOMNode
        For Each cnode In Node.childNodes
            Dim cmember As msgmember
            If cnode.nodeType = NODE_ELEMENT Then
                If cnode.baseName <> "rules" Then
                    Set cmember = build(member.namespace, cnode, member)
                    If member.memberstable Is Nothing Then Set member.memberstable = New Collection
                    member.memberstable.add cmember, cmember.name
                End If
            End If
        Next cnode
    End If
    member.initialworkstructure = member.workDocument.selectSingleNode(member.xpath).XML
    buildmembers = True
End Function

Public Property Get XML(member As msgmember)
    Dim copydoc As New MSXML2.DOMDocument30
    copydoc.LoadXML member.workDocument.XML
    Dim attr As MSXML2.IXMLDOMAttribute
    Set attr = copydoc.createAttribute("xmlns")
    attr.value = "http://www.nbg.gr/online/msg"
    'attr.value = member.rootmember.namespace
    copydoc.documentElement.Attributes.setNamedItem attr
    
    XML = copydoc.XML
End Property

Public Property Let XML(member As msgmember, value)
    Dim copydoc As New MSXML2.DOMDocument30
    copydoc.LoadXML value
    Dim anode As MSXML2.IXMLDOMNode
    Set anode = copydoc.selectSingleNode("//ROOT")
    If anode Is Nothing Then
    Else
        If anode.firstChild Is Nothing Then
        Else
            Set anode = anode.firstChild
            If anode.firstChild Is Nothing Then
            Else
                Dim astr As String
                Dim bstr As String
                astr = member.rootmember.workDocument.documentElement.baseName
                Dim childnode As MSXML2.IXMLDOMNode
                bstr = "<" & astr & ">"
                For Each childnode In anode.childNodes
                    bstr = bstr & childnode.XML
                Next childnode
                bstr = bstr & "</" & astr & ">"
                member.workDocument.LoadXML bstr
            End If
        End If
    End If
End Property

Public Function buildmessage(namespace As String, name As String) As msgmember
    Dim apath As String
    
    apath = "//*[local-name() = 'import' and namespace-uri() = '" & mdlNamespace & "' and @namespace='" + namespace + "']"
    
    Dim Node As MSXML2.IXMLDOMNode
    Dim attr As MSXML2.IXMLDOMAttribute
    Dim afilename As String
    Set Node = messagesdoc.selectSingleNode(apath)
    If Not (Node Is Nothing) Then
        Set attr = Node.Attributes.getNamedItem("location")
        If attr Is Nothing Then
            afilename = ""
        Else
            afilename = attr.value
        End If
    Else
        afilename = ""
    End If
    
    If afilename = "" Then
        LogMsgbox "Δεν βρέθηκε το αρχείο προδιαγραφών για το μύνημα:  " & namespace & ":" & name, vbCritical, "Λάθος"
        Set buildmessage = Nothing: Exit Function
    Else
        If messagedetails Is Nothing Then
            Set messagedetails = CreateObject("Msxml2.DOMDocument.6.0")
            messagedetails.async = False
            messagedetails.validateOnParse = False
        End If
        messagedetails.Load initialpath & "\" & afilename
    End If
    
    apath = "//*[local-name() = 'messages' and namespace-uri() = '" & mdlNamespace & "' and @tns='" & namespace & "']//*[local-name() = 'message' and namespace-uri() = '" & mdlNamespace & "' and @name='" + name + "']"
    Set Node = messagedetails.selectSingleNode(apath)
    If Node Is Nothing Then
        Set buildmessage = Nothing
    Else
        Set buildmessage = build(namespace, Node, Nothing)
    End If
End Function
