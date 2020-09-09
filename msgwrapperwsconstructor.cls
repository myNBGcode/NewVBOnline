VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msgwrapperwsconstructor"
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

Public Function build(namespace As String, element_ As MSXML2.IXMLDOMElement, wrapper As msgwrapper) As msgwrapper
    Dim name As String, memberclass As String, membertype As String
    Dim newelement As MSXML2.IXMLDOMElement
    Dim attr As MSXML2.IXMLDOMAttribute
    
    Set attr = element_.Attributes.getNamedItem("name")
    If (Not (attr Is Nothing)) Then name = attr.value
    Set attr = element_.Attributes.getNamedItem("type")
    If (Not (attr Is Nothing)) Then membertype = attr.value Else membertype = name
    memberclass = element_.baseName
    
    If wrapper Is Nothing Then
        Set wrapper = New msgwrapper
        With wrapper
            .namespace = namespace
            .name = "wrapper"
            .xpath = "/"
            .name = name
            Set .element = element_
            Set .constructor = Me
        End With
        Set newelement = buildworkelement(namespace, name, memberclass, membertype, wrapper, Nothing)
        
    Else
        Set newelement = buildworkelement(namespace, name, memberclass, membertype, wrapper, Nothing)
        
    End If
    Set build = wrapper
End Function

Public Function buildempty(namespace As String, wrapper As msgwrapper) As msgwrapper
    Dim name As String, memberclass As String, membertype As String
    Dim newelement As MSXML2.IXMLDOMElement
    Dim attr As MSXML2.IXMLDOMAttribute
    
    If wrapper Is Nothing Then
        Set wrapper = New msgwrapper
        With wrapper
            .namespace = namespace
            .name = "wrapper"
            .xpath = "/"
            Set .constructor = Me
            Set .element = Nothing
            Set .workDocument = New MSXML2.DOMDocument30
            Set .workElement = Nothing
        End With
    Else
        
    End If
    Set buildempty = wrapper
End Function

Public Function buildworkelement(namespace As String, name As String, memberclass As String, membertype As String, wrapper As msgwrapper, parentelement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement
    If wrapper.workDocument Is Nothing Then
        Set wrapper.workDocument = New MSXML2.DOMDocument30
        Set buildworkelement = wrapper.workDocument.createElement("MESSAGEDATA")
        Set wrapper.workElement = wrapper.workDocument.documentElement
        wrapper.workDocument.appendChild buildworkelement
        
    Else
        Set wrapper.workElement = wrapper.workDocument.createElement(name)
        Set buildworkelement = wrapper.workElement
        parentelement.appendChild buildworkelement
    End If
    buildmembers namespace, memberclass, membertype, wrapper, buildworkelement
End Function

Public Function description(namespace As String, memberclass As String, membertype As String, wrapper As msgwrapper) As MSXML2.IXMLDOMNode
    Dim Node As MSXML2.IXMLDOMNode
    Dim apath As String
    Set Node = Nothing
    If memberclass = "has" Then
        apath = "//*[local-name() = 'types' and namespace-uri() = '" & tpNamespace & "' and @tns='" & namespace & "']//*[local-name() = 'type' and namespace-uri() = '" & tpNamespace & "' and @name='" + membertype + "']"
        Set Node = wrapper.element.ownerDocument.selectSingleNode(apath)
    End If
    If memberclass = "part" Then
        apath = "//*[local-name() = 'types' and namespace-uri() = '" & tpNamespace & "' and @tns='" & namespace & "']//*[local-name() = 'type' and namespace-uri() = '" & tpNamespace & "' and @name='" + membertype + "']"
        Set Node = wrapper.element.ownerDocument.selectSingleNode(apath)
    End If
    If memberclass = "message" Then
        apath = "//*[local-name() = 'messages' and namespace-uri() = '" & mdlNamespace & "' and @tns='" & namespace & "']//*[local-name() = 'message' and namespace-uri() = '" & mdlNamespace & "' and @name='" + membertype + "']"
        Set Node = wrapper.element.selectSingleNode(apath)
    End If
    Set description = Node
End Function

Public Function buildmembers(namespace As String, memberclass As String, membertype As String, wrapper As msgwrapper, parentelement As MSXML2.IXMLDOMElement) As Boolean
    buildmembers = False
    Dim Node As MSXML2.IXMLDOMNode
    Set Node = description(namespace, memberclass, membertype, wrapper)
    If Not (Node Is Nothing) Then
        Dim cnode As MSXML2.IXMLDOMNode
        For Each cnode In Node.childNodes
            Dim cmember As msgmember
            If cnode.nodeType = NODE_ELEMENT Then
                If cnode.baseName <> "rules" Then
                    Dim attr As MSXML2.IXMLDOMAttribute, newelement As MSXML2.IXMLDOMElement, newname As String, newmemberclass As String, newmembertype As String

                    newmemberclass = cnode.baseName
                    Set attr = cnode.Attributes.getNamedItem("name")
                    If (Not (attr Is Nothing)) Then newname = attr.value
                    Set attr = cnode.Attributes.getNamedItem("type")
                    If (Not (attr Is Nothing)) Then newmembertype = attr.value

                    Set newelement = buildworkelement(namespace, newname, newmemberclass, newmembertype, wrapper, parentelement)
                End If
            End If
        Next cnode
    End If
    'member.initialworkstructure = member.workdocument.selectSingleNode(member.xpath).XML
    buildmembers = True
End Function

Public Function clonewrapper(inwrapper As msgwrapper) As msgwrapper
    Set clonewrapper = New msgwrapper
    With clonewrapper
        .name = inwrapper.name
        .namespace = inwrapper.namespace
        Set .element = inwrapper.element
        Set .workDocument = inwrapper.workDocument
        Set .workElement = inwrapper.workElement
        .xpath = inwrapper.xpath
        Set .constructor = Me
    End With
End Function


Public Function buildwrapper(namespace As String, Optional name As String) As msgwrapper
    Dim apath As String
    
    If Not IsMissing(name) And name <> "" Then
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
            Set buildwrapper = Nothing: Exit Function
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
            Set buildwrapper = Nothing
        Else
            Set buildwrapper = build(namespace, Node, Nothing)
        End If
    Else
        Set buildwrapper = buildempty(namespace, Nothing)
    End If
End Function

'Public Property Get XML(member As msgmember)
'    Dim copydoc As New MSXML2.DOMDocument30
'    copydoc.LoadXML member.workdocument.XML
'    Dim attr As MSXML2.IXMLDOMAttribute
'    Set attr = copydoc.createAttribute("xmlns")
'    attr.value = "http://www.nbg.gr/online/msg"
'    'attr.value = member.rootmember.namespace
'    copydoc.documentElement.Attributes.setNamedItem attr
'
'    XML = copydoc.XML
'End Property
'
Public Property Let XML(wrapper As msgwrapper, value)
    
    
    Dim copydoc
    Dim astr As String
    Dim bstr As String
    Set copydoc = CreateObject("Msxml2.DOMDocument.6.0")

    copydoc.LoadXML value
    Dim anode As MSXML2.IXMLDOMNode
    Set anode = copydoc.selectSingleNode("//ROOT")
    astr = wrapper.workDocument.documentElement.baseName
    If anode Is Nothing Then
    Else
        If anode.firstChild Is Nothing Then
            bstr = "<" & astr & ">"
            bstr = bstr & "</" & astr & ">"
            wrapper.workDocument.LoadXML bstr
        Else
            Set anode = anode.firstChild
            If anode.firstChild Is Nothing Then
                bstr = "<" & astr & ">"
                bstr = bstr & "</" & astr & ">"
                wrapper.workDocument.LoadXML bstr
            Else
'                astr = wrapper.workDocument.documentElement.baseName
                Dim childnode As MSXML2.IXMLDOMNode
                bstr = "<" & astr & ">"
                For Each childnode In anode.childNodes
                    bstr = bstr & childnode.XML
                Next childnode
                bstr = bstr & "</" & astr & ">"
                wrapper.workDocument.LoadXML bstr
            End If
        End If
        Set wrapper.workElement = wrapper.workDocument.documentElement
    End If
End Property
