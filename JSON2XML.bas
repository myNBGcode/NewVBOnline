Attribute VB_Name = "JSON2XML"
Public Function ToXML(json As String, rootxml As String) As MSXML2.DOMDocument60
    
    Dim xmldoc As New MSXML2.DOMDocument60
    Set xmldoc.documentElement = xmldoc.createElement(rootxml)
    
    Dim JB As New JsonBag
    JB.DecimalMode = True
    JB.json = json

    Call ToXMLElement(JB, xmldoc.documentElement)

    Set ToXML = xmldoc
End Function

Private Function ToXMLElement(ByRef JB As JsonBag, ByRef parent As IXMLDOMElement)
    Dim elm As IXMLDOMElement
    Dim i As Long, j As Long, k As Long
    Dim name As String
    
    For i = 1 To JB.count
        name = JB.name(i)
        If JB.ItemIsJSON(i) = False Then 'simple value
            Set elm = parent.ownerDocument.createElement(name)
            If Not IsNull(JB.Item(i)) Then elm.Text = JB.Item(i)
            parent.appendChild elm
        ElseIf JB.Item(i).isArray = True Then 'array
            For k = 1 To JB.Item(i).count
                Set elm = parent.ownerDocument.createElement(JB.name(i))
                parent.appendChild elm
                If JB.Item(i).ItemIsJSON(k) = False Then
                    If Not IsNull(JB.Item(i).Item(k)) Then elm.Text = JB.Item(i).Item(k)
                Else
                    Call ToXMLElement(JB.Item(i).Item(k), elm)
                End If
            Next
        ElseIf JB.Item(i).isArray = False Then 'object
            Set elm = parent.ownerDocument.createElement(name)
            parent.appendChild elm
            Call ToXMLElement(JB.Item(i), elm)
        End If
    Next

End Function

Public Function FromXML(XML As MSXML2.DOMDocument60, Optional ignorefirst As Boolean) As String
    FromXML = FromXML2(XML.documentElement, ignorefirst)
End Function
Public Function FromXML2(XML As MSXML2.IXMLDOMElement, Optional ignorefirst As Boolean) As String

    Dim JB As New JsonBag
    JB.DecimalMode = True
    
    If (Not IsMissing(ignorefirst)) And ignorefirst = True Then
        FromXML2 = FromXMLElement(JB, XML, XML.nodename)
    Else
        FromXML2 = FromXMLElement(JB, XML)
    End If
End Function

Private Function FromXMLElement(ByRef JB As JsonBag, elm As MSXML2.IXMLDOMElement, Optional item2serialize As String) As String
    
    Dim typeAttr As MSXML2.IXMLDOMAttribute, arrayattr As MSXML2.IXMLDOMAttribute
    Dim typeStr As String
    Dim isArray As Boolean
    Dim nodelist As IXMLDOMNodeList
    Dim anode As IXMLDOMElement, bNode As IXMLDOMElement
    Dim newJB As New JsonBag
    
    Set typeAttr = elm.getAttributeNode("type")
    Set arrayattr = elm.getAttributeNode("isArray")
    Set nodelist = elm.SelectNodes("*")
    
    If Not typeAttr Is Nothing Then
        typeStr = typeAttr.Text
    End If
    If Not arrayattr Is Nothing Then
        If UCase(arrayattr.Text) = "TRUE" Then
            isArray = True
        End If
    End If
    
    If nodelist.length > 0 Then
        If elm.parentNode.SelectNodes(elm.nodename).length = 1 And isArray = False Then
            With JB
                Set newJB = .AddNewObject(elm.nodename)
                For Each anode In nodelist
                    Call FromXMLElement(newJB, anode)
                Next
            End With
        Else
            With JB
                With .AddNewArray(elm.nodename)
                    For Each anode In elm.parentNode.SelectNodes(elm.nodename)
                        Set newJB = .AddNewObject()
                        For Each bNode In anode.childNodes
                            Call FromXMLElement(newJB, bNode)
                        Next
                    Next
                End With
            End With
        End If
    Else
        With JB
            If typeStr = "decimal" Then
                .Item(elm.nodename) = CDec(elm.Text)
            ElseIf typeStr = "guid" And Trim(elm.Text) = "" Then
                .Item(elm.nodename) = GetGuid
            Else
                .Item(elm.nodename) = elm.Text
            End If
        End With
    End If
    
    Set newJB = Nothing
    
    If (Not IsMissing(item2serialize)) And item2serialize <> "" Then
        FromXMLElement = JB.Item(item2serialize).json
    Else
        FromXMLElement = JB.json
    End If
End Function
