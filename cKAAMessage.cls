VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKAAMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private name As String
Private MessageDocument As MSXML2.DOMDocument30

Public Sub prepare(aname As String)
    name = aname
    
    Dim adoc As New MSXML2.DOMDocument30
    adoc.Load ReadDir & "/SAT APIS Προδιαγραφες/SATMessages.xml"
    Dim messageelement As IXMLDOMElement
    Set messageelement = adoc.selectSingleNode("//message[@name='" & aname & "']")
    If messageelement Is Nothing Then
        Err.Raise 1000, "cKAAMessage", "Δεν βρέθηκε το μήνυμα: " & aname
    End If
    
    Set MessageDocument = New MSXML2.DOMDocument30
    MessageDocument.LoadXML messageelement.XML
    
End Sub

Public Property Get Document() As MSXML2.DOMDocument30
    Set Document = MessageDocument
End Property

Public Property Get value(path As String)
    Dim valueelm As IXMLDOMElement
    Set valueelm = MessageDocument.selectSingleNode(path)
    If valueelm Is Nothing Then
        Err.Raise 1001, "cKAAMessage.Value get", "Δεν βρέθηκε η διαδρομή " & path
    End If
    
    Dim valueattr As IXMLDOMAttribute
    Set valueattr = valueelm.Attributes.getNamedItem("value")
    If valueattr Is Nothing Then
        Set valueattr = valueelm.ownerDocument.createAttribute("value")
        valueelm.Attributes.setNamedItem valueattr
    End If
    value = valueattr.value

End Property


Public Property Let value(path As String, invalue)
    Dim valueelm As IXMLDOMElement
    Set valueelm = MessageDocument.selectSingleNode(path)
    If valueelm Is Nothing Then
        Err.Raise 1001, "cKAAMessage.Value get", "Δεν βρέθηκε η διαδρομή " & path
    End If
    
    Dim valueattr As IXMLDOMAttribute
    Set valueattr = valueelm.Attributes.getNamedItem("value")
    If valueattr Is Nothing Then
        Set valueattr = valueelm.ownerDocument.createAttribute("value")
        valueelm.Attributes.setNamedItem valueattr
    End If
    valueattr.value = invalue
    
End Property

Public Function UpdateValues(inMessageRoot As IXMLDOMElement, inDocRoot As IXMLDOMElement) As Boolean
UpdateValues = False
Dim elm As IXMLDOMElement, selm As IXMLDOMElement, scount As Integer
    
    If inMessageRoot Is Nothing Then Exit Function
    If inDocRoot.hasChildNodes Then
        For Each elm In inDocRoot.childNodes
            Dim apart As IXMLDOMElement
            Dim docnameattr As IXMLDOMAttribute
            Set docnameattr = elm.Attributes.getNamedItem("name")
            If (docnameattr Is Nothing) Then
                Set apart = inMessageRoot.selectSingleNode("./" & elm.baseName)
            Else
                Set apart = inMessageRoot.selectSingleNode("./" & elm.baseName & "[@name='" & docnameattr.value & "']")
            End If
                
'            'Ελεγχος αν υπάρχουν περισσότερα απο ενα childs με το ιδιο ονομα
'            scount = 1:
'            Set selm = elm.previousSibling
'            Do While Not (selm Is Nothing)
'                If selm.baseName = elm.baseName Then
'                    scount = scount + 1
'                Else
'                    Exit Do
'                End If
'            Loop
'            Dim apart As IXMLDOMElement
'            Set apart = inMessageRoot.selectSingleNode("./" & elm.baseName & "[" & scount - 1 & "]")
            If apart Is Nothing Then
                LogMsgbox "Δεν βρέθηκε το στοιχείο για το: " & elm.XML, vbCritical, "Λάθος στην προετοιμασία του KAAMessage"
                Exit Function
            Else
                If Not UpdateValues(apart, elm) Then Exit Function
            End If
        Next elm
    Else
        Dim viewvalueattr As IXMLDOMAttribute
        Dim datavalueattr As IXMLDOMAttribute
        Set viewvalueattr = inMessageRoot.Attributes.getNamedItem("value")
        Set datavalueattr = inDocRoot.Attributes.getNamedItem("value")
        If viewvalueattr Is Nothing Then
            Set viewvalueattr = inMessageRoot.ownerDocument.createAttribute("value")
            inMessageRoot.Attributes.setNamedItem viewvalueattr
        End If
        If Not (datavalueattr Is Nothing) Then
            viewvalueattr.value = datavalueattr.value
        ElseIf inDocRoot.Text <> "" Then
            viewvalueattr.value = inDocRoot.Text
        End If
    End If
    
    UpdateValues = True
End Function

