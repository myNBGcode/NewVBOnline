Attribute VB_Name = "XMLSupport"
Option Explicit

Public Const TrnFrmNamespace = "http://www.nbg.gr/online/shine/TrnFrm/"


Public Function CreateXMLNode(Document As MSXML2.DOMDocument, namespace As String, nodename As String)
Dim mNamespace As IXMLDOMAttribute
Dim elm As IXMLDOMElement
    
    Set elm = Document.createElement(nodename)
    If namespace <> "none" Then
        Set mNamespace = Document.createAttribute("xmlns")
        mNamespace.value = namespace
        elm.setAttributeNode mNamespace
    End If
    Set CreateXMLNode = elm
    
    Set mNamespace = Nothing
    Set elm = Nothing
End Function


Public Function GetXmlNode(rootnode As IXMLDOMElement, path As String, _
    Optional partname As String, Optional partcontainer As String, Optional messagetitle As String) As IXMLDOMElement
    Set GetXmlNode = rootnode.selectSingleNode(path)
    If GetXmlNode Is Nothing Then
        Dim Message As String
        Message = "Δεν Βρέθηκε το τμήμα "
        If Not IsMissing(partname) Then Message = Message & partname Else Message = Message & path
        If Not IsMissing(partcontainer) Then Message = Message & " στο " & partcontainer
        If Not IsMissing(messagetitle) Then MsgBox Message, True, messagetitle Else MsgBox Message, True
    End If
End Function
Public Function GetXmlNodeIfPresent(rootnode As IXMLDOMElement, path As String) As IXMLDOMElement
    Set GetXmlNodeIfPresent = rootnode.selectSingleNode(path)
End Function
Public Function XmlLoadFile(filename As String, _
    Optional documentname As String, Optional messagetitle As String) As MSXML2.DOMDocument30
    Dim adoc As MSXML2.DOMDocument30
    Set adoc = New MSXML2.DOMDocument30
    
    With adoc
        .Load filename
        If Not (.parseError Is Nothing) Then
            If (.parseError.errorCode <> 0) Then
                Dim Message As String
                Message = "Προβλημα στη δημιουργία του "
                If Not IsMissing(documentname) Then Message = Message & documentname Else Message = Message & filename
                Message = Message & ": " & .parseError.errorCode & " " & _
                    .parseError.reason & " Θέση: " & .parseError.filepos & " Γραμμή: " & _
                    .parseError.Line & " Θέση Γραμμής: " & .parseError.linepos
                If Not IsMissing(messagetitle) Then MsgBox Message, True, messagetitle Else MsgBox Message, True
                Exit Function
            End If
        End If
    End With
    Set XmlLoadFile = adoc
End Function

Public Function XmlLoadString(Data As String, _
    Optional documentname As String, Optional messagetitle As String) As MSXML2.DOMDocument30
    Dim adoc As MSXML2.DOMDocument30
    Set adoc = New MSXML2.DOMDocument30
    
    With adoc
        .LoadXml Data
        If Not (.parseError Is Nothing) Then
            If (.parseError.errorCode <> 0) Then
                Dim Message As String
                Message = "Προβλημα στη δημιουργία του "
                If Not IsMissing(documentname) Then Message = Message & documentname Else Message = Message & " document"
                Message = Message & ": " & .parseError.errorCode & " " & _
                    .parseError.reason & " Θέση: " & .parseError.filepos & " Γραμμή: " & _
                    .parseError.Line & " Θέση Γραμμής: " & .parseError.linepos
                If Not IsMissing(messagetitle) Then MsgBox Message, True, messagetitle Else MsgBox Message, True
                Exit Function
            End If
        End If
    End With
    Set XmlLoadString = adoc
End Function

Public Function AsDate(FormatedText As String) As Date
'p.x. 21/01/2005
    FormatedText = Replace(FormatedText, ".", "/")
    If Not IsDate(FormatedText) Then
        AsDate = CDate("1900/01/01")
    ElseIf CInt(Left(FormatedText, 2)) <> Day(CDate(FormatedText)) Then
        AsDate = CDate("1900/01/01")
    ElseIf CInt(Mid(FormatedText, 4, 2)) <> Month(CDate(FormatedText)) Then
        AsDate = CDate("1900/01/01")
    Else
        AsDate = CDate(FormatedText)
    End If
End Function

