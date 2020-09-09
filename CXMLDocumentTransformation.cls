VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLDocumentTransformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement

Public Function LoadXml(value As String) As String
    Dim sourcedocument As New MSXML2.DOMDocument30
    Dim includeList As IXMLDOMNodeList
    Dim includeFile As String
    Dim includeAttr As IXMLDOMAttribute
    Dim i As Integer
    sourcedocument.LoadXml value
    Dim transformationDoc As New MSXML2.DOMDocument30
    If UCase(content.baseName) = "STYLESHEET" Then
        transformationDoc.LoadXml content.XML
    Else
        transformationDoc.LoadXml content.selectSingleNode("./xsl:stylesheet").XML
    End If
    
    Set includeList = content.SelectNodes(".//xsl:include")
    If Not (includeList Is Nothing) Then
       For i = 0 To includeList.length - 1
           Set includeAttr = includeList(i).Attributes.getNamedItem("href")
           If Not (includeAttr Is Nothing) Then
              includeFile = includeList(i).Attributes.getNamedItem("href").Text
              If includeFile <> "" Then
                 CopyINCLUDES (includeFile)
              End If
           End If
       Next
    End If
    On Error GoTo TransformationFailure
    LoadXml = sourcedocument.transformNode(transformationDoc)
    'sourcedocument.save "c:\res2.xml"
    'transformationDoc.save "c:\res3.xml"
    
    Exit Function
    
TransformationFailure:
    LogMsgbox "Απέτυχε ο μετασχηματισμός: " & Err.Number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
End Function

Public Function XML() As String

End Function

Private Sub CopyINCLUDES(ByVal includeFile As String)
    On Error GoTo ErrorHandler
    Dim include As New MSXML2.DOMDocument30
    include.Load (LogonServer & "\" & LogonShare & "\vbread\" & includeFile)
    If include.parseError.errorCode = 0 Then
       include.save (App.path & "\" & includeFile)
    End If
    'Dim fso
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'If fso.FileExists(LogonServer & "\vbonline\vbread\" & includeFile) Then
    '   fso.CopyFile LogonServer & "\vbonline\vbread\" & includeFile, App.path & "\" & includeFile
    'End If
    Exit Sub
ErrorHandler:

End Sub

Private Sub Class_Terminate()
    Set Manager = Nothing
    Set content = Nothing
End Sub
