VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLDocumentJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Title As String
Private mergedDoc As MSXML2.DOMDocument30
Private mergedxml As String, dirtyFlag As Boolean

Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement

Public Function ParseJob(ByRef rootElm As IXMLDOMElement) As Boolean
    Dim inputName As String, outputname As String
    Dim tranformationTitle As String, tranformationElm As IXMLDOMElement
    
    For Each Elm In rootElm.childNodes
        If Elm.baseName = "CLEAR" Then
            Set mergedDoc = New MSXML2.DOMDocument30: mergedxml = "": dirtyFlag = False
        ElseIf Elm.baseName = "INPUT" Then
            inputName = Elm.getAttribute("TITLE")
            mergedxml = mergedxml & Manager.XmlObjectList(inputName).xml: dirtyFlag = True
        ElseIf Elm.baseName = "FUNCTION" Then
            If dirtyFlag Then
                mergedDoc.loadXML "<" & Title & ">" & mergedDoc & "</" & Title & ">"
                dirtyFlag = False
            End If
            transformationTitle = Elm.getAttribute("TITLE")
            Set transformationElm = Manager.XmlObjectList(transformationTitle)
            Set mergedDoc = mergedDoc.documentElement.transformNode(transformationElm.selectSingleNode("./xsl:stylesheet"))
            mergedxml = mergedDoc.xml
        ElseIf Elm.baseName = "IF" Then
            If dirtyFlag Then
                mergedDoc.loadXML "<" & Title & ">" & mergedDoc & "</" & Title & ">"
                dirtyFlag = False
            End If
            Dim astatement As String
            astatement = Elm.getAttribute("select")
            If mergedDoc.selectSingleNode(astatement).Text = "true" Then
                Dim res As Boolean
                res = ParseJob(Elm)
            End If
        ElseIf Elm.baseName = "OUTPUT" Then
            outputname = Elm.getAttribute("TITLE")
            Manager.XmlObjectList(outputname).loadXML mergedDoc
        End If
    Next Elm
    ParseJob = True: Exit Function
    
End Function


Public Function BuildXMLView(Params As IXMLDOMElement) As IXMLDOMElement
    '<JOB TITLE="xxxxx">
    '   <INPUT TITLE="xxxx" />.....
    '   <FUNCTION TITLE="xxxx" />.....
    '   <OUTPUT TITLE="xxxx" />.....
    '</JOB>
    '
    Title = content.getAttribute("TITLE")
    
        
    
    
    
End Function

