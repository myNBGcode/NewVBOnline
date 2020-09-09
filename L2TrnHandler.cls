VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "L2TrnHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private backActiveL2TrnHandler As L2TrnHandler
Public filename As String
Public workDocument As MSXML2.DOMDocument30

Private FormUpdates As New Collection

Public Result As MSXML2.DOMDocument30

Public activeform
Public DocumentManager As cXMLDocumentManager

Public Sub addFormUpdate(Document As MSXML2.DOMDocument30, name As String)
    'Dim updatedoc As New MSXML2.DOMDocument30
    'updatedoc.LoadXml Document.XML
    
    FormUpdates.add Document, name
End Sub

Private Function PrepareL2Environment() As Boolean
    PrepareL2Environment = False
    
    On Error GoTo modelopenerrorpos
    If L2ModelFile Is Nothing Then
        Set L2ModelFile = New MSXML2.DOMDocument30
        L2ModelFile.Load ReadDir & "\l2model.xml"
    End If
    PrepareL2Environment = True
    Exit Function
modelopenerrorpos:
    LogMsgbox "Δεν βρέθηκε το αρχείο προδιαγραφών: " & ReadDir & "\l2model.xml" & " ΛΑΘΟΣ: " & Err.number & ", " & Err.description, vbOKOnly, "ΛΑΘΟΣ..."
End Function
Public Sub ExecuteForm(code As String)
Dim startupattribute As IXMLDOMAttribute
Dim formUpdate As MSXML2.DOMDocument30
Dim aStatus As Boolean

    If code = "9857" Or cDebug = 1 Then aStatus = True Else aStatus = ChkProfileAccessNew(CLng(code))
    If Not aStatus Then
        LogMsgbox "Δεν Επιτρέπεται η χρήση της συναλλαγής: " & (code), vbCritical
        Exit Sub
    End If
    If Not PrepareL2Environment Then Exit Sub
    Set backActiveL2TrnHandler = ActiveL2TrnHandler
    Set ActiveL2TrnHandler = Me
'    filename = TrnFileFromTrnCode(Me, code)
    Dim atrnnode As IXMLDOMElement
    Set atrnnode = TrnNodeFromTrnCode(code)
    If Not atrnnode Is Nothing Then
        filename = TrnFileFromTrnNode(Me, atrnnode)
    End If
    If filename <> "" Then
        Set workDocument = New MSXML2.DOMDocument30
        workDocument.Load filename
        If workDocument.parseError Is Nothing Then
        ElseIf workDocument.parseError.Reason <> "" Then
            LogMsgbox "Λαθος κατα το ανοιγμα του αρχείου:  " & filename & " " & workDocument.parseError.Reason & vbCrLf & workDocument.parseError.srcText, vbCritical, "ΛΑΘΟΣ...."
        End If
        If workDocument.documentElement Is Nothing Then
            Set startupattribute = Nothing
        Else
            Set startupattribute = workDocument.documentElement.Attributes.getNamedItem("startupjob")
        End If
    Else
        Set workDocument = Nothing
        Set startupattribute = Nothing
        LogMsgbox "Δεν βρέθηκε δήλωση για τη συναλλαγή " & CStr(code) & " στο αρχείο: L2TrnList ", vbCritical, "Λάθος..."
        Exit Sub
    End If
    
    Set DocumentManager = New cXMLDocumentManager
    Set DocumentManager.owner = Me
    
    DocumentManager.LoadTemplates workDocument.documentElement
    Dim oldtrncode
    oldtrncode = cTRNCode
    cTRNCode = code
    
    UpdateChiefKey ""
    UpdateManagerKey ""

    Dim i As Long
    For i = G0Data.count To 1 Step -1: G0Data.Remove (i): Next i
    
    Set activeform = New L2Form
    Set activeform.owner = Me
    For Each formUpdate In FormUpdates
        activeform.LoadXML formUpdate.XML
    Next formUpdate
    
    Dim exitformflag As Boolean
    exitformflag = False
    If startupattribute Is Nothing Then
    
    Else
        DocumentManager.XmlObjectList.Item(startupattribute.Text).XML
        If TypeOf DocumentManager.XmlObjectList.Item(startupattribute.Text) Is cXMLDocumentJob Then
            If DocumentManager.XmlObjectList.Item(startupattribute.Text).exitformflag Then
                exitformflag = True
            End If
        End If
    End If
    If Not exitformflag Then
        activeform.Show vbModal
    End If
    cTRNCode = oldtrncode
    
    Set ActiveL2TrnHandler = backActiveL2TrnHandler
End Sub

Public Function WebLinkWithNoExists(linkName As String) As String
On Error GoTo Err
    WebLinkWithNoExists = WebLink(linkName)
    Exit Function
Err:
    WebLinkWithNoExists = ""
End Function


Public Function WebLink(linkName As String) As String
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
        WebLink = WebLinks(UCase("EDUC" & linkName))
    ElseIf Left(Right(WorkEnvironment_, 8), 4) = "PROD" Then
        WebLink = WebLinks(UCase("PROD" & linkName))
    Else
        WebLink = ""
        MsgBox "Περιβάλλον λειτουργίας: " & WorkEnvironment_ & ". Δεν βρέθηκε το Virtual Directory:" & linkName
    End If
End Function

Public Sub LoadXML(invalue As String)
    If activeform Is Nothing Then
    Else
        activeform.LoadXML invalue
    End If
End Sub

Public Function XML() As String
    XML = activeform.XML
End Function

Private Sub Class_Terminate()
    CleanUp
End Sub

Public Sub CleanUp()
    Set backActiveL2TrnHandler = Nothing
    If Not FormUpdates Is Nothing Then
        While FormUpdates.count > 0
            FormUpdates.Remove 1
        Wend
        Set FormUpdates = Nothing
    End If
    Set Result = Nothing
    Set activeform = Nothing
    If Not (DocumentManager Is Nothing) Then
        DocumentManager.CleanUp
        Set DocumentManager = Nothing
    End If
    Set workDocument = Nothing
End Sub


