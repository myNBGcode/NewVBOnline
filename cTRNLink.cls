VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTRNLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement
Private TrnCode As String
Private LinkMode As String, LinkStringParams As String
Private OutputViewSystem As String, OutputViewModel As String, OutputViewName As String

'<trnlink name="findcustomer">
'   <trn name="7759">
'   <linkmode name="linktotrnv2" params="WCACU20BL2USE">
'   <outputview system="cra" model="WCACU20B" name="WCACU20BL2USE">
'</trnlink>

Public Function LoadXML(inDocument As String) As String

    Dim workDocument As New MSXML2.DOMDocument30
    workDocument.LoadXML inDocument
    Dim elm As IXMLDOMElement, attr As IXMLDOMAttribute
    Dim PArray()
    
    For Each elm In workDocument.documentElement.childNodes
        If UCase(elm.baseName) = "TRN" Then
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then TrnCode = attr.Text
        ElseIf UCase(elm.baseName) = "LINKMODE" Then
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then LinkMode = attr.Text
            If UCase(LinkMode) = "LINKTOTRNV2" Then
                Set attr = elm.getAttributeNode("params")
                If Not (attr Is Nothing) Then LinkStringParams = attr.Text
            End If
        ElseIf UCase(elm.baseName) = "INPUTVIEW" Then
            If UCase(LinkMode) = "LINKTOTRNV3" Then
                If elm.childNodes.length > 0 Then
                    ReDim PArray(elm.childNodes.length)
                    Dim ParamNode As IXMLDOMElement, pos As Integer
                    pos = 0
                    For Each ParamNode In elm.childNodes
                        PArray(pos) = ParamNode.Text
                        pos = pos + 1
                    Next ParamNode
                End If
            ElseIf UCase(LinkMode) = "LINKTOTRNV4" Then
                If elm.childNodes.length > 0 Then
                    ReDim PArray(elm.childNodes.length)
                    pos = 0
                    For Each ParamNode In elm.childNodes
                        If ParamNode.hasChildNodes Then
                           If ParamNode.childNodes(0).nodeType = NODE_ELEMENT Then
                              PArray(pos) = ParamNode.XML
                           Else
                             PArray(pos) = ParamNode.Text
                           End If
                        Else
                           PArray(pos) = ParamNode.Text
                        End If
                        pos = pos + 1
                    Next ParamNode
                End If
            End If
            
        ElseIf UCase(elm.baseName) = "OUTPUTVIEW" Then
            Set attr = elm.getAttributeNode("system")
            If Not (attr Is Nothing) Then OutputViewSystem = attr.Text
            Set attr = elm.getAttributeNode("model")
            If Not (attr Is Nothing) Then OutputViewModel = attr.Text
            Set attr = elm.getAttributeNode("name")
            If Not (attr Is Nothing) Then OutputViewName = attr.Text
        ElseIf UCase(elm.baseName) = "INPUTPARAM" Then
        
        End If
    Next elm
    
    Dim BuildViewFlag As Boolean, FindViewFlag As Boolean
    If OutputViewModel <> "" And OutputViewName <> "" Then
        If GenWorkForm.AppBuffers.Exists(OutputViewName) Then
            FindViewFlag = True
            GenWorkForm.AppBuffers.ByName(OutputViewName).ClearData
        Else
            If UCase(OutputViewSystem) = "CRA" Then
                BuildViewFlag = BuildCRAAppStruct(OutputViewModel, OutputViewName, True)
            ElseIf UCase(OutputViewSystem) = "IRIS" Then
                BuildViewFlag = BuildIRISAppStruct(OutputViewModel, OutputViewName, True)
            End If
        End If
    End If
    
    Dim oldTrnCD As String, anewTrnFrm As TRNFrm, oldFlag As Boolean
    If UCase(LinkMode) = "LINKTOTRNV2" And TrnCode <> "" Then
        oldTrnCD = cTRNCode: cTRNCode = TrnCode: oldFlag = cEnableHiddenTransactions: cEnableHiddenTransactions = True
        Set anewTrnFrm = New TRNFrm
        With anewTrnFrm
            Set .OwnerForm = Nothing
            .Params = LinkStringParams
            On Error Resume Next: .Show vbModal
        End With
        cTRNCode = oldTrnCD: cEnableHiddenTransactions = oldFlag
        Set anewTrnFrm = Nothing
    End If
    If (UCase(LinkMode) = "LINKTOTRNV3" Or UCase(LinkMode) = "LINKTOTRNV4") And TrnCode <> "" Then
        oldTrnCD = cTRNCode: cTRNCode = TrnCode: oldFlag = cEnableHiddenTransactions: cEnableHiddenTransactions = True
        Set anewTrnFrm = New TRNFrm
        With anewTrnFrm
            Set .OwnerForm = Nothing
            .PArray = PArray
            On Error Resume Next: .Show vbModal
        End With
        cTRNCode = oldTrnCD: cEnableHiddenTransactions = oldFlag
        Set anewTrnFrm = Nothing
    End If
    If BuildViewFlag Or FindViewFlag Then
        LoadXML = GenWorkForm.AppBuffers.ByName(OutputViewName).GetXMLView.XML
    End If
End Function

Public Function XML() As String
    If content Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(content.XML)
    End If
End Function

