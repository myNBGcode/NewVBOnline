VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.UserControl GenChart 
   BackColor       =   &H80000009&
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   DataSourceBehavior=   1  'vbDataSource
   ScaleHeight     =   1995
   ScaleWidth      =   1995
   Begin MSChart20Lib.MSChart Control 
      Height          =   1995
      Left            =   0
      OleObjectBlob   =   "GenChart.ctx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "GenChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Owner As Form
Public ChartNo As Integer, ChartName As String, LabelName As String
Public Prompt As Label
Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer
Private DisplayFlag(10) As Boolean
Private ValidationControl As ScriptControl

Public DocX As Integer, DocY As Integer, DocWidth As Integer, DocHeight As Integer
Public Title As String
Public TitleX As Integer, TitleY As Integer, TitleWidth As Integer, TitleHeight As Integer
Public DocAlign As Integer

'Private GetLinesFromBuffer As Boolean
'Private MinInLineLength As Integer, FromColumn As Integer, ToColumn As Integer
'Private ExcludeFirstLines As Integer, ExcludeLastLines As Integer

'Private DocTopMargin, DocLeftMargin, DocRightMargin, DocBottomMargin As Integer
'Private DocOrientation As Integer
'Private DocPageHead, DocPageFoot As String
'Private DocHeadLines, DocFootLines, DocClearLines As Integer
'Private DocLines, DocColumns As Integer

'Private OutCode(10) As Integer, OutBuffPos(10) As Integer, OutBuffLength(10) As Integer
'Private InBuffPos(10) As Integer, InBuffLength(10) As Integer
'Private JournalBeforeOut(10) As Boolean, JournalAfterIn(10) As Boolean
'Public QFldNo As Integer
'Public PasswordChar As String

Public ValidOk As Boolean, ValidationError As String
'Public FormatBeforeOutFlag As Boolean, FormatAfterInFlag As Boolean

'Public DisplayMask As String, Editmask As String, OutMask As String, DocMask As String
'Public EditLength As Integer, EditType As Integer
Public ValidationCode As Integer


Private ValidationFlag As Boolean
'Private ClearText As String, OLDTEXT As String, OutBuffText As String, InBuffText As String, EnableEditChk As Boolean
'Private ScrHelp As String
Public TTabIndex As Integer
'RestoreState As Boolean




'Private FormatString As String
'Dim ColumnPairs(30) As IntegerPair
Dim PairsCount As Integer
Private ExamineLineFlag As Boolean
Private curPageNo As Integer
'Private ERowsCount As Long

'Public PrintTruncatedCells As Boolean


Private Sub UserControl_Resize()
    
        Control.Left = 0:
        Control.Top = 0:
        Control.width = width:
        Control.height = height
End Sub

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Public Property Get Enabled() As Boolean
    Enabled = Control.Enabled
End Property

Public Property Let Enabled(Value As Boolean)
    Control.Enabled = Value
End Property

Public Property Get Tag() As String
    Tag = Control.Tag
End Property

Public Property Let Tag(Value As String)
    Control.Tag = Value
End Property

Public Property Get chartType() As Integer
  chartType = Control.chartType
End Property

Public Property Let chartType(Value As Integer)
  Control.chartType = Value
End Property

Public Property Get ChartData() As Variant
  ChartData = Control.ChartData
End Property


Public Property Let ChartData(Value As Variant)
Dim A() As Variant
ReDim Á(LBound(Value, 1) To UBound(Value, 1), LBound(Value, 2) To UBound(Value, 2))
A = Value
Control.ChartData = A
End Property

Public Sub LetChartData(Value As Variant)

Control.ChartData = Value
'Dim a(1 To 3)
'Dim i As Integer


' For i = 1 To 3
 ' a(i) = ar(i - 1)
  'MsgBox a(i)
'Next i

End Sub

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", chrModel)
    'EditFlag(inPhase) = NodeBooleanFld(inNode, "ScrEntry", fldModel)
    'OptionalFlag(inPhase) = NodeBooleanFld(inNode, "ScrOptional", fldModel)
    'OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", fldModel)
    'OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", fldModel)
    'OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", fldModel)
    'InBuffLength(inPhase) = NodeIntegerFld(inNode, "InBuffLengthA", fldModel)
    'InBuffPos(inPhase) = NodeIntegerFld(inNode, "InBuffPosA", fldModel)
    'JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", fldModel)
    'JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", fldModel)

    If inPhase = 1 Then
        Set Owner = inOwner
        ChartNo = NodeIntegerFld(inNode, "ChartNo", chrModel)
        ChartName = "Chr" & StrPad_(CStr(ChartNo), 3, "0", "L") 'NodeStringFld(inNode, "Name", fldModel)
        'TotalName = NodeStringFld(inNode, "TotalName", fldModel)
        'TotalPos = NodeIntegerFld(inNode, "TotalPos", fldModel)

        'aType = NodeIntegerFld(inNode, "ChrType", fldModel)
        'If aType = 0 Then
        '    Set aTypeNode = Nothing
        'Else
        '    Set aTypeNode = FldTypeList.Item("T" & Trim(Str(aType)))
        'End If
        
        chartType = NodeIntegerFld(inNode, "ChartType", chrModel)
        
        Set ValidationControl = inProcessControl
        ValidationControl.AddObject ChartName, Me, True
        
        ScrLeft = NodeIntegerFld(inNode, "ScrX", chrModel)
        ScrWidth = NodeIntegerFld(inNode, "ScrWidth", chrModel)
        ScrTop = NodeIntegerFld(inNode, "ScrY", chrModel) * 290
        ScrHeight = NodeIntegerFld(inNode, "ScrHeight", chrModel) * 285
        'ScrHelp = NodeStringFld(inNode, "ScrHelp", fldModel)
        
        DocX = NodeIntegerFld(inNode, "DocX", chrModel)
        DocY = NodeIntegerFld(inNode, "DocY", chrModel)
        DocWidth = NodeIntegerFld(inNode, "DocWidth", chrModel)
        DocHeight = NodeIntegerFld(inNode, "DocHeight", chrModel)
        Title = NodeStringFld(inNode, "DocTitle", chrModel)
        TitleX = NodeIntegerFld(inNode, "DocTitleX", chrModel)
        TitleY = NodeIntegerFld(inNode, "DocTitleY", chrModel)
        TitleWidth = NodeIntegerFld(inNode, "DocTitleWidth", chrModel)
        TitleHeight = NodeIntegerFld(inNode, "DocTitleHeight", chrModel)
        
        TTabIndex = NodeIntegerFld(inNode, "TabIndex", chrModel)


        'Select Case NodeIntegerFld(inNode, "DocAlign", fldModel)
        'Case 1
        '    DocAlign = 1
        'Case 2
        '    DocAlign = 2
        'End Select
        
        'If Not (aTypeNode Is Nothing) Then
        '    If aTypeNode.childnodes.Item("ValidationCode").Text <> "" Then
        '        ValidationCode = aTypeNode.childnodes.Item("ValidationCode").Text
        '    Else
        '        ValidationCode = 0
        '    End If
            
        'End If
        
        LabelName = "CLabel" & StrPad_(CStr(ChartNo), 3, "0", "L")
'        NodeStringFld(inNode, "Name", fldModel) & _
'            "_" & NodeStringFld(inNode, "Phase", fldModel)
        Set Prompt = Parent.Controls.Add("Vb.Label", LabelName)
        Prompt.BackColor = Parent.BackColor
        Prompt.AutoSize = False
        Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", chrModel)
        
        Parent.ScaleMode = vbCharacters
        
        Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", chrModel)
        Prompt.width = NodeIntegerFld(inNode, "ScrPromptWidth", chrModel)
        Parent.ScaleMode = vbTwips
        Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", chrModel) * 290
        Prompt.height = NodeIntegerFld(inNode, "ScrPromptHeight", chrModel) * 285
    
        'Dim astr As String
        'astr = NodeStringFld(inNode, "ScrValidationScript", fldModel)
        'If astr <> "" Then
        '    ValidationControl.AddCode " Public Sub " & _
        '        FldName & "_Validation " & vbCrLf & _
        '        astr & vbCrLf & "End Sub"
        '    ValidationFlag = True
        'End If
        'astr = NodeStringFld(inNode, "FormatBeforeOutScript", fldModel)
        'If astr <> "" Then
        '    ValidationControl.AddCode " Public Sub " & _
        '        FldName & "_FormatBeforeOut " & vbCrLf & _
        '        astr & vbCrLf & "End Sub"
        '    FormatBeforeOutFlag = True
        'End If
        'astr = NodeStringFld(inNode, "FormatAfterInScript", fldModel)
        'If astr <> "" Then
        '    ValidationControl.AddCode " Public Sub " & _
        '        FldName & "_FormatAfterIn " & vbCrLf & _
        '        astr & vbCrLf & "End Sub"
        '    FormatAfterInFlag = True
        'End If
        
        'Editmask = "": DisplayMask = "": OutMask = ""
        'If Not (aTypeNode Is Nothing) Then
        '    DisplayMask = aTypeNode.childnodes.Item("DisplayMask").Text
        '    Editmask = aTypeNode.childnodes.Item("EditMask").Text
        '    OutMask = aTypeNode.childnodes.Item("OutMask").Text
        '    EditLength = aTypeNode.childnodes.Item("EditLength").Text
        '    EditType = aTypeNode.childnodes.Item("EditType").Text
        'End If
        
        
        'If NodeStringFld(inNode, "ScrDisplayMask", fldModel) <> "" Then _
        '    DisplayMask = NodeStringFld(inNode, "ScrDisplayMask", fldModel)
        'If NodeStringFld(inNode, "ScrEditMask", fldModel) <> "" Then _
        '    Editmask = NodeStringFld(inNode, "ScrEditMask", fldModel)
        'If NodeStringFld(inNode, "OutMask", fldModel) <> "" Then _
        '    OutMask = NodeStringFld(inNode, "OutMask", fldModel)
        'DocMask = DisplayMask
        'If NodeStringFld(inNode, "DocDisplayMask", fldModel) <> "" Then _
        '    DocMask = NodeStringFld(inNode, "DocDisplayMask", fldModel)
        
'        NodeIntegerFld(inNode, "", fldModel)
'        NodeStringFld(inNode, "", fldModel)
        
        
        'QFldNo = NodeIntegerFld(inNode, "QFldNo", fldModel)
        'PasswordChar = NodeStringFld(inNode, "ScrPasswordChar", fldModel)
        'If PasswordChar <> "" Then
        '    VControl.Font.Name = "BBSecret"
            'VControl.Font.Name = "Wingdings"
            
'            VControl.PasswordChar = PasswordChar
'            VControl.ForeColor = VControl.BackColor
        'End If
        
'NoSelections:

    End If
End Sub


Private Sub UserControl_Show()
   If Not (Prompt Is Nothing) Then Prompt.Visible = True
End Sub

Public Function TranslateToProperties(inPhase) As IXMLDOMElement
Dim xml As DOMDocument30
Set xml = New DOMDocument30

Dim Elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set Elm = xml.createElement("CHART")
    Set attr = xml.createAttribute("NO")
    attr.nodeValue = UCase(Me.ChartNo)
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("NAME")
    attr.nodeValue = UCase(Me.ChartName)
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.ChartName)
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("ENABLED")
    attr.nodeValue = UCase(Me.Enabled)
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(inPhase))
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("TITLE")
    attr.nodeValue = UCase(Me.Title)
    Elm.setAttributeNode attr

End Function
Sub SetXMLValue(Elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    Me.Title = Elm.getAttributeNode("TITLE").nodeValue
    Me.ChartName = Elm.getAttributeNode("FULLNAME").nodeValue
    Me.Enabled = Elm.getAttributeNode("ENABLED").nodeValue
    'Me.SetDisplay inPhase, Elm.getAttributeNode("VISIBLE").nodeValue
End Sub

