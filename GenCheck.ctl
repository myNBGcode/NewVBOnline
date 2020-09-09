VERSION 5.00
Begin VB.UserControl GenCheck 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox Control 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "GenCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As Form
Public ChkNo As Integer, ChkName As String, CHKName2 As String, LabelName As String, name As String
Private DisplayFlag(10) As Boolean, EditFlag(10) As Boolean, OptionalFlag(10) As Boolean

Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer

Public DocX As Integer, DocY As Integer, DocWidth As Integer, DocHeight As Integer
Public Title As String
Public TitleX As Integer, TitleY As Integer, TitleWidth As Integer, TitleHeight As Integer
Public DocAlign As Integer

Private OutCode(10) As Integer, OutBuffPos(10) As Integer, OutBuffLength(10) As Integer
Private InBuffPos(10) As Integer, InBuffLength(10) As Integer
Private JournalBeforeOut(10) As Boolean, JournalAfterIn(10) As Boolean
Public QFldNo As Integer

Public ValidOk As Boolean, ValidationError As String
Public FormatBeforeOutFlag As Boolean, FormatAfterInFlag As Boolean

Public OutMask As String, DocMask As String
Public HPSOutStruct As String, HPSOutPart As String, HPSInStruct As String, HPSInPart As String

Public ValidationControl As ScriptControl
Private ValidationFlag As Boolean
Private ClearText As String, OLDTEXT As String, OutBuffText As String, InBuffText As String, EnableEditChk As Boolean
Private ScrHelp As String
Public TTabIndex As Integer

Private Sub Control_Click()
    Dim aFlag As Boolean
    aFlag = False
    With owner
        Control.Enabled = False
        .HandleEvent ChkName, 0, 0, aFlag
        Control.Enabled = True
    End With
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If owner.TabbedControls.count > 1 Then SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: .height = height
    End With
End Sub

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Public Function IsEditable(inPhase) As Boolean
    IsEditable = EditFlag(CInt(inPhase))
End Function

Public Function SetEditable(inPhase, SetFlag) As Boolean
    EditFlag(CInt(inPhase)) = CBool(SetFlag)
    HandleEdit inPhase
    owner.RefreshView
End Function

Public Function SetEditableNoRefresh(inPhase, SetFlag) As Boolean
    EditFlag(CInt(inPhase)) = CBool(SetFlag)
    HandleEdit inPhase
End Function

Public Sub HandleEdit(inPhase)
    Control.TabStop = EditFlag(inPhase)
    Control.Enabled = EditFlag(inPhase)
End Sub

Public Sub SetAsReadOnly()
    Control.TabStop = False
    Control.Enabled = False
    Control.BackColor = &HD0D0D0
End Sub

Public Sub FinalizeEdit()

End Sub

Public Property Get value() As Integer
    value = Control.value
End Property

Public Property Let value(invalue As Integer)
    Control.value = invalue
    PropertyChanged "Value"
End Property

Public Property Get Enabled() As Boolean
    Enabled = Control.Enabled
End Property

Public Property Let Enabled(invalue As Boolean)
    Control.Enabled = invalue
End Property

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", chkModel)
    EditFlag(inPhase) = NodeBooleanFld(inNode, "ScrEntry", chkModel)
    OptionalFlag(inPhase) = NodeBooleanFld(inNode, "ScrOptional", chkModel)
    OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", chkModel)
    OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", chkModel)
    OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", chkModel)
    InBuffLength(inPhase) = NodeIntegerFld(inNode, "InBuffLengthA", chkModel)
    InBuffPos(inPhase) = NodeIntegerFld(inNode, "InBuffPosA", chkModel)
    JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", chkModel)
    JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", chkModel)

    If inPhase = 1 Then
        Set owner = inOwner
        ChkNo = NodeIntegerFld(inNode, "ChkNo", chkModel)
        ChkName = "Chk" & StrPad_(CStr(ChkNo), 3, "0", "L")
        CHKName2 = UCase(NodeStringFld(inNode, "NAME", chkModel))
        If InStr(ReservedControlPrefixes, "," & Left(CHKName2, 3) & ",") > 0 Then CHKName2 = ""
        
'        If Left(CHKName2, 3) = "FLD" Or Left(CHKName2, 3) = "SPD" Or Left(CHKName2, 3) = "LST" _
        Or Left(CHKName2, 3) = "LBL" Or Left(CHKName2, 3) = "BTN" Or Left(CHKName2, 3) = "CMB" _
        Or Left(CHKName2, 3) = "CHK" Then CHKName2 = ""
        name = IIf(CHKName2 <> "", CHKName2, ChkName)

        Set ValidationControl = inProcessControl
        ValidationControl.AddObject ChkName, Me, True
        If Trim(CHKName2) <> "" And UCase(Trim(CHKName2)) <> UCase(Trim(ChkName)) Then
            On Error GoTo FldRegistrationError
            ValidationControl.ExecuteStatement "Set " & CHKName2 & "=" & ChkName
            GoTo FldRegistrationOk
FldRegistrationError:
            MsgBox "Λάθος κατα τη δήλωση του πεδίου: " & ChkName & ":" & CHKName2
FldRegistrationOk:
        End If
        
        ScrLeft = NodeIntegerFld(inNode, "ScrX", chkModel)
        ScrWidth = NodeIntegerFld(inNode, "ScrWidth", chkModel)
        ScrTop = NodeIntegerFld(inNode, "ScrY", chkModel) * 290
        ScrHeight = NodeIntegerFld(inNode, "ScrHeight", chkModel) * 285
        ScrHelp = NodeStringFld(inNode, "ScrHelp", chkModel)

        DocX = NodeIntegerFld(inNode, "DocX", chkModel)
        DocY = NodeIntegerFld(inNode, "DocY", chkModel)
        DocWidth = NodeIntegerFld(inNode, "DocWidth", chkModel)
        DocHeight = NodeIntegerFld(inNode, "DocHeight", chkModel)
        Title = NodeStringFld(inNode, "DocTitle", chkModel)
        TitleX = NodeIntegerFld(inNode, "DocTitleX", chkModel)
        TitleY = NodeIntegerFld(inNode, "DocTitleY", chkModel)
        TitleWidth = NodeIntegerFld(inNode, "DocTitleWidth", chkModel)
        TitleHeight = NodeIntegerFld(inNode, "DocTitleHeight", chkModel)

        Control.Caption = NodeStringFld(inNode, "ScrPrompt", chkModel)
        
        TTabIndex = NodeIntegerFld(inNode, "TabIndex", chkModel)
        
        OutMask = ""
        If Not (aTypeNode Is Nothing) Then
            OutMask = aTypeNode.childNodes.item("OutMask").Text
        End If
        
        HPSOutStruct = NodeStringFld(inNode, "HPSOutStruct", chkModel)
        HPSOutPart = NodeStringFld(inNode, "HPSOutPart", chkModel)
        HPSInStruct = NodeStringFld(inNode, "HPSInStruct", chkModel)
        HPSInPart = NodeStringFld(inNode, "HPSInPart", chkModel)
        
        QFldNo = NodeIntegerFld(inNode, "QFldNo", chkModel)
        
NoSelections:

    End If
End Sub


Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim xml As DOMDocument30
Set xml = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = xml.createElement("CHECK")
    Set attr = xml.createAttribute("NO")
    attr.nodeValue = UCase(Me.ChkNo)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("NAME")
    attr.nodeValue = UCase(Me.ChkName)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("TITLE")
    attr.nodeValue = UCase(Me.Title)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("VALUE")
    attr.nodeValue = UCase(Me.value)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("READONLY")
    attr.nodeValue = UCase(Not Me.IsEditable(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
                    
    Set TranslateToProperties = elm
End Function

Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    If inPhase = 0 Then inPhase = 1
    Me.Title = elm.getAttributeNode("TITLE").nodeValue
    'Me.IsVisible inphase, elm.getAttributeNode("VISIBLE").nodeValue
    
    Me.value = elm.getAttributeNode("VALUE").nodeValue
    Me.CHKName2 = elm.getAttributeNode("FULLNAME").nodeValue
    Me.SetEditable inPhase, Not CBool(elm.getAttributeNode("READONLY").nodeValue)
    
End Sub

