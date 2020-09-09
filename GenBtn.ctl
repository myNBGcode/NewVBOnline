VERSION 5.00
Begin VB.UserControl GenBtn 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton Control 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
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
      Width           =   1095
   End
End
Attribute VB_Name = "GenBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As Form
Public BtnNo As Integer, BtnName As String, BtnName2 As String, LabelName As String, name As String

Private DisplayFlag(10) As Boolean, EditFlag(10) As Boolean, OptionalFlag(10) As Boolean

Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer

'Public Prompt As Label
Public DocX As Integer, DocY As Integer, DocWidth As Integer, DocHeight As Integer
Public Title As String
Public TitleX As Integer, TitleY As Integer, TitleWidth As Integer, TitleHeight As Integer
Public DocAlign As Integer

Private OutCode(10) As Integer, OutBuffPos(10) As Integer, OutBuffLength(10) As Integer
Private InBuffPos(10) As Integer, InBuffLength(10) As Integer
Private JournalBeforeOut(10) As Boolean, JournalAfterIn(10) As Boolean
Public QFldNo As Integer
Public PasswordChar As String

Public ValidOk As Boolean, ValidationError As String
Public FormatBeforeOutFlag As Boolean, FormatAfterInFlag As Boolean

Public DisplayMask As String, Editmask As String, OutMask As String, DocMask As String
Public EditLength As Integer, EditType As Integer
Public ValidationCode As Integer

Public ValidationControl As ScriptControl
Private ValidationFlag As Boolean
Private ClearText As String, OLDTEXT As String, OutBuffText As String, InBuffText As String, EnableEditChk As Boolean
Private ScrHelp As String
Public TTabIndex As Integer, RestoreState As Boolean

Public Property Get Prompt() As String
    Prompt = Control.Caption
End Property

Public Property Let Prompt(invalue As String)
    Control.Caption = invalue
End Property

Public Sub SetDisplay(inPhase, SetFlag)
    DisplayFlag(CInt(inPhase)) = CBool(SetFlag)
End Sub

Private Sub Control_Click()
    Dim aFlag As Boolean
    aFlag = False
    If Not Control.Enabled Then Exit Sub
    With owner
        RestoreState = True:  Control.Enabled = False
        .HandleEvent BtnName, 0, 0, aFlag
        Control.Enabled = RestoreState
    End With
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: .height = height
    End With
End Sub

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Public Property Get Enabled() As Boolean
    Enabled = Control.Enabled
End Property

Public Property Let Enabled(value As Boolean)
    Control.Enabled = value
End Property

Public Property Get Tag() As String
    Tag = Control.Tag
End Property

Public Property Let Tag(value As String)
    Control.Tag = value
End Property

Public Sub FinalizeEdit()

End Sub


Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", btnModel)
'    EditFlag(inPhase) = NodeBooleanFld(inNode, "ScrEntry", fldModel)
'    OptionalFlag(inPhase) = NodeBooleanFld(inNode, "ScrOptional", fldModel)
'    OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", fldModel)
'    OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", fldModel)
'    OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", fldModel)
'    InBuffLength(inPhase) = NodeIntegerFld(inNode, "InBuffLengthA", fldModel)
'    InBuffPos(inPhase) = NodeIntegerFld(inNode, "InBuffPosA", fldModel)
'    JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", fldModel)
'    JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", fldModel)

    
    If inPhase = 1 Then
        Set owner = inOwner
        BtnNo = NodeIntegerFld(inNode, "BtnNo", btnModel)
        BtnName = "Btn" & StrPad_(CStr(BtnNo), 3, "0", "L")
        BtnName2 = UCase(NodeStringFld(inNode, "Name", btnModel))
        If InStr(ReservedControlPrefixes, "," & Left(BtnName2, 3) & ",") > 0 Then BtnName2 = ""
        
'        If Left(BtnName2, 3) = "FLD" Or Left(BtnName2, 3) = "SPD" Or Left(BtnName2, 3) = "LST" _
        Or Left(BtnName2, 3) = "LBL" Or Left(BtnName2, 3) = "BTN" Or Left(BtnName2, 3) = "CMB" _
        Or Left(BtnName2, 3) = "CHK" Then BtnName2 = ""
        
        name = IIf(BtnName2 <> "", BtnName2, BtnName)

        Set ValidationControl = inProcessControl
        ValidationControl.AddObject BtnName, Me, True
        If Trim(BtnName2) <> "" And UCase(Trim(BtnName2)) <> UCase(Trim(BtnName)) Then
            On Error GoTo FldRegistrationError
            ValidationControl.ExecuteStatement "Set " & BtnName2 & "=" & BtnName
            GoTo FldRegistrationOk
FldRegistrationError:
            MsgBox "Λάθος κατα τη δήλωση του πεδίου: " & BtnName & ":" & BtnName2
FldRegistrationOk:
        End If
        
        
        ScrLeft = NodeIntegerFld(inNode, "ScrX", btnModel)
        ScrWidth = NodeIntegerFld(inNode, "ScrWidth", btnModel)
        ScrTop = NodeIntegerFld(inNode, "ScrY", btnModel) * 290
        ScrHeight = NodeIntegerFld(inNode, "ScrHeight", btnModel) * 285
        Control.Caption = NodeStringFld(inNode, "ScrPrompt", btnModel)

        TTabIndex = NodeIntegerFld(inNode, "TabIndex", btnModel)
        
NoSelections:

    End If
End Sub

Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim xml As DOMDocument30
Set xml = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = xml.createElement("BUTTON")
    Set attr = xml.createAttribute("NO")
    attr.nodeValue = UCase(Me.BtnNo)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("NAME")
    attr.nodeValue = UCase(Me.BtnName)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.BtnName2)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("ENABLED")
    attr.nodeValue = UCase(Me.Enabled)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(inPhase))
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("PROMPT")
    attr.nodeValue = UCase(Me.Prompt)
    elm.setAttributeNode attr
                    
    Set TranslateToProperties = elm
End Function

Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    Dim aattr As IXMLDOMAttribute
    If inPhase = 0 Then inPhase = 1
    For Each aattr In elm.Attributes
        Select Case aattr.baseName
            Case "PROMPT"
                Prompt = aattr.value
            Case "ENABLED"
                If aattr.value = "FALSE" Then
                    Enabled = False
                ElseIf aattr.value = "TRUE" Then
                    Enabled = True
                End If
            Case "VISIBLE"
                If aattr.value = "FALSE" Then
                    Me.SetDisplay inPhase, False
                ElseIf aattr.value = "TRUE" Then
                    Me.SetDisplay inPhase, True
                End If
        End Select
    Next aattr
End Sub

