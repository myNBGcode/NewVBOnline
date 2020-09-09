VERSION 5.00
Begin VB.UserControl GenLabel 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   Enabled         =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   4800
   Begin VB.TextBox VControl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "GenLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Owner As Form
Public LabelNo As Integer, LabelName As String
Public ValidationControl As ScriptControl
Private DisplayFlag(10) As Boolean
Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer
Public DocX As Integer, DocY As Integer, DocWidth As Integer, DocHeight As Integer

Private OutCode(10) As Integer, OutBuffPos(10) As Integer, OutBuffLength(10) As Integer
Private InBuffPos(10) As Integer, InBuffLength(10) As Integer
Private JournalBeforeOut(10) As Boolean, JournalAfterIn(10) As Boolean


Public Function IsVisible(inPhase As Integer) As Boolean
    IsVisible = DisplayFlag(inPhase)
End Function

Public Function IsJournalBeforeOut(inPhase As Integer) As Boolean
    IsJournalBeforeOut = JournalBeforeOut(inPhase)
End Function

Public Function IsJournalAfterIN(inPhase As Integer) As Boolean
    IsJournalAfterIN = JournalAfterIn(inPhase)
End Function

Public Function GetOutCode(inPhase As Integer) As Integer
    GetOutCode = OutCode(inPhase)
End Function

Public Function GetOutBuffPos(inPhase As Integer) As Integer
    GetOutBuffPos = OutBuffPos(inPhase)
End Function

Public Function GetOutBuffLength(inPhase As Integer) As Integer
    GetOutBuffLength = OutBuffLength(inPhase)
End Function

Public Function GetInBuffPos(inPhase As Integer) As Integer
    GetInBuffPos = InBuffPos(inPhase)
End Function

Public Function GetInBuffLength(inPhase As Integer) As Integer
    GetInBuffLength = InBuffLength(inPhase)
End Function

Public Function WriteEJournal(inPhase As Integer, inTRNNum As Integer) As Boolean
    WriteEJournal = eJournalWriteFld(Owner, LabelNo, "", VControl.Text) ', CStr(inTRNNum), cTRNNum)
End Function

Public Property Get Text() As String
    Text = VControl.Text
End Property

Public Property Let Text(aText As String)
    VControl.Text = aText
    PropertyChanged "Text"
End Property

Public Sub UserControl_Hide()
    VControl.Visible = False
End Sub

Private Sub UserControl_Initialize()
    VControl.Visible = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub UserControl_Resize()
    VControl.Left = 0
    VControl.Top = 0
    VControl.width = width
    VControl.height = height
    
End Sub

Private Sub UserControl_Show()
    VControl.Visible = True
End Sub

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    If inNode Is Nothing Then
        DisplayFlag(inPhase) = True
        OutCode(inPhase) = 0
        OutBuffLength(inPhase) = 0
        OutBuffPos(inPhase) = 0
        InBuffLength(inPhase) = 0
        InBuffPos(inPhase) = 0
        JournalBeforeOut(inPhase) = False
        JournalAfterIn(inPhase) = False
    
        If inPhase = 1 Then
            VControl.Text = ""
            Set Owner = inOwner
    
            Set ValidationControl = inProcessControl
        End If
    Else
        DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", lblModel)
        OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", lblModel)
        OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", lblModel)
        OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", lblModel)
        InBuffLength(inPhase) = NodeIntegerFld(inNode, "InBuffLengthA", lblModel)
        InBuffPos(inPhase) = NodeIntegerFld(inNode, "InBuffPosA", lblModel)
        JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", lblModel)
        JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", lblModel)
    
        If inPhase = 1 Then
            VControl.Text = NodeStringFld(inNode, "Text", lblModel)
            Set Owner = inOwner
            LabelNo = NodeIntegerFld(inNode, "LabelNo", lblModel)
            LabelName = "Lbl" & StrPad_(CStr(LabelNo), 3, "0", "L")
    
            Set ValidationControl = inProcessControl
            ValidationControl.AddObject LabelName, Me, True
            
            ScrLeft = NodeIntegerFld(inNode, "ScrX", lblModel)
            ScrWidth = NodeIntegerFld(inNode, "ScrWidth", lblModel)
            ScrTop = NodeIntegerFld(inNode, "ScrY", lblModel) * 290
            ScrHeight = NodeIntegerFld(inNode, "ScrHeight", lblModel) * 285
    
            DocX = NodeIntegerFld(inNode, "DocX", lblModel)
            DocY = NodeIntegerFld(inNode, "DocY", lblModel)
            DocWidth = NodeIntegerFld(inNode, "DocWidth", lblModel)
            DocHeight = NodeIntegerFld(inNode, "DocHeight", lblModel)
    
            Select Case NodeIntegerFld(inNode, "ScrAlign", lblModel)
            Case 1
                VControl.Alignment = vbLeftJustify
            Case 2
                VControl.Alignment = vbRightJustify
            End Select
        End If
    End If
End Sub


Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim xml As DOMDocument30
Set xml = New DOMDocument30

Dim Elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set Elm = xml.createElement("LABEL")
    
    Set attr = xml.createAttribute("NO")
    attr.nodeValue = UCase(Me.LabelNo)
    Elm.setAttributeNode attr
    Set attr = xml.createAttribute("NAME")
    attr.nodeValue = UCase(Me.LabelName)
    Elm.setAttributeNode attr
    
    Set attr = xml.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(CInt(inPhase))) '(Owner.ProcessPhase))
    Elm.setAttributeNode attr
    
    Set attr = xml.createAttribute("TEXT")
    attr.nodeValue = UCase(Me.Text)
    Elm.setAttributeNode attr
    
    
    
    
    
    Set TranslateToProperties = Elm
End Function

Sub SetXMLValue(Elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    If inPhase = 0 Then inPhase = 1
    
    Me.Text = Elm.getAttributeNode("TEXT").nodeValue
    
End Sub




