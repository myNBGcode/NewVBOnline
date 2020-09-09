VERSION 5.00
Begin VB.UserControl GenCombo 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox Control 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "GenCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As Form
Public CmbNo As Integer, CmbName As String, CMBName2 As String, LabelName As String, name As String
Private DisplayFlag(10) As Boolean, EditFlag(10) As Boolean, OptionalFlag(10) As Boolean

Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer

Public Prompt As Label
Public DocX As Integer, DocY As Integer, DocWidth As Integer, DocHeight As Integer
Public Title As String
Public TitleX As Integer, TitleY As Integer, TitleWidth As Integer, TitleHeight As Integer
Public DocAlign As Integer

Private OutCode(10) As Integer, OutBuffPos(10) As Integer, OutBuffLength(10) As Integer, OutCodeEx(10) As String
Private InBuffPos(10) As Integer, InBuffLength(10) As Integer
Private JournalBeforeOut(10) As Boolean, JournalAfterIn(10) As Boolean
Public QFldNo As Integer
Public PasswordChar As String

'Public ValidOk As Boolean, ValidationError As String
Public FormatBeforeOutFlag As Boolean, FormatAfterInFlag As Boolean

Public DisplayMask As String, Editmask As String, OutMask As String, DocMask As String
Public EditLength As Integer, EditType As Integer
Public ValidationCode As Integer

Public ValidationControl As ScriptControl
Private ValidationFlag As Boolean
Private ClearText As String, OLDTEXT As String, OutBuffText As String, InBuffText As String, EnableEditChk As Boolean
Private ScrHelp As String

'------------------------------------------------------------------------------------------------------------------------------------
Public TTabIndex As Integer
Public LastListIndex As Long

Private Choices() As HelpLine
Private ChoiceCount_ As Integer

Private ChoicesSuperSet() As HelpLine
Private ChoiceSuperSetCount As Integer

Private Sub Control_Change()
    Dim aFlag As Boolean
    aFlag = False
    With owner
        .HandleEvent CmbName, Control.ListIndex, 0, aFlag
    End With
    LastListIndex = Control.ListIndex
End Sub

Private Sub Control_Click()
    Dim aFlag As Boolean
    aFlag = False
    If Not (owner Is Nothing) Then
        With owner
            .HandleEvent CmbName, Control.ListIndex, 0, aFlag
        End With
        LastListIndex = Control.ListIndex
    End If
End Sub

Private Sub Control_GotFocus()
    If Control.Locked And owner.TabbedControls.count > 1 Then SendKeys "{TAB}"
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If owner.TabbedControls.count > 1 Then SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Control_Validate(Cancel As Boolean)
    Dim aFlag As Boolean
    aFlag = False
    With owner
        .HandleEvent CmbName, Control.ListIndex, 1, aFlag
    End With
    LastListIndex = Control.ListIndex
End Sub

Private Sub UserControl_Resize()
    With Control
        .Left = 0: .Top = 0: .width = width: height = .height
    End With
End Sub

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Public Function IsEditable(inPhase) As Boolean
    IsEditable = EditFlag(CInt(inPhase))
End Function

Public Sub SetDisplay(inPhase, SetFlag)
    DisplayFlag(CInt(inPhase)) = CBool(SetFlag)
End Sub

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
    Control.Locked = Not EditFlag(inPhase)
    If Not EditFlag(inPhase) Then
        Control.BackColor = &HD0D0D0
    Else
        Control.BackColor = &H80000005
    End If
End Sub

Public Function GetOutCode(inPhase) As Integer
    GetOutCode = OutCode(CInt(inPhase))
End Function

Public Function GetOutCodeEx(inPhase) As String
    GetOutCodeEx = Trim(OutCodeEx(CInt(inPhase)))
End Function

Public Function GetOutBuffPos(inPhase) As Integer
    GetOutBuffPos = OutBuffPos(CInt(inPhase))
End Function

Public Function GetOutBuffLength(inPhase) As Integer
    GetOutBuffLength = OutBuffLength(CInt(inPhase))
End Function

Public Sub FormatBeforeOut()
    Dim astr As String
    If Choice <> "" And OutMask <> "" Then
        OutBuffText = format(Choice, OutMask)
    Else
        OutBuffText = Choice
    End If
    
End Sub

Public Function IsOptional(inPhase) As Boolean
    IsOptional = OptionalFlag(CInt(inPhase))
End Function

Public Property Get OutText() As String
    OutText = OutBuffText
End Property

Public Sub AddItem(value, Optional Index)
    On Error GoTo GenError
    If IsMissing(Index) Then Control.AddItem CStr(value) Else Control.AddItem CStr(value), Index
    Exit Sub
GenError:
    MsgBox "Πρόβλημα στην προσθήκη εγγραφής στο Combo Box: " & name, vbCritical, "ΛΑΘΟΣ"
End Sub

Public Property Get ListIndex() As Integer
    On Error GoTo GenError
    ListIndex = Control.ListIndex
    Exit Property
GenError:
    MsgBox "Πρόβλημα στην ανάκτηση ListIndex απο το Combo Box: " & name, vbCritical, "ΛΑΘΟΣ"
End Property

Public Property Let ListIndex(value As Integer)
    On Error GoTo GenError
    Control.ListIndex = CInt(value)
    Exit Property
GenError:
    MsgBox "Πρόβλημα στην ανάθεση ListIndex στο Combo Box: " & name, vbCritical, "ΛΑΘΟΣ"
End Property

Public Property Get Enabled() As Boolean
    Enabled = Control.Enabled
End Property

Public Property Let Enabled(value As Boolean)
    If value Then owner.EnableTabForControl Me Else owner.DisableTabForControl Me
'    Control.Enabled = value
    Control.Locked = Not value
    Control.TabStop = value
    Control.BackColor = IIf(value, &H80000005, &H80000004)
End Property

Public Sub FinalizeEdit()
    'do nothing
End Sub

Public Property Get Text() As String
    Text = Control.list(Control.ListIndex)
End Property

Public Sub ReadFromStruct(ByRef inPart As BufferPart, Optional KeyPart As String, Optional VisiblePart As String)
Dim i As Long, k As Long
    
    Control.Clear
    For i = 1 To inPart.Times
        If Not IsMissing(KeyPart) And KeyPart <> "" Then
            If inPart.ByName(KeyPart, i).value = 0 Then Exit For
        End If
        
        If Not IsMissing(VisiblePart) And VisiblePart <> "" Then
            For k = 1 To inPart.SubStruct.PartNum
                If UCase(VisiblePart) = UCase(inPart.ByIndex(k, i).name) Then
                    Control.AddItem Trim(CStr(inPart.ByIndex(k, i).value))
                    Exit For
                End If
            Next k
        Else
            Control.AddItem CStr(inPart.ByIndex(1, i).value)
        End If
    Next i
End Sub

Public Function LocateText(invalue As String) As Long
Dim i As Long
    LocateText = -1
    For i = 0 To Control.ListCount
        If UCase(Control.list(i)) = UCase(invalue) Then
            LocateText = i: Exit Function
        End If
    Next i
End Function

Public Function WriteEJournal(inPhase As Integer, inTrnCode As String) As Boolean
'    eJournalWrite (Prompt & ": " & ClearText)
Dim res As Boolean, astr As String
    If IsOptional(inPhase) And Text = "" Then
        WriteEJournal = True
    Else
        If Prompt.Caption <> "" Then astr = Prompt.Caption & ": "
        WriteEJournal = eJournalWriteFld(owner, CmbNo, astr, Text) ', inTrnCode, CInt(cTRNNum))
    End If
End Function

Public Function IsJournalBeforeOut(inPhase) As Boolean
    IsJournalBeforeOut = JournalBeforeOut(CInt(inPhase))
End Function

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", cmbModel)
    EditFlag(inPhase) = NodeBooleanFld(inNode, "ScrEntry", cmbModel)
    OptionalFlag(inPhase) = NodeBooleanFld(inNode, "ScrOptional", cmbModel)
    OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", cmbModel)
    OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", cmbModel)
    OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", cmbModel)
    JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", cmbModel)
'    JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", fldModel)

    If inPhase = 1 Then
        Set owner = inOwner
        CmbNo = NodeIntegerFld(inNode, "CmbNo", cmbModel)
        CmbName = "CMB" & StrPad_(CStr(CmbNo), 3, "0", "L")
        CMBName2 = UCase(NodeStringFld(inNode, "NAME", cmbModel))
        If InStr(ReservedControlPrefixes, "," & Left(CMBName2, 3) & ",") > 0 Then CMBName2 = ""
        
        name = IIf(CMBName2 <> "", CMBName2, CmbName)

        Set ValidationControl = inProcessControl
        ValidationControl.AddObject CmbName, Me, True
        If Trim(CMBName2) <> "" And UCase(Trim(CMBName2)) <> UCase(Trim(CmbName)) Then
            On Error GoTo FldRegistrationError
            ValidationControl.ExecuteStatement "Set " & CMBName2 & "=" & CmbName
            GoTo FldRegistrationOk
FldRegistrationError:
            MsgBox "Λάθος κατα τη δήλωση του πεδίου: " & CmbName & ":" & CMBName2
FldRegistrationOk:
        End If
        
        LabelName = "CMBLabel" & StrPad_(CStr(CmbNo), 3, "0", "L")
'        NodeStringFld(inNode, "Name", fldModel) & _
'            "_" & NodeStringFld(inNode, "Phase", fldModel)
        Set Prompt = parent.Controls.add("Vb.Label", LabelName)
        Prompt.BackColor = parent.BackColor
        Prompt.AutoSize = False
        Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", cmbModel)
        
        TTabIndex = NodeIntegerFld(inNode, "TabIndex", cmbModel)
        
        parent.ScaleMode = vbCharacters
        Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", cmbModel)
        Prompt.width = NodeIntegerFld(inNode, "ScrPromptWidth", cmbModel)
        parent.ScaleMode = vbTwips
        Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", cmbModel) * 290
        Prompt.height = NodeIntegerFld(inNode, "ScrPromptHeight", cmbModel) * 285
    
        
        ScrLeft = NodeIntegerFld(inNode, "ScrX", cmbModel)
        ScrWidth = NodeIntegerFld(inNode, "ScrWidth", cmbModel)
        ScrTop = NodeIntegerFld(inNode, "ScrY", cmbModel) * 290
        ScrHeight = NodeIntegerFld(inNode, "ScrHeight", cmbModel) * 285
        
'        Control.Locked = True

        On Error GoTo NoSelections
        If Not (inNode.selectSingleNode("SELECTIONS") Is Nothing) Then
            Dim selNode As MSXML2.IXMLDOMElement, anode As MSXML2.IXMLDOMElement
            On Error GoTo 0
            Set selNode = inNode.selectSingleNode("SELECTIONS")
'            If Not (selNode.children Is Nothing) Then
                ChoiceCount_ = selNode.childNodes.length
                ChoiceSuperSetCount = selNode.childNodes.length
                ReDim Choices(selNode.childNodes.length)
                ReDim ChoicesSuperSet(selNode.childNodes.length)
                For i = 0 To selNode.childNodes.length - 1
                    Set anode = selNode.childNodes.item(i)
                    Choices(i).LineCD = Right(anode.tagName, Len(anode.tagName) - 2)
                    Choices(i).LineText = anode.Text
                    Control.AddItem anode.Text
                    ChoicesSuperSet(i).LineCD = Right(anode.tagName, Len(anode.tagName) - 2)
                    ChoicesSuperSet(i).LineText = anode.Text
                Next i
'            End If
        End If
        
        
NoSelections:

    End If
End Sub

Public Sub CopyList(inControl As GenCombo)
    ChoiceCount_ = inControl.ChoiceCount
    ChoiceSuperSetCount = inControl.ChoiceCount
    ReDim Choices(inControl.ChoiceCount)
    ReDim ChoicesSuperSet(inControl.ChoiceCount)
    Control.Clear
    Dim i As Long
    For i = 0 To inControl.ChoiceCount - 1
        Choices(i).LineCD = inControl.ChoiceLineCD(i)
        Choices(i).LineText = inControl.ChoiceLineText(i)
        ChoicesSuperSet(i).LineCD = inControl.ChoiceLineCD(i)
        ChoicesSuperSet(i).LineText = inControl.ChoiceLineText(i)
        Control.AddItem inControl.ChoiceLineText(i)
    Next i
    
End Sub

Public Sub AddChoice(ChoiceValue, TextValue)
    If ChoiceCount = 0 Then
        ReDim Choices(ChoiceCount_)
    Else
        ReDim Preserve Choices(ChoiceCount_)
    End If
    ChoiceCount_ = ChoiceCount_ + 1
    Choices(ChoiceCount_ - 1).LineCD = ChoiceValue
    Choices(ChoiceCount_ - 1).LineText = TextValue
    Control.AddItem TextValue
End Sub

Property Get Choice() As String
    If Control.ListIndex = -1 Then Choice = "": Exit Property
    If Control.ListIndex < ChoiceCount_ And ChoiceCount_ > 0 Then
        Choice = Choices(Control.ListIndex).LineCD
    End If
End Property

Property Let Choice(invalue As String)
    Dim i As Long
    If invalue = "" Then Control.ListIndex = -1: Exit Property
    For i = 0 To ChoiceCount_ - 1
        If UCase(Choices(i).LineCD) = UCase(invalue) Then
            Control.ListIndex = i
        End If
    Next i
End Property

Public Property Get ChoiceCount() As Integer
    ChoiceCount = ChoiceCount_
End Property

Public Property Get ChoiceLineCD(Index As Long) As String
    ChoiceLineCD = Choices(Index).LineCD
End Property

Public Property Get ChoiceLineText(Index As Long) As String
    ChoiceLineText = Choices(Index).LineText
End Property

Public Sub Clear()
    Control.Clear
    ChoiceCount_ = 0
    ChoiceSuperSetCount = 0
End Sub

Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim xml As DOMDocument30
Set xml = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = xml.createElement("COMBOBOX")
    Set attr = xml.createAttribute("NO")
    attr.nodeValue = UCase(Me.CmbNo)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("NAME")
    attr.nodeValue = UCase(Me.CmbName)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("PROMPT")
    attr.nodeValue = UCase(Me.Prompt)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("CHOICE")
    attr.nodeValue = UCase(Me.Choice)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("TEXT")
    attr.nodeValue = UCase(Me.Text)
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("READONLY")
    attr.nodeValue = UCase(Not Me.IsEditable(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
    Set attr = xml.createAttribute("CHOICECOUNT")
    attr.nodeValue = UCase(Me.ChoiceCount)  '(Owner.ProcessPhase))
    elm.setAttributeNode attr
                    
    Dim rowelm As IXMLDOMElement
    Dim i As Long
    For i = 0 To Me.ChoiceCount - 1
        Set rowelm = xml.createElement("CHOICE")
        elm.appendChild rowelm
        Set attr = xml.createAttribute("VALUE")
        attr.nodeValue = Me.ChoiceLineCD(i)
        rowelm.setAttributeNode attr
        Set attr = xml.createAttribute("TEXT")
        attr.nodeValue = Me.ChoiceLineText(i)
        rowelm.setAttributeNode attr
    Next
    Set TranslateToProperties = elm
End Function
Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    Dim aattr As IXMLDOMAttribute
    If inPhase = 0 Then inPhase = 1
    If elm.SelectNodes("./CHOICE").length > 0 Then
        Dim aval As String
        aval = Choice
        Clear
        Dim aChoice As IXMLDOMElement
        For Each aChoice In elm.SelectNodes("./CHOICE")
            AddChoice aChoice.getAttribute("VALUE"), aChoice.getAttribute("TEXT")
        Next aChoice
        Choice = aval
    End If
    For Each aattr In elm.Attributes
        Select Case aattr.baseName
            Case "CHOICE"
                Me.Choice = aattr.value
            Case "READONLY"
                If aattr.value = "FALSE" Then
                    Me.SetEditableNoRefresh inPhase, True
                ElseIf aattr.value = "TRUE" Then
                    Me.SetEditableNoRefresh inPhase, False
                End If
            Case "PROMPT"
                Me.Prompt.Caption = aattr.value
            Case "VISIBLE"
                If aattr.value = "FALSE" Then
                    Me.SetDisplay inPhase, False
                ElseIf aattr.value = "TRUE" Then
                    Me.SetDisplay inPhase, True
                End If
        End Select
    Next aattr
End Sub

