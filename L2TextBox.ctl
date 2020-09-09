VERSION 5.00
Begin VB.UserControl L2TextBox 
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1590
   ScaleWidth      =   3765
   Begin VB.TextBox Control 
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "L2TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form
Public name As String
Public Mandatory As Boolean

Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long

Public PasswordChar As String

Private ClearText As String, OLDTEXT As String

Private EditLength As Integer, DisplayMask As String
Private FieldType As String
Private EditType As String
Private OutMask As String
Private Signed As String
'EditType: String, Number, Date
Public Align As String
'Align: Left, Right
Private Validation As String, ValidationCD As Long
'Validation example: Account2CD, Account1CD
Private cLocked As Boolean

Public ValidOk As Boolean, ValidationError As String
Public ChangeFocusOk As Boolean, ChangeFocusError As String

Private ValidationFailed As Boolean, ValidationErrMessage As String

Public TTabIndex As Integer, tTabStop As Boolean
Private OnF1 As String, OnEnterKey As String
Public Caption As String, Label As String

Private disableChangeControl As Boolean

Public Sub Activate()
On Error Resume Next
    Control.SetFocus
End Sub

Public Property Get Text() As String
    Text = ClearText
End Property

Public Property Let Text(value As String)
    ClearText = value
    If owner.ActiveControl Is Nothing Then
        Control.Text = FormatedText
    Else
        If owner.ActiveControl.name <> name Then
            Control.Text = FormatedText
        Else
            Control.Text = ClearText
        End If
    End If
End Property

Public Property Get FormatedText() As String
Dim astr  As String
Dim aPos As Integer, bpos As Integer
On Error GoTo ErrorPos:
    If DisplayMask <> "" Then
        If ClearText = "" Or Trim(Replace(ClearText, Chr(160), " ")) = "" Then Exit Property
        Select Case UCase(FieldType)
        Case UCase("Account2CD")
            astr = ClearText
            If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        Case UCase("Account1CD")
            astr = ClearText
            If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
            If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = astr & CalcCd2_(Left(astr, 10))
        Case UCase("Account0CD")
            astr = ClearText
            If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
            If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 9, "0", "L")
            astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
            astr = astr & CalcCd2_(Left(astr, 10))
        Case Else
            If ClearText <> "" Then
                aPos = InStr(DisplayMask, ".")
                If aPos > 0 Then
                    bpos = Len(DisplayMask) - aPos
                    
'                    If Right(ClearText, 1) = "-" Then bpos = bpos - 1
                    astr = CStr(CDbl(ClearText) / 10 ^ bpos)
                Else
                    astr = ClearText
                End If
            Else
                astr = ClearText
            End If
        End Select
        
        If ClearText <> "" Then
            FormatedText = format(astr, DisplayMask)
        Else
            FormatedText = ""
        End If
    Else: FormatedText = ClearText
    End If
    Exit Property
ErrorPos:
    Call NBG_LOG_MsgBox("Πεδίο :" & name & vbCrLf & " Λάθος:" & Err.number & Err.description & " ClearText = " & ClearText, True)
End Property

Property Get OutText() As String
    Dim astr As String
    Select Case UCase(FieldType)
    Case UCase("Account2CD")
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        OutText = astr
    Case UCase("Account1CD")
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
            If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        OutText = astr
    Case UCase("ACCOUNTIRISCD")
        astr = ClearText
        If Len(astr) <= 10 Then
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = Mid(astr, 1, 9) & CalcCd1_(Left(astr, 9), 9)
        Else
            astr = StrPad_(astr, 11, "0", "L")
            If Len(astr) < 11 Then astr = StrPad_(astr, 11, "0", "L")
            astr = Mid(astr, 1, 10) & CalcCd2_(Left(astr, 10))
        End If
    Case UCase("Account0CD")
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
            If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 9, "0", "L")
            astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        OutText = astr
    Case UCase("GermanyAccount")
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 8 Then astr = cBRANCH & StrPad_(astr, 7, "0", "L") & "0"
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        OutText = astr
    Case UCase("date08")
        astr = ClearText
        If astr <> "" Then
            astr = Right("00000000" & astr, 8)
            astr = Right(astr, 4) & "-" & Mid(astr, 3, 2) & "-" & Left(astr, 2)
        End If
        OutText = astr
    Case UCase("time06")
        astr = ClearText
        If astr <> "" Then
            astr = Right("000000" & astr, 6)
            astr = Left(astr, 2) & ":" & Mid(astr, 3, 2) & ":" & Right(astr, 2)
        End If
        OutText = astr
    Case UCase("time04")
        astr = ClearText
        If astr <> "" Then
            astr = Right("0000" & astr, 4)
            astr = Left(astr, 2) & ":" & Right(astr, 2) & ":00"
        End If
        OutText = astr
    Case Else
        
        If OutMask = "" Then
            OutText = ClearText
        Else
            If EditType = "number" Then
                If ClearText = "" Then
                    OutText = format(0, OutMask): Exit Property
                End If
            ElseIf EditType = "none" Or EditType = "string" Then
                If Trim(ClearText) = "" Then
                    OutText = "": Exit Property
                End If
            End If
            OutText = format(ClearText, OutMask)
        End If
    End Select
    Exit Property
    
End Property

Private Sub Control_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyF1 Then
        If OnF1 <> "" Then
        
            owner.Enabled = False
            owner.owner.DocumentManager.XmlObjectList.item(OnF1).XML
            owner.Enabled = True
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If OnEnterKey <> "" Then
            owner.Enabled = False
            owner.owner.DocumentManager.XmlObjectList.item(OnEnterKey).XML
            owner.Enabled = True
        End If
        
    End If
End Sub

'Private Function OtherActive() As Control
'Dim acontrol As Control
'    On Error Resume Next
'    For Each acontrol In owner.Controls
'        If acontrol.Enabled And acontrol.TabStop And acontrol.Visible And acontrol.name <> name Then
'            Set OtherActive = acontrol
'            Exit Function
'        End If
'
'    Next acontrol
'    Set OtherActive = Nothing
'    Exit Function
'End Function
'
Private Sub Control_KeyPress(KeyAscii As Integer)
Dim aPos As Integer
'Dim acontrol As Control
    If KeyAscii = 13 Then
        'Set acontrol = OtherActive
        'If acontrol Is Nothing Then Exit Sub
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = 9 Then
        If Enabled Then
            'Set acontrol = OtherActive
            'If acontrol Is Nothing Then Exit Sub
            SendKeys "{TAB}"
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub FinalizeEdit()
    If Trim(FormatedText) <> Trim(Control.Text) Then
        ClearText = Trim(Control.Text)
    End If
'Συμπληρώνει τα δεκαδικά στα αριθμητικά πεδία άν ο χρήστης έχει πατήσει το ,
    If EditType = "number" Then
        Dim fstr As String, astr As String, pos1 As Integer, pos2 As Integer, i As Integer
        If Len(ClearText) > 1 Then
            If Right(ClearText, 1) = "-" Or Right(ClearText, 1) = "+" Then
                ClearText = Right(ClearText, 1) & Left(ClearText, Len(ClearText) - 1)
            End If
        End If
        fstr = Trim(DisplayMask)
        If ClearText <> "" And fstr <> "" Then
            pos1 = InStr(fstr, "."):  pos2 = InStr(ClearText, ".")
            If pos2 = 0 Then pos2 = InStr(ClearText, ",")
            If pos1 > 0 And pos2 > 0 Then
                
                If Mid(ClearText, pos2, 1) = "," Then ClearText = Replace(ClearText, ",", ".")
                
                pos1 = Len(fstr) - pos1: pos2 = Len(ClearText) - pos2
                If pos1 > pos2 Then ClearText = Trim(ClearText) & String(pos1 - pos2, "0") _
                Else If pos1 < pos2 Then ClearText = Left(ClearText, Len(ClearText) - pos2 + pos1)
                astr = ""
                For i = 1 To Len(ClearText)
                    If Mid(ClearText, i, 1) <> "." Then astr = astr + Mid(ClearText, i, 1)
                Next i
                ClearText = astr
            End If
        End If
    End If
    
End Sub

Private Sub Control_LostFocus()
    disableChangeControl = True
    FinalizeEdit
    Control.Text = FormatedText
    disableChangeControl = False
    
    Dim Message As String
    If Not Enabled Or Control.Locked Then Exit Sub
    If Not ChkValid(ClearText) Then Control.SetFocus: Control.BackColor = &HC0FFFF: Exit Sub
    If UCase(Validation) <> "" And UCase(Validation) <> "NONE" Then
        If Not L2ChkFldType(Message, ClearText, Validation) Then
            owner.WriteStatusMessage Message
            Control.SetFocus: Beep: Control.BackColor = &HC0FFFF: Exit Sub
        End If
    End If
    owner.WriteStatusMessage "": Control.BackColor = &HFFFFFF:
End Sub

Private Function ChkValid(ByVal value As String) As Boolean
    On Error GoTo ValidationFailed
    If UCase(EditType) = UCase("Number") Then
        Dim anumber As Double
        If value = "" Then value = "0"
        anumber = CDbl(value)
    End If
    
    GoTo validationok
ValidationFailed:
    ChkValid = False: Exit Function
validationok:
    ChkValid = True: Exit Function
End Function

Private Sub RestoreText()
    Control.Text = OLDTEXT
    Control.SelStart = Len(Control.Text)
    Control.SelLength = 0
    Beep
End Sub

Private Sub Control_Change()
If disableChangeControl Then Exit Sub
Dim astr As String
Dim aNum As Double
Dim aselpos As Integer, asellength As Integer
    If EditLength > 0 Then
        If Right(Control.Text, 1) = vbTab Then
            RestoreText
            Exit Sub
        End If
        If Len(Control.Text) > EditLength Then
            RestoreText
            Exit Sub
        End If
    End If
    If Len(Control.Text) > 0 Then
        If UCase(EditType) = UCase("Number") Or UCase(EditType) = UCase("date") Then
            astr = Control.Text
            On Error GoTo ErrorPos
            If UCase(EditType) = UCase("Number") And UCase(Signed) = UCase("true") Then
                If Not (astr = "+" Or astr = "-") Then
                    aNum = CDbl(astr)
                End If
            ElseIf UCase(EditType) = UCase("Number") And UCase(Signed) = UCase("false") Then
                If InStr(astr, "+") > 0 Then GoTo ErrorPos
                If InStr(astr, "-") > 0 Then GoTo ErrorPos
                aNum = CDbl(astr)
            Else
                aNum = CDbl(astr)
                If Mid(astr, Len(astr), 1) = "+" Then GoTo ErrorPos
                If Mid(astr, Len(astr), 1) = "-" Then GoTo ErrorPos
            
            End If
            'If InStr(DisplayMask, ".") <= 0 And Mid(astr, Len(astr), 1) = "." Then GoTo ErrorPos
            'If InStr(DisplayMask, ",") <= 0 And Mid(astr, Len(astr), 1) = "," Then GoTo ErrorPos
            GoTo noErrorPos
ErrorPos:
            'If astr = "-" Or astr = "+" Or astr = "." Then GoTo noErrorPos
            If astr = "." Then GoTo noErrorPos
            RestoreText
            Exit Sub
noErrorPos:
        End If
        
    End If
    aselpos = Control.SelStart
    asellength = Control.SelLength
    Control.Text = UCase(Control.Text)
    Control.SelStart = aselpos
    Control.SelLength = asellength
        
    ClearText = Control.Text
        
    OLDTEXT = Control.Text
End Sub

Private Sub Control_GotFocus()
    Control.Text = ClearText
    Control.SelStart = 0
    Control.SelLength = Len(ClearText)
End Sub

Public Sub setBackColor()
    If Enabled And Not Control.Locked Then
        Control.BackColor = &HFFFFFF
    Else
        'Control.BackColor = &H8000000F
        Control.BackColor = &HD0D0D0
    End If
End Sub

Private Sub Control_Validate(Cancel As Boolean)
     Dim Message As String
    Dim padClearText As String
    If Not Enabled Or Control.Locked Then Exit Sub
    If Not ChkValid(ClearText) Then Cancel = True: Beep: Control.BackColor = &HC0FFFF: Exit Sub
    If UCase(Validation) <> "" And UCase(Validation) <> "NONE" Then
        If (EditLength > 0 And ClearText <> "" And UCase(Validation) <> "ACCOUNT2CD") Then
            If UCase(Validation) = "ACCOUNT2CD" Then
               padClearText = ClearText
            ElseIf UCase(Validation) = "ACCOUNT1CD" Then
               padClearText = ClearText
            ElseIf UCase(Validation) = "ACCOUNTIRISCD" Then
               padClearText = ClearText
            Else
                padClearText = StrPad_(ClearText, EditLength, "0")
            End If
        Else
            padClearText = ClearText
        End If
        If Not L2ChkFldType(Message, padClearText, Validation) Then
            owner.WriteStatusMessage Message
            Cancel = True: Beep: Control.BackColor = &HC0FFFF: Exit Sub
        End If
    End If
    owner.WriteStatusMessage "": Control.BackColor = &HFFFFFF:
End Sub

Private Sub UserControl_Resize()
    Control.Left = 0
    Control.Top = 0
    Control.width = width
    Control.height = height
End Sub

Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)
    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.value
    
    LoadFromIXMLDOMElement inNode
    
'        PasswordChar = NodeStringFld(inNode, "ScrPasswordChar", fldModel)
'        If PasswordChar <> "" Then
'            VControl.Font.Name = "BBSecret"
'        End If

End Sub

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute, fieldTypeNode As IXMLDOMElement
    disableChangeControl = True
    Set aattr = elm.getAttributeNode("fieldtype")
    If Not (aattr Is Nothing) Then
        FieldType = UCase(aattr.Text)
        Set fieldTypeNode = L2ModelFile.documentElement.selectSingleNode("//datatype[@name='" & FieldType & "']")
        If fieldTypeNode Is Nothing Then
        Else
            For Each aattr In fieldTypeNode.Attributes
                Select Case UCase(aattr.baseName)
                    Case "ALIGN"
                        Align = aattr.value
                    Case "DISPLAYMASK"
                        DisplayMask = aattr.value
                    Case "EDITLENGTH"
                        EditLength = aattr.value
                    Case "EDITTYPE"
                        EditType = aattr.value
                        If UCase(EditType) = UCase("number") Then
                            If Align = "" Then Align = "right"
                        Else
                            If Align = "" Then Align = "left"
                        End If
                    Case "VALIDATION"
                        Validation = aattr.value
                    Case "VALIDATIONCD"
                        ValidationCD = aattr.value
                    Case "OUTMASK"
                        OutMask = aattr.value
                    Case "SIGNED"
                       Signed = aattr.value
                End Select
            Next aattr
        End If
    End If
    For Each aattr In elm.Attributes
        Select Case UCase(aattr.baseName)
            Case "LEFT"
                tLeft = aattr.value
            Case "TOP"
                tTop = aattr.value
            Case "WIDTH"
                tWidth = aattr.value
            Case "HEIGHT"
                tHeight = aattr.value
            Case "TABSTOP"
                If UCase(aattr.value) = bvTrue Then
                    tTabStop = True 'And Control.Enabled
                ElseIf UCase(aattr.value) = bvFalse Then
                    tTabStop = False
                End If
            Case "TABINDEX"
                TTabIndex = aattr.value
            Case "VISIBLE"
                If UCase(aattr.value) = bvFalse Then
                    Me.tVisible = False
                ElseIf UCase(aattr.value) = bvTrue Then
                    Me.tVisible = True
                End If
            
            Case "VALIDATION"
                If (UCase(aattr.value) = "NONE" And fieldTypeNode Is Nothing) Or (UCase(aattr.value) <> "NONE") Then
                    Validation = aattr.value
                End If
            Case "VALIDATIONCD"
                ValidationCD = aattr.value
            Case "DISPLAYMASK"
                If (aattr.value = "" And fieldTypeNode Is Nothing) Or (aattr.value <> "") Then
                    DisplayMask = aattr.value
                End If
            Case "EDITLENGTH"
                If (aattr.value = "0" And fieldTypeNode Is Nothing) Or (aattr.value <> "0") Then
                    EditLength = aattr.value
                End If
            Case "OUTMASK"
                OutMask = aattr.value
            Case "EDITTYPE"
                If (UCase(aattr.value) = "NONE" And fieldTypeNode Is Nothing) Or (UCase(aattr.value) <> "NONE") Then
                    EditType = aattr.value
                    If UCase(EditType) = UCase("number") Then
                        If Align = "" Then Align = "right"
                    Else
                        If Align = "" Then Align = "left"
                    End If
                End If
            Case "ALIGN"
                If (aattr.value = "" And fieldTypeNode Is Nothing) Or (aattr.value <> "") Then
                    If fieldTypeNode Is Nothing Then Align = aattr.value
                End If
                
            Case "TEXT"
                Me.Text = aattr.value
            Case "ENABLED"
                If UCase(aattr.value) = bvTrue Then
                    Enabled = True
                    Control.BackColor = &HFFFFFF
                ElseIf UCase(aattr.value) = bvFalse Then
                    Enabled = False
                    Control.BackColor = &HD0D0D0
                End If
                setBackColor
            Case "READONLY"
                If Trim(aattr.value) = bvTrue And Not Enabled Then
                    Control.Locked = True
                    Enabled = True
                    cLocked = True
                ElseIf Trim(aattr.value) = bvFalse Then
                    Control.Locked = False
                    cLocked = False
                End If
            Case "OPTIONAL"
                If UCase(aattr.value) = bvFalse Then
                    Me.Mandatory = True
                ElseIf UCase(aattr.value) = bvTrue Then
                    Me.Mandatory = False
                End If
            Case "ONF1"
                OnF1 = aattr.value
            Case "ONENTERKEY"
                OnEnterKey = aattr.value
            Case "CAPTION"
                Caption = aattr.value
            Case "LABEL"
                Label = aattr.value
            Case "PASSWORDCHAR"
                If UCase(aattr.value) = bvTrue Then
                    Control.Font.name = "BBSecret"
                End If
        End Select
    Next aattr
    If UCase(Align) = UCase("left") Or Align = "" Then
        Control.Alignment = 0
    ElseIf UCase(Align) = UCase("right") Then
        Control.Alignment = 1
    End If
    disableChangeControl = False
End Sub


Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("textbox")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = Me.tVisible
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("text")
    attr.nodeValue = Me.Text
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("formatedtext")
    attr.nodeValue = Me.FormatedText
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("outtext")
    attr.nodeValue = Me.OutText
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("enabled")
    attr.nodeValue = Enabled
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("optional")
    attr.nodeValue = Not Me.Mandatory
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Me.Caption
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("label")
    attr.nodeValue = Me.Label
    elm.setAttributeNode attr
    
    Set attr = XML.createAttribute("readonly")
    attr.nodeValue = cLocked
    elm.setAttributeNode attr
                    
    Dim Message As String
    If ClearText = "" And Mandatory And Enabled And Not Control.Locked Then
        Set attr = XML.createAttribute("validationerror")
        attr.nodeValue = "Υποχρεωτικό πεδίο"
        elm.setAttributeNode attr
        'If Enabled Then
        Control.BackColor = &HC0FFFF:
        'End If
    Else
        If Not L2ChkFldType(Message, ClearText, Validation) Then
            Set attr = XML.createAttribute("validationerror")
            attr.nodeValue = Message
            elm.setAttributeNode attr
            
            'If Enabled Then
            '    Control.BackColor = &HC0FFFF:
            'End If
        ElseIf Enabled And Not Control.Locked Then
            'Control.BackColor = &HFFFFFF:
        End If
        setBackColor
    End If
    
    
    Set IXMLDOMElementView = elm
End Function

Public Sub CleanUp()
    Set owner = Nothing
End Sub
