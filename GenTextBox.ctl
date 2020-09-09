VERSION 5.00
Begin VB.UserControl GenTextBox 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2220
   ScaleWidth      =   4605
   Begin VB.TextBox VControl 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox HiddenFld 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "GenTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Const etNONE = 0
'Const etTEXT = 1
'Const etNUMBER = 2

Private owner As Form
Public FldNo As Integer, FldName As String, FldName2 As String, LabelName As String, name As String
Public TotalName As String, TotalPos As Integer
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

Public ValidOk As Boolean, ValidationError As String
Public ChangeFocusOk As Boolean, ChangeFocusError As String

Public FormatBeforeOutFlag As Boolean, FormatAfterInFlag As Boolean
Public DisplayMask As String, Editmask As String, OutMask As String, DocMask As String
Public EditLength As Integer, EditType As Integer
Public ValidationCode As Integer

Public ValidationControl As ScriptControl
Private ValidationFlag As Boolean
Private ClearText As String, OLDTEXT As String, OutBuffText As String, InBuffText As String, EnableEditChk As Boolean
Private ScrHelp As String
Private ValidationFailed As Boolean, ValidationErrMessage As String

Public HPSOutStruct As String, HPSOutPart As String, HPSInStruct As String, HPSInPart As String
Public TTabIndex As Integer, Tabbed As Boolean

Public HelpFormWidth As Long
Public HelpFormHeight As Long

Private Type HelpLine
    LineCD As String
    LineText As String
End Type

Private Choices() As HelpLine
Private ChoiceCount As Integer

Private ChoicesSuperSet() As HelpLine
Private ChoiceSuperSetCount As Integer


Private InvertTab As Boolean
Private ClearLastKey As Boolean

'------------------------------------------------
'Record βοήθειας για αρχείο 3ων
Private Type ThrdRecord
    CD As String
    account As String
    name As String
    fld(1 To 11) As String
End Type
Private ThrdSelection As ThrdRecord

Private Type StkRecord
    CD As String
    name As String
End Type
Private StkSelection As StkRecord
'------------------------------------------------
'------------------------------------------------

Public Property Get ThrdCD() As String
    ThrdCD = ThrdSelection.CD
End Property

Public Property Get ThrdAccount() As String
    ThrdAccount = ThrdSelection.account
End Property

Public Property Get ThrdName() As String
    ThrdName = ThrdSelection.name
End Property
Public Property Get ThrdFld(inFldNo) As String
    ThrdFld = ThrdSelection.fld(inFldNo)
End Property

Public Property Get StkCD() As String
    StkCD = StkSelection.CD
End Property

Public Sub SetAsActive()
    If VControl.Visible Then VControl.SetFocus
End Sub

Public Sub SetDisplay(inPhase, SetFlag)
    DisplayFlag(CInt(inPhase)) = CBool(SetFlag)
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

Public Function IsOptional(inPhase) As Boolean
    IsOptional = OptionalFlag(CInt(inPhase))
End Function

Public Function IsJournalBeforeOut(inPhase) As Boolean
    IsJournalBeforeOut = JournalBeforeOut(CInt(inPhase))
End Function

Public Function IsJournalAfterIN(inPhase) As Boolean
    IsJournalAfterIN = JournalAfterIn(CInt(inPhase))
End Function

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

Public Sub SetOutBuffLength(inPhase, inLength)
    OutBuffLength(CInt(inPhase)) = CInt(inLength)
End Sub

Public Sub SetInBuffLength(inPhase, inLength)
    InBuffLength(CInt(inPhase)) = CInt(inLength)
End Sub

Public Function GetInBuffPos(inPhase) As Integer
    GetInBuffPos = InBuffPos(CInt(inPhase))
End Function

Public Function GetInBuffLength(inPhase) As Integer
    GetInBuffLength = InBuffLength(CInt(inPhase))
End Function

Public Function GetHelpFromValue(invalue) As String
Dim i As Integer
    GetHelpFromValue = ""
    If IsNumeric(invalue) Then
        For i = 0 To ChoiceCount - 1
            If IsNumeric(Choices(i).LineCD) Then
                If CLng(Choices(i).LineCD) = CLng(invalue) Then GetHelpFromValue = Choices(i).LineText: Exit Function
            Else
                If Choices(i).LineCD = CStr(invalue) Then GetHelpFromValue = Choices(i).LineText: Exit Function
            End If
        Next i
    Else
        For i = 0 To ChoiceCount - 1
            If Choices(i).LineCD = CStr(invalue) Then GetHelpFromValue = Choices(i).LineText: Exit Function
        Next i
    End If
End Function

Public Function GetDBL() As Double
    On Error GoTo ExitError
    GetDBL = CDbl(ClearText)
    Exit Function
ExitError:
    GetDBL = 0
End Function

Public Function WriteEJournal(inPhase As Integer, inTrnCode As String) As Boolean
'    eJournalWrite (Prompt & ": " & ClearText)
Dim res As Boolean, astr As String
    If IsOptional(inPhase) And Text = "" Then
        WriteEJournal = True
    Else
        If Prompt.Caption <> "" Then astr = Prompt.Caption & ": "
        WriteEJournal = eJournalWriteFld(owner, FldNo, astr, FormatedText) ', inTrnCode, CInt(cTRNNum))
    End If
End Function

Public Function ChkValidL1(inPhase As Integer) As Boolean
'01: ΧΩΡΙΣ VALIDATION
'02: Λογαριασμός με CD,
'03: Λογαριασμός χωρίς CD,
'04: Δάνειο με CD,
'05: Δάνειο χωρίς CD,
'06: Ημερομηνία
'07: Αριθμός Εγγραφής
'08: Ειδικός με CD
'09: Γενικός Λογαριασμός Δανείου
'10: Λογαριασμός Καταθέσεων με 1 CD
'11: Τραπεζική Επιταγή
'12: ΕΘΝΟΚΑΡΤΑ
'13: Τραπεζική Εντολή
'14: ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
'21: Λογαριασμος Καταθέσεων Γερμανία

Dim astr As String, ares As Integer, aFlag As Boolean
On Error GoTo chkFailed
    astr = ClearText
    ChkValidL1 = True
    If Trim(astr) = "" Then Exit Function
    
    Select Case ValidationCode
    Case 2
        If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
        If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 11, "0", "L")
        
        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
        End If
        
        ares = CalcCd2_(Left(astr, 10))
        If CInt(Mid(astr, 11, 1)) <> ares Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
    Case 3
    
    Case 10
        If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
        If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 10, "0", "L")
        
        If cKMODEValue <> "IRIS" Then
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
        End If
        
        ares = CalcCd2_(Left(astr, 10))
            
    Case 11 'Τραπεζική Επιταγή
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
        
    Case 13 'Τραπεζική Εντολή
        astr = Trim(astr)
        aFlag = False
        On Error GoTo chkFailed
        If astr = "" Or astr = "0" Then
            aFlag = True
        Else
            If CLng(astr) >= 550000000 And CLng(astr) <= 600000000 And cVersion >= 20010101 Then
                aFlag = ChkGenBankCheque_(astr)
            Else
                ''FBB 17/06/2013
                'aFlag = ChkGenBankCheque_(astr)
                If Len(astr) > 1 Then
                    ares = CLng(Left(astr, Len(astr) - 1)) Mod 11
                    If ares = 10 Then ares = 0
                    If CInt(Right(astr, 1)) = ares Then aFlag = True
                Else
                    aFlag = True
                End If
            End If
        End If
        
        If Not aFlag Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
        
    Case 14 'ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
    
    Case 4
        If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
        ares = CalcCd1_(Left(astr, 9), 9)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
    Case 5
    Case 6
        astr = GetFormatedText
        If Not IsDate(astr) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Μορφή Ημερομηνίας": GoTo chkFailed
        End If
        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Μορφή Ημερομηνίας": GoTo chkFailed
        End If
        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Μορφή Ημερομηνίας": GoTo chkFailed
        End If
    Case 7
        If Len(astr) < 13 Then astr = StrPad_(astr, 13, "0", "L")
        ares = CalcCd1_(Left(astr, 12), 12)
        If CInt(Right(astr, 1)) <> ares Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
    Case 8
    Case 9
        astr = Trim(astr)
        If Len(astr) > 3 Then
            ares = CalcSAccCd(Left(astr, Len(astr) - 1))
            If CInt(Right(astr, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
        End If
    Case 12
        astr = Trim(astr)
        If Not ChkCard(astr) Then
            ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
        End If
    Case 21 'Λογαριασμος Καταθέσεων Γερμανία
        If Len(astr) < 8 Then 'astr = StrPad_(astr, 8, "0", "L")
            astr = StrPad_(astr, 7, "0", "L")
            ares = CalcCd1_(Mid(astr, 1, 6), 6)
            If CInt(Mid(astr, 7, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
            
'            DisplayMask = "000/000000-00"
        Else
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
            ares = CalcCd2_(Left(astr, 10))
            If CInt(Mid(astr, 11, 1)) <> ares Then
                ValidationFailed = True: ValidationErrMessage = "Λάθος Ψηφίο Ελέγχου": GoTo chkFailed
            End If
            
'            DisplayMask = "000/000000-00"
        End If
    End Select
    ChkValidL1 = True
'28-01-2003
'    Screen.ActiveForm.sbWriteStatusMessage ""
     If ValidationFailed Then 'ειχε αποτύχει προηγούμενο validation
        ValidationFailed = False: ValidationErrMessage = "": Screen.activeform.sbWriteStatusMessage ""
     End If
'    If ValidationFlag And EditFlag(inPhase) Then
'        ValidOk = True
'        ValidationControl.Run FldName & "_Validation"
'        If Not ValidOk Then
'            If ValidationError <> "" Then
'                Screen.ActiveForm.sbWriteStatusMessage ValidationError
'            Else
'                Screen.ActiveForm.sbWriteStatusMessage "Λάθος Κατα τον Ελεγχο του Πεδίου"
'            End If
'            GoTo chkFailed
'        End If
'    End If
    
    Exit Function
chkFailed:
    Screen.activeform.sbWriteStatusMessage ValidationErrMessage: ChkValidL1 = False
End Function


Public Function ChkValid(inPhase As Integer) As Boolean
'01: ΧΩΡΙΣ VALIDATION
'02: Λογαριασμός με CD,
'03: Λογαριασμός χωρίς CD,
'04: Δάνειο με CD,
'05: Δάνειο χωρίς CD,
'06: Ημερομηνία
'07: Αριθμός Εγγραφής
'08: Ειδικός με CD
'09: Γενικός Λογαριασμός Δανείου
'10: Λογαριασμός Καταθέσεων με 1 CD
'11: Τραπεζική Επιταγή
'12: ΕΘΝΟΚΑΡΤΑ
'13: Τραπεζική Εντολή
'21: Λογαριασμος Καταθέσεων Γερμανία

Dim astr As String, ares As Integer
Dim aFlag As Boolean
On Error GoTo chkFailed
    
    astr = ClearText
    aFlag = ChkValidL1(inPhase)
    If Not aFlag Then
        GoTo chkFailed
    End If
    If aFlag And Not OptionalFlag(inPhase) And EditFlag(inPhase) _
    And Trim(astr) = "" Then
        Screen.activeform.sbWriteStatusMessage "Υποχρεωτικό Πεδίο: " & Prompt
        GoTo chkFailed
    Else
        ChkValid = True
        Exit Function
    End If
    
        
'------------------------------------------------------------------
    
    If Not OptionalFlag(inPhase) And EditFlag(inPhase) _
    And Trim(astr) = "" Then
        Screen.activeform.sbWriteStatusMessage "Υποχρεωτικό Πεδίο"
        GoTo chkFailed
    ElseIf (OptionalFlag(inPhase) Or Not EditFlag(inPhase)) _
    And Trim(astr) = "" Then
        ChkValid = True
        Exit Function
    End If
    
    Select Case ValidationCode
    Case 2
        If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
'        If Len(astr) = 8 Then astr = fnReadConst_("BranchID") & astr
        If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 11, "0", "L")
        ares = CalcCd1_(Mid(astr, 4, 6), 6)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
        ares = CalcCd2_(Left(astr, 10))
        If CInt(Mid(astr, 11, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 3
    
    Case 10 'Λογαριασμός Καταθέσεων με 1 CD
        If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
'        If Len(astr) = 8 Then astr = fnReadConst_("BranchID") & astr
        If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        astr = StrPad_(astr, 10, "0", "L")
        ares = CalcCd1_(Mid(astr, 4, 6), 6)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
        ares = CalcCd2_(Left(astr, 10))
    Case 11 'Τραπεζική Επιταγή
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 13 'Τραπεζική Εντολή
        astr = Trim(astr)
        aFlag = False
        On Error GoTo chkFailed
        If astr = "" Or astr = "0" Then
            aFlag = True
        Else
            If CLng(astr) >= 550000000 And CLng(astr) <= 600000000 And cVersion >= 20010101 Then
                aFlag = ChkGenBankCheque_(astr)
            Else
                ''FBB 17/06/2013
                'aFlag = ChkGenBankCheque_(astr)
                If Len(astr) > 1 Then
                    ares = CLng(Left(astr, Len(astr) - 1)) Mod 11
                    If ares = 10 Then ares = 0
                    If CInt(Right(astr, 1)) = ares Then aFlag = True
                Else
                    aFlag = True
                End If
            End If
        End If
        
        If Not aFlag Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 14 'ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
        astr = Trim(astr)
        If Not ChkETECheque_(CLng(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 4
        If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
        ares = CalcCd1_(Left(astr, 9), 9)
        If CInt(Mid(astr, 10, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 5
    Case 6
        astr = GetFormatedText
        If Not IsDate(astr) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
            GoTo chkFailed
        End If
    Case 7
        If Len(astr) < 13 Then astr = StrPad_(astr, 13, "0", "L")
        ares = CalcCd1_(Left(astr, 12), 12)
        If CInt(Right(astr, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 8
    Case 9
        astr = Trim(astr)
        ares = CalcSAccCd(Left(astr, Len(astr) - 1))
        If CInt(Right(astr, 1)) <> ares Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 12
        astr = Trim(astr)
        If Not ChkCard(astr) Then
            Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
            GoTo chkFailed
        End If
    Case 21 'Λογαριασμος Καταθέσεων Γερμανία
        If Len(astr) < 8 Then 'astr = StrPad_(astr, 8, "0", "L")
            astr = StrPad_(astr, 7, "0", "L")
            ares = CalcCd1_(Mid(astr, 1, 6), 6)
            If CInt(Mid(astr, 7, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
'            DisplayMask = "000000-0"
        Else
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
            ares = CalcCd1_(Mid(astr, 4, 6), 6)
            If CInt(Mid(astr, 10, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Λάθος Ψηφίο Ελέγχου"
                GoTo chkFailed
            End If
            ares = CalcCd2_(Left(astr, 10))
            If CInt(Mid(astr, 11, 1)) <> ares Then
                Screen.activeform.sbWriteStatusMessage "Υποχρεωτικό πεδίο"
                GoTo chkFailed
            End If
'            DisplayMask = "000/000000-00"
        End If
    End Select
    ChkValid = True
    Screen.activeform.sbWriteStatusMessage ""
    If ValidationFlag And EditFlag(inPhase) Then
        ValidOk = True
        ValidationControl.Run FldName & "_Validation"
        If Not ValidOk Then
            If ValidationError <> "" Then
                Screen.activeform.sbWriteStatusMessage ValidationError
            Else
                Screen.activeform.sbWriteStatusMessage "Λάθος Κατα τον Ελεγχο του Πεδίου"
            End If
            GoTo chkFailed
        End If
    End If
    
    Exit Function
chkFailed:
    ChkValid = False
End Function

Private Function GetFormatedText() As String
Dim astr  As String
Dim apos As Integer, bpos As Integer
On Error GoTo ErrorPos:
    If DisplayMask <> "" Then
        If ClearText = "" Or Trim(Replace(ClearText, Chr(160), " ")) = "" Then Exit Function
        Select Case ValidationCode
        Case 2
            astr = ClearText
            If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
        Case 10
            astr = ClearText
            If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
            If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = astr & CalcCd2_(Left(astr, 10))
        Case 3
            astr = ClearText
            If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
            If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 9, "0", "L")
            astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
            astr = astr & CalcCd2_(Left(astr, 10))
        Case Else
            If ClearText <> "" Then
                apos = InStr(DisplayMask, ".")
                If apos > 0 Then
                    bpos = Len(DisplayMask) - apos
                    
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
            GetFormatedText = format(astr, DisplayMask)
        Else
            GetFormatedText = ""
        End If
    Else: GetFormatedText = ClearText
    End If
    Exit Function
ErrorPos:
    Call NBG_LOG_MsgBox("Πεδίο :" & CStr(FldNo) & vbCrLf & " Λάθος:" & Err.number & Err.description, True)
End Function

Public Property Get Text() As String
    Text = ClearText
End Property

Public Property Let Text(aText As String)
    ClearText = aText
    VControl.Text = GetFormatedText
    PropertyChanged "Text"
End Property

Public Property Get ControlText() As String
    ControlText = VControl.Text
End Property

Public Property Let ControlText(aText As String)
    EnableEditChk = False
    ClearText = aText
    VControl.Text = aText
    PropertyChanged "ControlText"
    EnableEditChk = True
End Property

Public Property Let DisplayText(aText As String)
    EnableEditChk = False
    VControl.Text = aText
    PropertyChanged "DisplayText"
    EnableEditChk = True
End Property

Public Property Get FormatedText() As String
    FormatedText = GetFormatedText
End Property

Public Sub Clear()
    ClearText = ""
    VControl.Text = GetFormatedText
End Sub

Public Sub GetDate8(invalue)
Dim astr As String
    astr = Right("00" & CStr(Day(invalue)), 2) & Right("00" & CStr(Month(invalue)), 2) & Right("0000" & CStr(Year(invalue)), 4)
    If astr <> "00000000" And astr <> "01010100" Then
        EnableEditChk = False
        ClearText = astr
        VControl.Text = GetFormatedText
        EnableEditChk = True
    End If
End Sub

Public Sub GetDate6(invalue)
Dim astr As String
    astr = Right("00" & CStr(Day(invalue)), 2) & Right("00" & CStr(Month(invalue)), 2) & Right("00" & CStr(Year(invalue)), 2)
    If astr <> "000000" Then
        EnableEditChk = False
        ClearText = astr
        VControl.Text = GetFormatedText
        EnableEditChk = True
    End If
End Sub

Public Property Get AsDate() As Date
Dim astr As String
    If Trim(ClearText) = "" Then
        AsDate = CDate("1900/01/01")
    Else
        astr = GetFormatedText
        If Not IsDate(astr) Then
            AsDate = CDate("1900/01/01")
        ElseIf CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
            AsDate = CDate("1900/01/01")
        ElseIf CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
            AsDate = CDate("1900/01/01")
        Else
            AsDate = CDate(astr)
        End If
    End If
        
'        astr = GetFormatedText
'        If Not IsDate(astr) Then
'            Screen.ActiveForm.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
'            GoTo chkFailed
'        End If
'        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
'            Screen.ActiveForm.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
'            GoTo chkFailed
'        End If
'        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
'            Screen.ActiveForm.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
'            GoTo chkFailed
'        End If
        
        
        
'        If CInt(Left(astr, 2)) <> Day(CDate(astr)) Then
'            Screen.ActiveForm.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
'            GoTo chkFailed
'        End If
'        If CInt(Mid(astr, 4, 2)) <> Month(CDate(astr)) Then
'            Screen.ActiveForm.sbWriteStatusMessage "Λάθος Μορφή Ημερομηνίας"
'            GoTo chkFailed
'        End If
        
End Property

Public Property Get AsInteger() As Long
On Error GoTo alpha
    If Trim(ClearText) = "`" Then AsInteger = 0 Else AsInteger = CLng("0" & Trim(ClearText))
    Exit Property
alpha:
    AsInteger = 0
End Property

Public Property Get AsDouble() As Double
On Error GoTo alpha
    If Trim(ClearText) = "`" Then
        AsDouble = 0
    Else
        If Trim(ClearText) = "" Then AsDouble = 0 Else AsDouble = CDbl(Trim(ClearText))
    End If
    Exit Property
alpha:
    AsDouble = 0
End Property

Public Property Get AsString() As String
    AsString = Trim(ClearText)
End Property

Public Property Get OutText() As String
    OutText = OutBuffText
End Property

Public Property Let OutText(aText As String)
    OutBuffText = aText
    PropertyChanged "OutText"
End Property

Public Property Get InText() As String
    InText = InBuffText
End Property

Public Property Let InText(aText As String)
    InBuffText = aText
    PropertyChanged "InText"
End Property

Public Sub SetFocus()
    VControl.SetFocus
End Sub

Public Sub FormatBeforeOut()
'01: ΧΩΡΙΣ VALIDATION
'02: Λογαριασμός με CD,
'03: Λογαριασμός χωρίς CD,
'04: Δάνειο με CD,
'05: Δάνειο χωρίς CD,
'06: Ημερομηνία
'07: Αριθμός Εγγραφής
'08: Ειδικός με CD
'09: Γενικός Λογαριασμός Δανείου
'10: Λογαριασμός Καταθέσεων με 1 CD
'11: Τραπεζική Επιταγή
'12: ΕΘΝΟΚΑΡΤΑ
'13: Τραπεζική Εντολή
'14: ΙΔΙΩΤΙΚΗ ΕΠΙΤΑΓΗ
'21: Λογαριασμος Καταθέσεων Γερμανία
    Dim astr As String
    If ValidationCode = 2 Then
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 8 Then astr = StrPad_(astr, 8, "0", "L")
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        OutBuffText = astr
    ElseIf ValidationCode = 3 Then
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 6 Then astr = StrPad_(astr, 6, "0", "L")
            If Len(astr) = 6 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 9, "0", "L")
            astr = astr & CalcCd1_(Mid(astr, 4, 6), 6)
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        OutBuffText = astr
    ElseIf ValidationCode = 10 Then
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 7 Then astr = StrPad_(astr, 7, "0", "L")
            If Len(astr) = 7 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            If Len(astr) < 10 Then astr = StrPad_(astr, 10, "0", "L")
            astr = astr & CalcCd2_(Left(astr, 10))
        End If
        OutBuffText = astr
    ElseIf ValidationCode = 21 Then
        astr = ClearText
        If astr <> "" Then
            If Len(astr) < 8 Then astr = cBRANCH & StrPad_(astr, 7, "0", "L") & "0"
            If Len(astr) = 8 Then astr = IIf(cBRANCH > 1000, cBRANCH \ 100, cBRANCH) & astr
            astr = StrPad_(astr, 11, "0", "L")
        End If
        OutBuffText = astr
    Else
        If ClearText <> "" And OutMask <> "" Then
            OutBuffText = format(ClearText, OutMask)
        Else
            OutBuffText = ClearText
        End If
    End If
    If FormatBeforeOutFlag Then
        ValidationControl.Run FldName & "_FormatBeforeOut"
    End If
    
End Sub

Public Sub FormatAfterIn()
    ClearText = InText
    VControl.Text = GetFormatedText
    If FormatAfterInFlag Then
        ValidationControl.Run FldName & "_FormatAfterIn"
    End If
End Sub

Public Sub HandleEdit(inPhase)
    VControl.TabStop = EditFlag(inPhase)
    VControl.Locked = Not EditFlag(inPhase)
    If Not EditFlag(inPhase) Then
        VControl.BackColor = &HD0D0D0
    Else
        If ChoiceCount = 0 Then VControl.BackColor = &H80000005 Else VControl.BackColor = &HFFFF&
    End If
End Sub

'Public Sub SetAsReadOnly()
'    EditFlag(owner.ProcessPhase) = False
'    VControl.TabStop = False
'    VControl.Locked = True
'    VControl.BackColor = &HD0D0D0
'End Sub

Public Sub SetPrompt(inPrompt As Label)
    Set Prompt = inPrompt
End Sub

Public Sub FinalizeEdit()
    EnableEditChk = False
    If Trim(GetFormatedText) <> Trim(VControl.Text) Then
        ClearText = Trim(VControl.Text)
    End If
'Συμπληρώνει τα δεκαδικά στα αριθμητικά πεδία άν ο χρήστης έχει πατήσει το ,
    If EditType = etNUMBER Then
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
    
    EnableEditChk = True
End Sub

Private Sub VControl_Change()
Dim i As Integer
Dim astr As String
Dim aNum As Double
    If Not EnableEditChk Then Exit Sub
    If Right(VControl.Text, 1) = vbTab Then
        VControl.Text = OLDTEXT
            VControl.SelStart = Len(VControl.Text)
            VControl.SelLength = 0
            Beep
            Exit Sub
    End If
    
    If EditLength > 0 Then
        If Len(VControl.Text) > EditLength Then
            VControl.Text = OLDTEXT
            VControl.SelStart = Len(VControl.Text)
            VControl.SelLength = 0
            Beep
        ElseIf Len(VControl.Text) > 0 Then
            
            If EditType = etNUMBER Then
                astr = VControl.Text
                On Error GoTo ErrorPos
                aNum = CDbl(astr)
                GoTo noErrorPos
ErrorPos:
                If astr = "-" Or astr = "+" Or astr = "." Then GoTo noErrorPos
                VControl.Text = OLDTEXT
                VControl.SelStart = Len(VControl.Text)
                VControl.SelLength = 0
                Beep
noErrorPos:
            Else
                astr = VControl.Text
                If astr <> UCase(astr) Then
                    Dim aselpos As Integer, asellength As Integer
                    aselpos = VControl.SelStart
                    asellength = VControl.SelLength
                    
                    astr = UCase(astr)
                    VControl.Text = astr
                    VControl.SelStart = aselpos
                    VControl.SelLength = asellength
                    
                End If
            End If
            
'            For I = 1 To Len(VControl.Text)
'                If EditType = etNUMBER _
'                And (Mid(VControl.Text, I, 1) < "0" _
'                Or Mid(VControl.Text, I, 1) > "9") _
'                And Mid(VControl.Text, I, 1) <> "-" Then
'                    VControl.Text = oldText
'                    VControl.SelStart = Len(VControl.Text)
'                    VControl.SelLength = 0
'                    Beep
'                    Exit For
'                End If
'            Next I
        End If
    End If
    OLDTEXT = VControl.Text
End Sub

Private Sub VControl_GotFocus()
    VControl.Text = ClearText
    VControl.SelStart = 0
    VControl.SelLength = Len(ClearText)
    EnableEditChk = True
    Set owner.ActiveTextBox = Me
End Sub
Public Sub ProcessSendAction()
Dim res As Boolean
    FinalizeEdit
    res = owner.ControlLostFocus(Me)
    owner.SEND VControl
End Sub

Public Sub GetHelpFromStkList()
    StkSelection.CD = ""
    StkSelection.name = ""
    
    owner.ScaleMode = vbTwips
    HelpFrm.HelpList.Clear
    HelpFrm.HelpTxt.BackColor = &H80000004
    HelpFrm.HelpTxt.ForeColor = &H8000000D

Dim aSize As Integer, i As Integer, k As Integer, astr  As String
Dim aList() As StkRecord

Dim ars As New ADODB.Recordset
    
    Dim allflag As Boolean, selectedrow As Integer
    allflag = True
    
    ars.open ReadDir & "stocks.xml", "Provider=msPersist", adOpenStatic, adLockReadOnly
    If Trim(VControl.Text) <> "" Then
        allflag = False
        ars.Filter = " CD = " & VControl.Text
        
        'ars.Open "Select count(*) as res from tbl_stocklist where CD = " & VControl.Text, ado_DB, adOpenStatic, adLockOptimistic
        On Error Resume Next: ars.MoveFirst
        aSize = ars.RecordCount '  ars!res
        If aSize = 0 Then
            allflag = True
            ars.Filter = ""
            'ars.Close
            'ars.Open "Select count(*) as res from tbl_stocklist", ado_DB, adOpenStatic, adLockOptimistic
            aSize = ars.RecordCount '  aSize = ars!res
        End If
        ReDim aList(aSize)
        'ars.Close
    Else
        'ars.Open "Select count(*) as res from tbl_stocklist", ado_DB, adOpenStatic, adLockOptimistic
        ars.Filter = ""
        On Error Resume Next: 'ars.MoveFirst
        aSize = ars.RecordCount ' aSize = ars!res
        ReDim aList(aSize)
        'ars.Close
    End If
    
    'ars.Open "Select * from tbl_stocklist " & IIf(allflag, "", " where CD = " & VControl.Text) & " order by XAACD", ado_DB, adOpenStatic, adLockOptimistic
    
    ars.MoveFirst
    i = 0: selectedrow = -1
    Do While Not ars.Eof
        If VControl.Text <> "" And VControl.Text = CStr(ars!CD) Then selectedrow = i
        aList(i).CD = ars.fields("CD").value
        'aList(i).Name = Left(ars!XAACD & String(6, " "), 6) & Left(ars!Name & String(30, " "), 29) & " " & Left(ars!CD & String(5, " "), 5) & _
        '    " " & StrPad_(CStr(ars!unit), 4, "0", "L") & Right(String(10, " ") & CStr(ars!hvalue), 10) & _
        '    Right(String(10, " ") & CStr(ars!lvalue), 10) & Right(String(10, " ") & CStr(ars!Value), 10)
        
        aList(i).name = Left(ars!XAACD & String(6, " "), 6) & Left(ars!name & String(30, " "), 29) & " " & Right(String(5, "0") & ars!CD, 5) & _
            " " & Right(String(5, " ") & ars!unit, 5) & _
             Right(String(8, " ") & CStr(ars!value), 8)
        
        astr = ars!valueDate
        i = i + 1
        ars.MoveNext
        If ars.RecordCount = 1 Then Exit Do
    Loop
    'astr = "ΚΧΑΑ " & " ΟΝΟΜΑΣΙΑ                     ΚΩΔ. " & " ΜΟΝΑΔΑ " & "  AN. TIMH" & " KAT. TIMH " & astr
    astr = "ΚΧΑΑ " & " ΟΝΟΜΑΣΙΑ                      ΚΩΔΙΚΟΣ " & " ΜΟΝ. " & astr
    HelpFrm.TitlesLbl.Caption = astr
    
    For i = 0 To aSize - 1
        HelpFrm.HelpList.AddItem (aList(i).name)
        HelpFrm.Selections.add (i)
    Next i
    
'    If SelectedRow <> -1 Then HelpFrm.HelpList.ListIndex = SelectedRow
        

    'HelpFrm.Width = VControl.Width
    HelpFrm.height = 5115
    HelpFrm.width = 9900
    
    HelpFrm.Left = 100
    HelpFrm.Top = (owner.height - HelpFrm.height) \ 2 '  owner.Top + owner.ActiveControl.Top + 340 + owner.ActiveControl.Height
        
    If HelpFrm.HelpList.ListCount > 0 Then
        If HelpFrm.HelpList.ListCount > 1 Then HelpFrm.SelectedIndex = selectedrow
'        If SelectedRow <> -1 Then HelpFrm.HelpList.ListIndex = SelectedRow
        HelpFrm.Show vbModal, Me
        If HelpRetValue <> "" Then
            StkSelection = aList(HelpRetValue)
        End If
            
    End If
End Sub

Public Sub GetHelpFrom3rdList()

End Sub

Private Sub VControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim aMsg As String, res As Boolean, i As Integer
    
    If KeyCode = vbKeyUp Then
        If owner.EditFieldsCount > 1 Then SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        If owner.EditFieldsCount > 1 Then SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyF1 And IsEditable(owner.processphase) Then
        owner.ScaleMode = vbTwips
    
        HelpFrm.HelpList.Clear
        
        HelpFrm.HelpTxt.BackColor = &H80000004
        HelpFrm.HelpTxt.ForeColor = &H8000000D
        If ScrHelp <> "" Then HelpFrm.HelpTxt.Text = ScrHelp
        For i = 0 To ChoiceCount - 1
            HelpFrm.HelpList.AddItem Choices(i).LineCD + "  " + Choices(i).LineText
            HelpFrm.Selections.add Choices(i).LineCD
            
        Next i
        

        HelpFrm.height = 5115
        
        If HelpFrm.HelpList.ListCount > 0 Then
            HelpFrm.Show vbModal, Me
            If HelpRetValue <> "" Then
                
                EnableEditChk = False
                ClearText = HelpRetValue
                VControl.Text = HelpRetValue
                EnableEditChk = True
            End If
            
            
        Else
            aMsg = MsgBox("Δεν υπάρχει βοήθεια για αυτό το πεδίο", vbOKOnly, "On Line Εφαρμογή")
        End If
    ElseIf KeyCode = 111 Then '/
        If owner.EditFieldsCount > 1 Then SendKeys "000": ClearLastKey = True
    ElseIf KeyCode = 106 Then '*
        If owner.EditFieldsCount > 1 Then SendKeys "00": ClearLastKey = True
    End If

End Sub

Private Sub VControl_KeyPress(KeyAscii As Integer)
Dim apos As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        If owner.EditFieldsCount > 1 Or owner.TabbedControls.count > 1 Then
            If Not owner.DisableEnterKey Then
                SendKeys "{TAB}"
            End If
        End If
    ElseIf KeyAscii = 47 Or KeyAscii = 42 Then
        If ClearLastKey Then KeyAscii = 0: ClearLastKey = False
    
    End If
End Sub

Private Sub VControl_LostFocus()
Dim res As Boolean
    FinalizeEdit
    res = owner.ControlLostFocus(Me)
    If res Then
        EnableEditChk = False
        VControl.Text = GetFormatedText
        EnableEditChk = True
    End If
End Sub

Private Sub HiddenFld_GotFocus()
    If InvertTab Then
        InvertTab = False
        SendKeys "+{TAB}"
    Else
        SendKeys "{TAB}"
    End If
End Sub

Private Sub UserControl_Hide()
On Error Resume Next
    Prompt.Visible = False
End Sub

Private Sub UserControl_Initialize()
    ValidationFlag = False
    VControl.CausesValidation = True
    VControl.Visible = True
End Sub

Private Sub UserControl_Resize()
    VControl.Left = 0
    VControl.Top = 0
    VControl.width = width
    VControl.height = height
    
End Sub

Private Sub UserControl_Show()
    If Not (Prompt Is Nothing) Then Prompt.Visible = True
End Sub

Public Property Get Enabled() As Boolean
    Enabled = VControl.Enabled
End Property

Public Property Let Enabled(value As Boolean)
    VControl.Enabled = value
    VControl.TabStop = value
    VControl.BackColor = IIf(value And Not VControl.Locked, IIf(ChoiceCount = 0, &H80000005, &HFFFF&), &H80000004)
End Property

Public Sub SetCurrHelp()

ReDim Choices(ChoiceSuperSetCount)
Dim i As Integer
If cVersion >= 20020101 Then

        ReDim Choices(11): ReDim ChoicesSuperSet(11)
        ChoiceCount = 11: ChoiceSuperSetCount = 11
        
        Choices(0).LineCD = "002": Choices(0).LineText = "USD - ΔΟΛΑΡΙΟ ΗΠΑ"
        Choices(1).LineCD = "008": Choices(1).LineText = "CHF - ΦΡΑΓΚΟ ΕΛΒΕΤΙΑΣ"
        Choices(2).LineCD = "010": Choices(2).LineText = "CAD - ΔΟΛΑΡΙΟ ΚΑΝΑΔΑ"
        Choices(3).LineCD = "012": Choices(3).LineText = "SEK - ΚΟΡΟΝΑ ΣΟΥΗΔΙΑΣ"
        Choices(4).LineCD = "013": Choices(4).LineText = "NOK - ΚΟΡΟΝΑ ΝΟΡΒΗΓΙΑΣ"
        Choices(5).LineCD = "014": Choices(5).LineText = "DKK - ΚΟΡΟΝΑ ΔΑΝΙΑΣ"
        Choices(6).LineCD = "032": Choices(6).LineText = "CYP - ΛΙΡΑ ΚΥΠΡΟΥ"
        Choices(7).LineCD = "043": Choices(7).LineText = "JPY - ΓΕΝ ΙΑΠΩΝΙΑΣ"
        Choices(8).LineCD = "049": Choices(8).LineText = "AUD - ΔΟΛΑΡΙΟ ΑΥΣΤΡΑΛΙΑΣ"
        Choices(9).LineCD = "050": Choices(9).LineText = "GBP - ΛΙΡΑ ΑΓΓΛΙΑΣ"
        Choices(10).LineCD = "070": Choices(10).LineText = "EUR - ΕΥΡΩ"
Else
        ReDim Choices(22): ReDim ChoicesSuperSet(22)
        ChoiceCount = 22: ChoiceSuperSetCount = 22

        Choices(0).LineCD = "001": Choices(0).LineText = "GRD - ΔΡΑΧΜΕΣ"
        Choices(1).LineCD = "002": Choices(1).LineText = "USD - ΔΟΛΑΡΙΟ ΗΠΑ"
        Choices(2).LineCD = "003": Choices(2).LineText = "FRF - ΦΡΑΓΚΟ ΓΑΛΛΙΑΣ"
        Choices(3).LineCD = "004": Choices(3).LineText = "ITL - ΛΙΡΕΤΑ ΙΤΑΛΙΑΣ"
        Choices(4).LineCD = "005": Choices(4).LineText = "DEM - ΜΑΡΚΟ ΓΕΡΜΑΝΙΑΣ"
        Choices(5).LineCD = "008": Choices(5).LineText = "CHF - ΦΡΑΓΚΟ ΕΛΒΕΤΙΑΣ"
        Choices(6).LineCD = "010": Choices(6).LineText = "CAD - ΔΟΛΑΡΙΟ ΚΑΝΑΔΑ"
        Choices(7).LineCD = "012": Choices(7).LineText = "SEK - ΚΟΡΟΝΑ ΣΟΥΗΔΙΑΣ"
        Choices(8).LineCD = "013": Choices(8).LineText = "NOK - ΚΟΡΟΝΑ ΝΟΡΒΗΓΙΑΣ"
        Choices(9).LineCD = "014": Choices(9).LineText = "DKK - ΚΟΡΟΝΑ ΔΑΝΙΑΣ"
        Choices(10).LineCD = "017": Choices(10).LineText = "NLG - ΦΙΟΡΙΝΙ ΟΛΛΑΝΔΙΑΣ"
        Choices(11).LineCD = "018": Choices(11).LineText = "ESP - ΠΕΣΕΤΑ ΙΣΠΑΝΙΑΣ"
        Choices(12).LineCD = "021": Choices(12).LineText = "ATS - ΣΕΛΙΝΙ ΑΥΣΤΡΙΑΣ"
        Choices(13).LineCD = "023": Choices(13).LineText = "PTE - ΕΣΚΟΥΔΟΣ ΠΟΡΤΟΓΑΛΙΑΣ"
        Choices(14).LineCD = "032": Choices(14).LineText = "CYP - ΛΙΡΑ ΚΥΠΡΟΥ"
        Choices(15).LineCD = "035": Choices(15).LineText = "FIM - ΜΑΡΚΟ ΦΙΝΛΑΝΔΙΑΣ"
        Choices(16).LineCD = "043": Choices(16).LineText = "JPY - ΓΕΝ ΙΑΠΩΝΙΑΣ"
        Choices(17).LineCD = "049": Choices(17).LineText = "AUD - ΔΟΛΑΡΙΟ ΑΥΣΤΡΑΛΙΑΣ"
        Choices(18).LineCD = "050": Choices(18).LineText = "GBP - ΛΙΡΑ ΑΓΓΛΙΑΣ"
        Choices(19).LineCD = "057": Choices(19).LineText = "IEP - ΛΙΡΑ ΙΡΛΑΝΔΙΑΣ"
        Choices(20).LineCD = "059": Choices(20).LineText = "BEF - ΦΡΑΓΚΟ ΒΕΛΓΙΟΥ"
        Choices(21).LineCD = "070": Choices(21).LineText = "EUR - ΕΥΡΩ"
End If

End Sub


Public Sub RemoveHelpItem(invalue)
Dim i As Integer
Dim j As Integer
Dim CompareChoice
    
If IsNumeric(invalue) Then invalue = CLng(invalue) Else invalue = CStr(invalue)
    
    For i = 0 To ChoiceCount - 1
        If IsNumeric(Choices(i).LineCD) Then CompareChoice = CLng(Choices(i).LineCD) Else CompareChoice = CStr(Choices(i).LineCD)
        If CompareChoice = invalue Then
            If i < ChoiceCount - 1 Then
                Choices(i).LineCD = Choices(i + 1).LineCD: Choices(i).LineText = Choices(i + 1).LineText
                For j = i + 1 To ChoiceCount - 2
                    Choices(j).LineCD = Choices(j + 1).LineCD: Choices(j).LineText = Choices(j + 1).LineText
                Next j
            End If
            Choices(ChoiceCount - 1).LineCD = "": Choices(ChoiceCount - 1).LineText = ""
            ChoiceCount = ChoiceCount - 1
            Exit Sub
        End If
    Next i

End Sub

Public Sub GetHelpSuperSet()
ReDim Choices(ChoiceSuperSetCount)
Dim i As Integer
    For i = 0 To ChoiceSuperSetCount - 1
        Choices(i).LineCD = ChoicesSuperSet(i).LineCD
        Choices(i).LineText = ChoicesSuperSet(i).LineText
    Next i
    ChoiceCount = ChoiceSuperSetCount
End Sub

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement, inPhase As Integer)
Dim i As Integer, aType As Integer, aTypeNode As MSXML2.IXMLDOMElement
    
    DisplayFlag(inPhase) = NodeBooleanFld(inNode, "ScrDisplay", fldModel)
    EditFlag(inPhase) = NodeBooleanFld(inNode, "ScrEntry", fldModel)
    OptionalFlag(inPhase) = NodeBooleanFld(inNode, "ScrOptional", fldModel)
    OutCode(inPhase) = NodeIntegerFld(inNode, "OutBuffCDA", fldModel)
    OutCodeEx(inPhase) = NodeStringFld(inNode, "OutBuffCDAex", fldModel)
    OutBuffLength(inPhase) = NodeIntegerFld(inNode, "OutBuffLengthA", fldModel)
    OutBuffPos(inPhase) = NodeIntegerFld(inNode, "OutBuffPosA", fldModel)
    InBuffLength(inPhase) = NodeIntegerFld(inNode, "InBuffLengthA", fldModel)
    InBuffPos(inPhase) = NodeIntegerFld(inNode, "InBuffPosA", fldModel)
    JournalBeforeOut(inPhase) = NodeBooleanFld(inNode, "JournalBeforeOut", fldModel)
    JournalAfterIn(inPhase) = NodeBooleanFld(inNode, "JournalAfterIn", fldModel)

    If inPhase = 1 Then
        Set owner = inOwner
        FldNo = NodeIntegerFld(inNode, "FldNo", fldModel)
        FldName = "Fld" & StrPad_(CStr(FldNo), 3, "0", "L") 'NodeStringFld(inNode, "Name", fldModel)
        FldName2 = UCase(NodeStringFld(inNode, "NAME", fldModel))
        If InStr(ReservedControlPrefixes, "," & Left(FldName2, 3) & ",") > 0 Then FldName2 = ""
        name = IIf(FldName2 <> "", FldName2, FldName)
            
        TotalName = NodeStringFld(inNode, "TotalName", fldModel)
        TotalPos = NodeIntegerFld(inNode, "TotalPos", fldModel)

        aType = NodeIntegerFld(inNode, "FldType", fldModel)
        If aType = 0 Then
            Set aTypeNode = Nothing
        Else
            Set aTypeNode = FldTypeList.item("T" & Trim(Str(aType)))
        End If
        
        Set ValidationControl = inProcessControl
        ValidationControl.AddObject FldName, Me, True
        If Trim(FldName2) <> "" And UCase(Trim(FldName2)) <> UCase(Trim(FldName)) Then
            On Error GoTo FldRegistrationError
            ValidationControl.ExecuteStatement "Set " & FldName2 & "=" & FldName
            GoTo FldRegistrationOk
FldRegistrationError:
            MsgBox "Λάθος κατα τη δήλωση του πεδίου: " & FldName & ":" & FldName2
FldRegistrationOk:
        End If
        
        ScrLeft = NodeIntegerFld(inNode, "ScrX", fldModel)
        ScrWidth = NodeIntegerFld(inNode, "ScrWidth", fldModel)
        ScrTop = NodeIntegerFld(inNode, "ScrY", fldModel) * 290
        ScrHeight = NodeIntegerFld(inNode, "ScrHeight", fldModel) * 285
        ScrHelp = NodeStringFld(inNode, "ScrHelp", fldModel)

        DocX = NodeIntegerFld(inNode, "DocX", fldModel)
        DocY = NodeIntegerFld(inNode, "DocY", fldModel)
        DocWidth = NodeIntegerFld(inNode, "DocWidth", fldModel)
        DocHeight = NodeIntegerFld(inNode, "DocHeight", fldModel)
        Title = NodeStringFld(inNode, "DocTitle", fldModel)
        TitleX = NodeIntegerFld(inNode, "DocTitleX", fldModel)
        TitleY = NodeIntegerFld(inNode, "DocTitleY", fldModel)
        TitleWidth = NodeIntegerFld(inNode, "DocTitleWidth", fldModel)
        TitleHeight = NodeIntegerFld(inNode, "DocTitleHeight", fldModel)
        
        TTabIndex = NodeIntegerFld(inNode, "TabIndex", fldModel)

        Select Case NodeIntegerFld(inNode, "ScrAlign", fldModel)
        Case 1
            VControl.Alignment = vbLeftJustify
            DocAlign = 1
        Case 2
            VControl.Alignment = vbRightJustify
            DocAlign = 2
        Case Else
            If Not (aTypeNode Is Nothing) Then
                If Trim(aTypeNode.selectSingleNode("ALIGN").Text) = "1" Then
                    VControl.Alignment = vbLeftJustify
                    DocAlign = 1
                ElseIf Trim(aTypeNode.selectSingleNode("ALIGN").Text) = "2" Then
                    VControl.Alignment = vbRightJustify
                    DocAlign = 2
                End If
            End If
        End Select
        Select Case NodeIntegerFld(inNode, "DocAlign", fldModel)
        Case 1
            DocAlign = 1
        Case 2
            DocAlign = 2
        End Select
        
        If Not (aTypeNode Is Nothing) Then
            If aTypeNode.selectSingleNode("VALIDATIONCODE").Text <> "" Then
                ValidationCode = aTypeNode.selectSingleNode("VALIDATIONCODE").Text
            Else
                ValidationCode = 0
            End If
            
        End If
        
        LabelName = "TLabel" & StrPad_(CStr(FldNo), 3, "0", "L")
        Set Prompt = parent.Controls.add("Vb.Label", LabelName)
        Prompt.BackColor = parent.BackColor
        Prompt.AutoSize = False
        Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", fldModel)
        
        parent.ScaleMode = vbCharacters
        
        Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", fldModel)
        Prompt.width = NodeIntegerFld(inNode, "ScrPromptWidth", fldModel)
        parent.ScaleMode = vbTwips
        Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", fldModel) * 290
        Prompt.height = NodeIntegerFld(inNode, "ScrPromptHeight", fldModel) * 285
    
        Dim astr As String
        astr = NodeStringFld(inNode, "ScrValidationScript", fldModel)
        If astr <> "" Then
            ValidationControl.AddCode " Public Sub " & _
                FldName & "_Validation " & vbCrLf & _
                astr & vbCrLf & "End Sub"
            ValidationFlag = True
        End If
        astr = NodeStringFld(inNode, "FormatBeforeOutScript", fldModel)
        If astr <> "" Then
            ValidationControl.AddCode " Public Sub " & _
                FldName & "_FormatBeforeOut " & vbCrLf & _
                astr & vbCrLf & "End Sub"
            FormatBeforeOutFlag = True
        End If
        astr = NodeStringFld(inNode, "FormatAfterInScript", fldModel)
        If astr <> "" Then
            ValidationControl.AddCode " Public Sub " & _
                FldName & "_FormatAfterIn " & vbCrLf & _
                astr & vbCrLf & "End Sub"
            FormatAfterInFlag = True
        End If
        
        Editmask = "": DisplayMask = "": OutMask = ""
        If Not (aTypeNode Is Nothing) Then
            DisplayMask = aTypeNode.selectSingleNode("DISPLAYMASK").Text
            Editmask = aTypeNode.selectSingleNode("EDITMASK").Text
            OutMask = aTypeNode.selectSingleNode("OUTMASK").Text
            EditLength = aTypeNode.selectSingleNode("EDITLENGTH").Text
            EditType = aTypeNode.selectSingleNode("EDITTYPE").Text
        End If
        
        HPSOutStruct = NodeStringFld(inNode, "HPSOutStruct", fldModel)
        HPSOutPart = NodeStringFld(inNode, "HPSOutPart", fldModel)
        HPSInStruct = NodeStringFld(inNode, "HPSInStruct", fldModel)
        HPSInPart = NodeStringFld(inNode, "HPSInPart", fldModel)
        
        If NodeStringFld(inNode, "ScrDisplayMask", fldModel) <> "" Then _
            DisplayMask = NodeStringFld(inNode, "ScrDisplayMask", fldModel)
        If NodeStringFld(inNode, "ScrEditMask", fldModel) <> "" Then _
            Editmask = NodeStringFld(inNode, "ScrEditMask", fldModel)
        If NodeStringFld(inNode, "OutMask", fldModel) <> "" Then _
            OutMask = NodeStringFld(inNode, "OutMask", fldModel)
        DocMask = DisplayMask
        If NodeStringFld(inNode, "DocDisplayMask", fldModel) <> "" Then _
            DocMask = NodeStringFld(inNode, "DocDisplayMask", fldModel)
        
        If NodeIntegerFld(inNode, "ScrEditLength", fldModel) <> 0 Then _
            EditLength = NodeIntegerFld(inNode, "ScrEditLength", fldModel)
        If NodeIntegerFld(inNode, "ScrEditType", fldModel) <> 0 Then _
            EditType = NodeIntegerFld(inNode, "ScrEditType", fldModel)
            
        On Error GoTo NoSelections
        HelpFormWidth = 5805: HelpFormHeight = 5115
        If Not (inNode.selectSingleNode("SELECTIONS") Is Nothing) Then
            Dim selNode As MSXML2.IXMLDOMElement, anode As MSXML2.IXMLDOMElement
            On Error GoTo 0
            Set selNode = inNode.selectSingleNode("SELECTIONS")
            ChoiceCount = selNode.childNodes.length
            ChoiceSuperSetCount = selNode.childNodes.length
            ReDim Choices(selNode.childNodes.length)
            ReDim ChoicesSuperSet(selNode.childNodes.length)
            For i = 0 To selNode.childNodes.length - 1
                Set anode = selNode.childNodes.item(i)
                Choices(i).LineCD = Right(anode.tagName, Len(anode.tagName) - 2)
                Choices(i).LineText = anode.Text
                ChoicesSuperSet(i).LineCD = Right(anode.tagName, Len(anode.tagName) - 2)
                ChoicesSuperSet(i).LineText = anode.Text
            Next i
        End If
        
        QFldNo = NodeIntegerFld(inNode, "QFldNo", fldModel)
        PasswordChar = NodeStringFld(inNode, "ScrPasswordChar", fldModel)
        If PasswordChar <> "" Then
            VControl.Font.name = "BBSecret"
        End If
        If NodeStringFld(inNode, "SCRHELP", fldModel) <> "" Then
            VControl.ToolTipText = NodeStringFld(inNode, "SCRHELP", fldModel)
        End If
        
NoSelections:

    End If
End Sub

Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("TEXTBOX")
    Set attr = XML.createAttribute("NO")
    attr.nodeValue = UCase(Me.FldNo)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("NAME")
    attr.nodeValue = UCase(Me.FldName)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("PROMPT")
    attr.nodeValue = UCase(Me.Prompt.Caption)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.IsVisible(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("TEXT")
    attr.nodeValue = UCase(Me.Text)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FORMATEDTEXT")
    attr.nodeValue = UCase(Me.FormatedText)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("READONLY")
    attr.nodeValue = UCase(Not Me.IsEditable(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("OPTIONAL")
    attr.nodeValue = UCase(Me.IsOptional(inPhase)) '(Owner.ProcessPhase))
    elm.setAttributeNode attr
                    
    Set TranslateToProperties = elm
End Function

Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    Dim aattr As IXMLDOMAttribute
    If inPhase = 0 Then inPhase = 1
    For Each aattr In elm.Attributes
        Select Case aattr.baseName
            Case "TEXT"
                Me.Text = aattr.value
            Case "READONLY"
                If UCase(aattr.value) = "FALSE" Then
                    Me.SetEditableNoRefresh inPhase, True
                ElseIf UCase(aattr.value) = "TRUE" Then
                    Me.SetEditableNoRefresh inPhase, False
                End If
            Case "VISIBLE"
                If UCase(aattr.value) = "FALSE" Then
                    Me.SetDisplay inPhase, False
                ElseIf UCase(aattr.value) = "TRUE" Then
                    Me.SetDisplay inPhase, True
                End If
            Case "OPTIONAL"
                If UCase(aattr.value) = "FALSE" Then
                    OptionalFlag(inPhase) = False
                ElseIf UCase(aattr.value) = "TRUE" Then
                    OptionalFlag(inPhase) = True
                End If
            Case "PROMPT"
                Me.Prompt.Caption = aattr.value
            Case "DISPLAYMASK"
                DisplayMask = aattr.value
            Case "OUTMASK"
                OutMask = aattr.value
            Case "DOCMASK"
                DocMask = aattr.value
            Case "EDITLENGTH"
                EditLength = aattr.value
            Case "EDITTYPE"
                If UCase(aattr.value) = "NONE" Then
                    EditType = etNONE
                ElseIf UCase(aattr.value) = "TEXT" Then
                    EditType = etTEXT
                ElseIf UCase(aattr.value) = "NUMBER" Then
                    EditType = etNUMBER
                End If
            Case "TABSTOP"
                If UCase(aattr.value) = "TRUE" Then
                    Tabbed = True
                ElseIf UCase(aattr.value) = "FALSE" Then
                    Tabbed = False
                End If
            Case "TABINDEX"
                TTabIndex = aattr.value
        End Select
    Next aattr
End Sub


