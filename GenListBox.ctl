VERSION 5.00
Begin VB.UserControl GenListBox 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   ScaleHeight     =   1920
   ScaleWidth      =   7965
   Begin VB.ListBox Control 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "GenListBox.ctx":0000
      Left            =   120
      List            =   "GenListBox.ctx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "GenListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As Form
Public LstNo As Integer, LstName As String, LstName2 As String, LabelName As String, name As String
Public Prompt As Label
Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer
Private DisplayFlag As Boolean
Private ValidationControl As ScriptControl
Private GetLinesFromBuffer As Boolean
Private MinInLineLength As Integer, FromColumn As Integer, ToColumn As Integer
Private ExcludeFirstLines As Integer, ExcludeLastLines As Integer

Public DocTopMargin, DocLeftMargin, DocRightMargin, DocBottomMargin As Integer
Private DocOrientation As Integer
Private DocPageHead, DocPageFoot As String
Private DocHeadLines, DocFootLines, DocClearLines As Integer
Private DocLines, DocColumns As Integer

Private ExamineLineFlag As Boolean
Private curPageNo As Integer

Public TTabIndex As Integer
Public EnableLaserPrinter As Boolean

Dim ForcedPageLines() As Integer, ForcedPageCount As Integer, ForcedPageNum As Integer
'Γραμμές στις οποίες γίνεται αλλαγή σελίδας, Συνολικός διαθέσιμος αριθμός εγγραφών στο πίνακα,
'αριθμός εγγραφών στο πίνακα που έχουν γίνει

Public Sub PrintToJournal()
Dim i As Integer, astr As String
If Control.ListCount = 0 Then Exit Sub
For i = 0 To Control.ListCount - 1
    astr = Control.list(i)
    If Trim(astr) <> "" Then eJournalWrite astr
Next i
    SaveJournal
End Sub

Public Sub ShowScrollBar(Optional maxlinelength As Integer)
    If IsMissing(maxlinelength) Then maxlinelength = 200
    Dim ascalemode As Integer, aSize As Integer, astr As String
    ascalemode = owner.ScaleMode: owner.ScaleMode = 3
    astr = String(maxlinelength, "W")
    aSize = owner.TextWidth(astr)
    SendMessage Control.hWnd, LB_SETHORIZONTALEXTENT, aSize, 0
    owner.ScaleMode = ascalemode
End Sub

Public Sub UnLockPrinter(owner As Form)
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
        owner.SPCPanel.UnLockPrinter
    End If
End Sub

Public Sub PrintPageHead()
Dim aLineNo As Integer, aLine As String
Dim bstr As String
    bstr = DocPageHead
    aLineNo = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        aLine = EmbedValues(owner, aLine, curPageNo)
        aLineNo = aLineNo + 1
        Printer.CurrentX = DocLeftMargin
        Printer.CurrentY = DocTopMargin + aLineNo
        Printer.Print aLine
    Loop
End Sub

Public Sub PrintPageHead_P()
Dim aLineNo As Integer, aLine As String
Dim bstr As String
    bstr = DocPageHead
    aLineNo = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        aLine = EmbedValues(owner, aLine, curPageNo)
        aLineNo = aLineNo + 1
'        Printer.CurrentX = DocLeftMargin
'        Printer.CurrentY = DocTopMargin + aLineNo
'        Printer.Print aLine
        xSetDocLine_ aLineNo, aLine
    Loop
End Sub

Public Sub PrintPageFoot()
Dim aLineNo As Integer, bstr As String, aLine As String
    bstr = DocPageFoot
    aLineNo = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        aLine = EmbedValues(owner, aLine, curPageNo)
        aLineNo = aLineNo + 1
        Printer.CurrentX = DocLeftMargin
        Printer.CurrentY = DocTopMargin + DocHeadLines + DocClearLines + aLineNo
        Printer.Print aLine
    Loop
End Sub

Public Sub PrintPageFoot_P()
Dim aLineNo As Integer, bstr As String, aLine As String
    bstr = DocPageFoot
    aLineNo = DocClearLines + DocTopMargin + DocHeadLines
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        aLine = EmbedValues(owner, aLine, curPageNo)
        aLineNo = aLineNo + 1
'        Printer.CurrentX = DocLeftMargin
'        Printer.CurrentY = DocTopMargin + DocHeadLines + DocClearLines + aLineNo
'        Printer.Print aLine
        xSetDocLine_ aLineNo, aLine
    Loop
End Sub

Public Sub PrintLines_L()
Dim oldFlag As Integer
    oldFlag = cListToPassbook
    cListToPassbook = 0
    On Error Resume Next
    PrintLines
    cListToPassbook = oldFlag
End Sub

Public Sub PrintLines(Optional PrintOCR As Boolean)
Dim OCRFlag As Boolean
If IsMissing(PrintOCR) Then OCRFlag = False Else OCRFlag = PrintOCR
eJournalWrite "ΕΚΤΥΠΩΣΗ ΣΤΟΙΧΕΙΩΝ"
eJournalWrite "Τ ΩΡΑ ΣΥΝΑΛΛΑΓΗΣ: " & FormatDateTime(Time, vbShortTime)
    SaveJournal

If cListToPassbook = 1 And Not EnableLaserPrinter Then PrintLines_P , OCRFlag: Exit Sub

Dim aFrm As SelectPrinterFrm
Set aFrm = New SelectPrinterFrm
Load aFrm
Dim res As Long
aFrm.Show vbModal, owner

Dim aPrinterName As String
aPrinterName = aFrm.SelectedPrinter

If UCase(aPrinterName) = "PASSBOOK" Then PrintLines_P False, OCRFlag: Exit Sub
Dim x As Printer
    For Each x In Printers
        If aPrinterName = x.DeviceName Then _
            Set Printer = x: Exit For
    Next
Set aFrm = Nothing
If aPrinterName = "" Then Exit Sub

If DocOrientation = 1 Then
    Printer.Orientation = vbPRORPortrait
Else
    Printer.Orientation = vbPRORLandscape
End If
Printer.ScaleMode = vbCharacters
Printer.FontSize = Control.FontSize
Printer.FontName = Control.FontName
Printer.FontName = "Courier"
Printer.Font = Control.Font
    
Dim aLine As Integer, i As Integer, allLine As Integer, PageLine As Integer, CurForcedPage As Integer
Dim astr As String
    
    curPageNo = 1: CurForcedPage = 1: allLine = 0: PageLine = 0
    PrintPageHead
    For i = 0 To Control.ListCount - 1
        allLine = allLine + 1: PageLine = PageLine + 1
        
        astr = Control.list(i)
        If Len(astr) > DocColumns - DocLeftMargin - DocRightMargin Then
            astr = Left(astr, DocColumns - DocLeftMargin - DocRightMargin)
        End If
        
        Printer.CurrentX = DocLeftMargin
        
        aLine = PageLine Mod DocClearLines
        
        If aLine > 0 Then Printer.CurrentY = aLine + DocTopMargin + DocHeadLines
        If aLine = 0 Then Printer.CurrentY = DocClearLines + DocTopMargin + DocHeadLines
        
        Printer.Print astr
        
        If CurForcedPage <= ForcedPageNum Then
            If ForcedPageLines(CurForcedPage) = allLine Then aLine = 0: CurForcedPage = CurForcedPage + 1
        End If
        
        If aLine = 0 Then PrintPageFoot
        
        If aLine = 0 And (i < Control.ListCount - 1) Then
            Printer.NewPage
            curPageNo = curPageNo + 1
            PageLine = 0
            PrintPageHead
        End If
        
    Next i
    Printer.EndDoc
End Sub

Public Sub PrintLines_P(Optional UnlockAfterPrint As Boolean, Optional OCRFlag As Boolean)
If IsMissing(UnlockAfterPrint) Then UnlockAfterPrint = False
If IsMissing(OCRFlag) Then OCRFlag = False
    
Dim PrintMsg As String
PrintMsg = owner.PrintPromptMessage
'If G0Data.count > 0 Then DocForm.Show vbModal, owner _
'Else NBG_MsgBox PrintMsg, True, PrintMsg
    
Dim allLine As Integer, aLine As Integer, bLine As Integer, i As Integer, CurForcedPage As Integer
Dim astr As String
    xClearDoc_
    curPageNo = 1: PrintPageHead_P: aLine = 0: allLine = 0: CurForcedPage = 1
    For i = 0 To Control.ListCount - 1
        aLine = aLine + 1: allLine = allLine + 1
        astr = Control.list(i)
        If Len(astr) > DocColumns - DocLeftMargin - DocRightMargin Then _
            astr = Left(astr, DocColumns - DocLeftMargin - DocRightMargin)
        bLine = aLine + DocTopMargin + DocHeadLines
        xSetDocLine_ bLine, IIf(DocLeftMargin > 0, String(DocLeftMargin, " "), "") & astr
        
        If CurForcedPage <= ForcedPageNum Then
            If ForcedPageLines(CurForcedPage) = allLine Then aLine = DocClearLines: CurForcedPage = CurForcedPage + 1
        End If
        If aLine = DocClearLines Then PrintPageFoot_P
        If aLine = DocClearLines And (i < Control.ListCount - 1) Then _
            xPrintDoc_ owner: curPageNo = curPageNo + 1: xClearDoc_: PrintPageHead_P: aLine = 0
    Next i
    If aLine <> 0 Then PrintPageFoot_P
    xPrintDoc_ owner, , OCRFlag
End Sub

Public Function ReadLines() As Boolean
ReadLines = False
Dim i As Integer
Dim astr As String
If GetLinesFromBuffer Then
    Control.Clear
    For i = 1 To ReceivedData.count
        astr = ReceivedData(i)
        astr = eJournalClearString(astr)
        
        Dim bstr As String, k As Integer
        bstr = ""
        For k = 1 To Len(astr)
            If InStr(Mid(astr, k, 1), PASSBOOK_CLEAR_STRING) <> 0 Then bstr = bstr & " " Else bstr = bstr & Mid(astr, k, 1)
        Next k
        astr = bstr
        
        If (i > ExcludeFirstLines) And (i <= ReceivedData.count - ExcludeLastLines) Then
        If Len(astr) > MinInLineLength Then
            If FromColumn > 0 And FromColumn <= Len(astr) Then
                If ToColumn > FromColumn Then
                    astr = Mid(astr, FromColumn, ToColumn - FromColumn + 1)
                Else
                    astr = Right(astr, Len(astr) - FromColumn + 1)
                End If
            ElseIf FromColumn <= Len(astr) Then
                If ToColumn > 0 Then
                    astr = Left(astr, ToColumn)
                End If
            ElseIf FromColumn > Len(astr) Then
                astr = ""
            End If
            Control.AddItem (astr)
        End If
        End If
    Next i
End If
ReadLines = True
End Function

Public Function ReadLines_28() As Boolean
Dim i As Integer, k As Integer, l As Integer
Dim astr As String, bstr As String
Dim DSTR
Clear

ReDim ForcedPageLines(10) As Integer: ForcedPageCount = 10: ForcedPageNum = 0

For i = 1 To owner.ListData.count
    astr = owner.ListData.item(i)
    astr = eJournalClearString(astr)
    
    bstr = ""
    For k = 1 To Len(astr)
        If InStr(Mid(astr, k, 1), PASSBOOK_CLEAR_STRING) <> 0 Then bstr = bstr & " " Else bstr = bstr & Mid(astr, k, 1)
    Next k
    astr = bstr

    If (Len(astr) > 5) And (Left(astr, 1) = "7") Then
        astr = Right(astr, Len(astr) - 5)
        While (Len(Trim(astr)) > 1) Or (Len(Trim(astr)) = 1 And Trim(astr) = "0")
            bstr = Left(astr, 81)
            If Len(astr) > 81 Then
                astr = Right(astr, Len(astr) - 81)
            Else
                astr = ""
            End If
            If Trim(bstr) <> "" Then
                If IsNumeric(Left(bstr, 1)) Then
                    If CInt(Left(bstr, 1)) = 0 Then
                        If ForcedPageCount = ForcedPageNum Then
                            ForcedPageCount = ForcedPageCount + 10
                            ReDim Preserve ForcedPageLines(ForcedPageCount)
                        End If
                        ForcedPageNum = ForcedPageNum + 1
                        ForcedPageLines(ForcedPageNum) = Control.ListCount + 1
                    Else
                        For k = 1 To CInt(Left(bstr, 1)) - 1
                            Control.AddItem ("")
                        Next k
                    End If
                End If
                Control.AddItem (Right(bstr, Len(bstr) - 1))
            End If
        Wend
    End If
Next i

End Function

Public Sub AddForcedPage()
    If ForcedPageCount = ForcedPageNum Then
        ForcedPageCount = ForcedPageCount + 10
        ReDim Preserve ForcedPageLines(ForcedPageCount)
    End If
    ForcedPageNum = ForcedPageNum + 1
    ForcedPageLines(ForcedPageNum) = Control.ListCount
End Sub

Public Function ReadLines_28_B() As Boolean
Dim i As Integer, k As Integer, l As Integer
Dim astr As String, bstr As String
Dim DSTR
Clear

ReDim ForcedPageLines(10) As Integer: ForcedPageCount = 10: ForcedPageNum = 0

For i = 1 To owner.ListData.count
    astr = owner.ListData.item(i)
    astr = eJournalClearString(astr)

    bstr = ""
    For k = 1 To Len(astr)
        If InStr(Mid(astr, k, 1), PASSBOOK_CLEAR_STRING) <> 0 Then bstr = bstr & " " Else bstr = bstr & Mid(astr, k, 1)
    Next k
    astr = bstr
        
    If (Len(astr) > 5) And (Left(astr, 1) = "7") Then
        
        astr = Right(astr, Len(astr) - 5)
        If Len(Trim(astr)) = 1 Then
            Control.AddItem ""
            If IsNumeric(Left(astr, 1)) Then
                If CInt(Left(astr, 1)) = 0 Then
                    AddForcedPage
                Else
                    For k = 1 To CInt(Left(astr, 1)) - 1
                        Control.AddItem ("")
                    Next k
                End If
            End If

        End If
        While (Len(Trim(astr)) > 1) Or (Len(Trim(astr)) = 1 And Trim(astr) = "0")
            bstr = Left(astr, 81)
            If Len(astr) > 81 Then
                astr = Right(astr, Len(astr) - 81)
            Else
                astr = ""
            End If
            If Trim(bstr) <> "" Then
                Control.AddItem (Right(bstr, Len(bstr) - 1))
                If IsNumeric(Left(bstr, 1)) Then
                    If CInt(Left(bstr, 1)) = 0 Then
                        AddForcedPage
                    Else
                        For k = 1 To CInt(Left(bstr, 1)) - 1
                            Control.AddItem ("")
                        Next k
                    End If
                End If
            End If
        Wend
    End If
Next i

End Function

Public Function ReadLines_16() As Boolean
Dim i As Integer, k As Integer, l As Integer
Dim astr As String, bstr As String
Dim DSTR
Clear
ReDim ForcedPageLines(10) As Integer: ForcedPageCount = 10: ForcedPageNum = 0

For i = 1 To owner.ListData.count
    astr = owner.ListData.item(i)
    astr = eJournalClearString(astr)
    
    bstr = ""
    For k = 1 To Len(astr)
        If InStr(Mid(astr, k, 1), PASSBOOK_CLEAR_STRING) <> 0 Then bstr = bstr & " " Else bstr = bstr & Mid(astr, k, 1)
    Next k
    astr = bstr
        
    If (Len(astr) > 5) And (Left(astr, 1) = "7") Then
        astr = Right(astr, Len(astr) - 5)
        While Len(Trim(astr)) > 1
            bstr = Left(astr, 81)
            If Len(astr) > 81 Then astr = Right(astr, Len(astr) - 81) Else astr = ""
            If Trim(bstr) <> "" Then
                Control.AddItem (Right(bstr, Len(bstr) - 1))
                If IsNumeric(Left(bstr, 1)) Then
                    If CInt(Left(bstr, 1)) = 0 Then
                        If ForcedPageCount = ForcedPageNum Then
                            ForcedPageCount = ForcedPageCount + 10
                            ReDim Preserve ForcedPageLines(ForcedPageCount)
                        End If
                        ForcedPageNum = ForcedPageNum + 1
                        ForcedPageLines(ForcedPageNum) = Control.ListCount
                    Else
                        For k = 1 To CInt(Left(bstr, 1)) - 1
                            Control.AddItem ("")
                        Next k
                    End If
                End If
            End If
        Wend
    End If
Next i

End Function

Public Property Get Visible(Optional processphase As Integer) As Boolean
    Visible = DisplayFlag
End Property

Public Property Get ItemCount() As Integer
    ItemCount = Control.ListCount
End Property

Public Property Get item(inItemNo) As String
    item = Control.list(inItemNo - 1)
End Property

Public Sub RemoveItem(inItemNo)
    Control.RemoveItem inItemNo - 1
End Sub

Public Property Let Visible(Optional processphase As Integer, aFlag As Boolean)
    DisplayFlag = aFlag
    PropertyChanged "Visible"
End Property

Public Function IsVisible(Optional processphase As Integer) As Boolean
    IsVisible = DisplayFlag
End Function

Private Sub Control_DblClick()
    Dim aFlag As Boolean
    aFlag = False
    If Not (owner Is Nothing) Then
        With owner
            .HandleEvent LstName, Control.ListIndex, 0, aFlag
        End With
     End If
End Sub

Private Sub Control_GotFocus()
    Set owner.ActiveListBox = Me
End Sub

Private Sub UserControl_Initialize()
    Control.Clear
End Sub

Public Sub AddItem(astr)
    Control.AddItem astr
End Sub

Public Property Get Selected(Index) As Boolean
    Selected = Control.Selected(Index)
End Property

Public Sub Clear()
    Control.Clear
End Sub

Private Sub UserControl_Resize()
   Control.Left = 0
   Control.Top = 0
   Control.width = width
   Control.height = height
End Sub

Public Sub FinalizeEdit()

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

Public Property Get PageHead() As String
    PageHead = DocPageHead
End Property

Public Property Let PageHead(value As String)
    DocPageHead = value

Dim bstr As String, aLine As String
    bstr = DocPageHead
    DocHeadLines = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        DocHeadLines = DocHeadLines + 1
    Loop
    DocClearLines = DocLines - DocTopMargin - DocBottomMargin - DocHeadLines - DocFootLines

End Property

Public Property Get PageFoot() As String
    PageFoot = DocPageFoot
End Property

Public Property Let PageFoot(value As String)
    DocPageFoot = value

Dim bstr As String, aLine As String
    bstr = DocPageFoot
    DocFootLines = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        DocFootLines = DocFootLines + 1
    Loop
    DocClearLines = DocLines - DocTopMargin - DocBottomMargin - DocHeadLines - DocFootLines

End Property

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML2.IXMLDOMElement)
    
    Dim astr As String
    
    Set owner = inOwner
    LstNo = NodeIntegerFld(inNode, "LstNo", listModel)
    LstName = "Lst" & StrPad_(CStr(LstNo), 3, "0", "L")
    
    LstName2 = UCase(NodeStringFld(inNode, "NAME", listModel))
    If InStr(ReservedControlPrefixes, "," & Left(LstName2, 3) & ",") > 0 Then LstName2 = ""
    
    name = IIf(LstName2 <> "", LstName2, LstName)

    LabelName = "LLabel" & StrPad_(CStr(LstNo), 3, "0", "L")
    Set ValidationControl = inProcessControl
    ValidationControl.AddObject LstName, Me, True
    
    If Trim(LstName2) <> "" And UCase(Trim(LstName2)) <> UCase(Trim(LstName)) Then
        On Error GoTo LstRegistrationError
        ValidationControl.ExecuteStatement "Set " & LstName2 & "=" & LstName
        GoTo LstRegistrationOk
LstRegistrationError:
        MsgBox "Λάθος κατα τη δήλωση του πεδίου: " & LstName & ":" & LstName2
LstRegistrationOk:
    End If
    
    ScrLeft = NodeIntegerFld(inNode, "ScrX", listModel)
    ScrWidth = NodeIntegerFld(inNode, "ScrWidth", listModel)
    ScrTop = NodeIntegerFld(inNode, "ScrY", listModel) * 290
    ScrHeight = NodeIntegerFld(inNode, "ScrHeight", listModel) * 285
    
    LabelName = "LLabel" & NodeStringFld(inNode, "LstNo", listModel)
    
    Set Prompt = parent.Controls.add("Vb.Label", LabelName)
    Prompt.BackColor = parent.BackColor
    Prompt.AutoSize = False
    Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", listModel)
        
    parent.ScaleMode = vbCharacters
    Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", listModel)
    Prompt.width = NodeIntegerFld(inNode, "ScrPromptWidth", listModel)
    parent.ScaleMode = vbTwips
    Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", listModel) * 290
    Prompt.height = NodeIntegerFld(inNode, "ScrPromptHeight", listModel) * 285
    DisplayFlag = NodeBooleanFld(inNode, "ScrDisplay", listModel)
    
    TTabIndex = NodeIntegerFld(inNode, "TabIndex", listModel)
    
    GetLinesFromBuffer = NodeBooleanFld(inNode, "GetLinesFromBuffer", listModel)
    MinInLineLength = NodeIntegerFld(inNode, "MinInLineLength", listModel)
    FromColumn = NodeIntegerFld(inNode, "FromColumn", listModel)
    ToColumn = NodeIntegerFld(inNode, "ToColumn", listModel)
    ExcludeFirstLines = NodeIntegerFld(inNode, "ExcludeFirstLines", listModel)
    ExcludeLastLines = NodeIntegerFld(inNode, "ExcludeLastLines", listModel)
    
    DocTopMargin = NodeIntegerFld(inNode, "DocTopMargin", listModel)
    DocLeftMargin = NodeIntegerFld(inNode, "DocLeftMargin", listModel)
    DocRightMargin = NodeIntegerFld(inNode, "DocRightMargin", listModel)
    DocBottomMargin = NodeIntegerFld(inNode, "DocBottomMargin", listModel)
    DocOrientation = NodeIntegerFld(inNode, "DocOrientation", listModel)
    DocLines = NodeIntegerFld(inNode, "DocLines", listModel)
    If DocLines > DocumentLines Then DocLines = DocumentLines
    DocColumns = NodeIntegerFld(inNode, "DocColumns", listModel)
    
    astr = NodeStringFld(inNode, "ExamineLineScript", listModel)
If astr <> "" Then
    ExamineLineFlag = True
    ValidationControl.AddCode "Public Sub " & LstName & "_ExamineLine_Script " & vbCrLf & _
        astr & vbCrLf & "End Sub"
End If
    
    PageHead = NodeStringFld(inNode, "DocPageHead", listModel)
    PageFoot = NodeStringFld(inNode, "DocPageFoot", listModel)
     
End Sub

Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("LISTBOX")
    Set attr = XML.createAttribute("NO")
    attr.nodeValue = UCase(Me.LstNo)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("NAME")
    attr.nodeValue = UCase(Me.LstName)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("VISIBLE")
    attr.nodeValue = UCase(Me.Visible)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("ITEMCOUNT")
    attr.nodeValue = UCase(Me.ItemCount)
    elm.setAttributeNode attr

    Dim i As Integer
    For i = 1 To Me.ItemCount
        Set attr = XML.createAttribute("ITEM" + CStr(i))
        attr.nodeValue = UCase(Me.item(i))
        elm.setAttributeNode attr
    Next
    

    Set TranslateToProperties = elm
End Function

Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    If inPhase = 0 Then inPhase = 1
    Me.LstName2 = elm.getAttributeNode("FULLNAME").nodeValue
    Me.Visible = elm.getAttributeNode("VISIBLE").nodeValue
    Dim iCount As Integer
    iCount = elm.getAttributeNode("ITEMCOUNT").nodeValue
    Me.Clear
    Dim i As Integer
    For i = 1 To iCount
        
        Me.AddItem (elm.getAttributeNode("ITEM" + CStr(i)).nodeValue)
        
    Next
    
End Sub


