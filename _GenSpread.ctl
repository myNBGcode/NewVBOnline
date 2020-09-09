VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.UserControl GenSpread 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Control 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
      FixedRows       =   2
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New Greek"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "GenSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As Form
Public SprdNo As Integer, SprdName As String, LabelName As String
Public Prompt As Label
Public ScrLeft As Integer, ScrTop As Integer, ScrWidth As Integer, ScrHeight As Integer
Private DisplayFlag As Boolean
Private ValidationControl As ScriptControl

Private GetLinesFromBuffer As Boolean
Private MinInLineLength As Integer, FromColumn As Integer, ToColumn As Integer
Private ExcludeFirstLines As Integer, ExcludeLastLines As Integer

Private DocTopMargin, DocLeftMargin, DocRightMargin, DocBottomMargin As Integer
Private DocOrientation As Integer
Private DocPageHead, DocPageFoot As String
Private DocHeadLines, DocFootLines, DocClearLines As Integer
Private DocLines, DocColumns As Integer

'Private FormatString As String
Dim ColumnPairs(30) As IntegerPair
Dim PairsCount As Integer
Private ExamineLineFlag As Boolean
Private curPageNo As Integer
Public TTabIndex As Integer

Public Sub UnlockPrinter(owner As Form)
    If cPassbookPrinter = 5 Then
        owner.SPCPanel.UnlockPrinter
    ElseIf cPassbookPrinter = 1 Or cPassbookPrinter = 2 Then
        docPrinter.WUnlock owner
    End If
End Sub

Public Sub PrintPageHead()
If cListToPassbook = 1 Then PrintPageHead_P: Exit Sub
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
        aLine = String(DocLeftMargin, " ") & aLine
        aLineNo = aLineNo + 1
        xSetDocLine_ aLineNo, aLine
    Loop
End Sub

Public Sub PrintPageFoot()
If cListToPassbook = 1 Then PrintPageFoot_P: Exit Sub
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
    aLineNo = 0
    
    aLineNo = DocClearLines + DocTopMargin + DocHeadLines
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        aLine = EmbedValues(owner, aLine, curPageNo)
        aLine = String(DocLeftMargin, " ") & aLine
        aLineNo = aLineNo + 1
        xSetDocLine_ aLineNo, aLine
    Loop
End Sub

Public Function GetPrintLine(inLine, formcharwidth As Integer) As String
Dim astr As String, OutStr As String, aWidth As Integer, i, adiff As Integer

OutStr = ""
For i = 0 To Control.Cols - 1
    aWidth = (Control.ColWidth(i) / formcharwidth) - 1
    astr = Trim(Control.TextMatrix(inLine, i))
    Select Case Control.ColAlignment(i)
        Case flexAlignLeftTop, flexAlignLeftCenter, flexAlignLeftBottom, flexAlignGeneral
            If Len(astr) <= aWidth Then
                astr = StrPad_(astr, aWidth, " ", "R")
            Else
                astr = StrPad_("", aWidth, "*", "R")
            End If
        Case flexAlignRightTop, flexAlignRightCenter, flexAlignRightBottom
            If Len(astr) <= aWidth Then
                astr = StrPad_(astr, aWidth, " ", "L")
            Else
                astr = StrPad_("", aWidth, "*", "L")
            End If
        Case flexAlignCenterTop, flexAlignCenterCenter, flexAlignCenterBottom
            If Len(astr) <= aWidth Then
                adiff = (aWidth - Len(astr)) \ 2
                astr = StrPad_(astr, adiff + Len(astr), " ", "L")
                astr = StrPad_(astr, aWidth, " ", "R")
            Else
                astr = StrPad_("", aWidth, "*", "L")
            End If
    End Select
    
    If OutStr <> "" Then
        OutStr = OutStr & " " & astr
    Else
        OutStr = astr
    End If
    
Next i
GetPrintLine = OutStr
End Function

Public Sub PrintLines()

eJournalWrite owner, "ÅÊÔÕÐÙÓÇ ÓÔÏÉ×ÅÉÙÍ"
eJournalWrite owner, "Ô ÙÑÁ ÓÕÍÁËËÁÃÇÓ: " & FormatDateTime(Time, vbShortTime)

If cListToPassbook = 1 Then PrintLines_P: Exit Sub
If DocOrientation = 1 Then
    Printer.Orientation = vbPRORPortrait
Else
    Printer.Orientation = vbPRORLandscape
End If
Printer.ScaleMode = vbCharacters
Printer.FontName = Control.FontName
Printer.FontSize = Control.FontSize
Printer.FontBold = Control.FontBold

Dim aFontName As String, aFontSize As Integer, aFontBold As Boolean
aFontName = owner.FontName
aFontSize = owner.FontSize
aFontBold = owner.FontBold

owner.FontName = Control.FontName
owner.FontSize = Control.FontSize
owner.FontBold = Control.FontBold

owner.ScaleMode = vbTwips
Dim aWidth, formcharwidth As Integer
    aWidth = TextWidth("ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ")
    formcharwidth = aWidth / 24

owner.FontName = aFontName
owner.FontSize = aFontSize
owner.FontBold = aFontBold
    
Dim aLine, i As Integer
Dim astr As String
    
    curPageNo = 1
    PrintPageHead
    For i = 0 To Control.Rows - 1

        astr = GetPrintLine(i, formcharwidth)
        If Len(astr) > DocColumns - DocLeftMargin - DocRightMargin Then
            astr = Left(astr, DocColumns - DocLeftMargin - DocRightMargin)
        End If

        Printer.CurrentX = DocLeftMargin

        aLine = (i + 1) Mod DocClearLines

        If aLine > 0 Then Printer.CurrentY = aLine + DocTopMargin + DocHeadLines
        If aLine = 0 Then Printer.CurrentY = DocClearLines + DocTopMargin

        Printer.Print astr
        
        If aLine = 0 Then PrintPageFoot
        If aLine = 0 And (i < Control.Rows - 1) Then
            Printer.NewPage
            curPageNo = curPageNo + 1
            PrintPageHead
        End If

    Next i
    Printer.EndDoc
End Sub


Public Sub PrintLines_P(Optional UnlockAfterPrint As Boolean)
If IsMissing(UnlockAfterPrint) Then UnlockAfterPrint = False

Dim aFontName As String, aFontSize As Integer, aFontBold As Boolean
aFontName = owner.FontName
aFontSize = owner.FontSize
aFontBold = owner.FontBold

owner.FontName = Control.FontName
owner.FontSize = Control.FontSize
owner.FontBold = Control.FontBold

owner.ScaleMode = vbTwips
Dim aWidth, formcharwidth As Integer
    aWidth = TextWidth("ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ")
    formcharwidth = aWidth / 24

owner.FontName = aFontName
owner.FontSize = aFontSize
owner.FontBold = aFontBold
    
Dim aLine As Integer, bLine As Integer, i As Integer
Dim astr As String
    
    xClearDoc_
    curPageNo = 1
    PrintPageHead_P
    For i = 0 To Control.Rows - 1

        astr = GetPrintLine(i, formcharwidth)
        If Len(astr) > DocColumns - DocLeftMargin - DocRightMargin Then
            astr = Left(astr, DocColumns - DocLeftMargin - DocRightMargin)
        End If
        astr = String(DocLeftMargin, " ") & astr

        aLine = (i + 1) Mod DocClearLines

        If aLine > 0 Then bLine = aLine + DocTopMargin + DocHeadLines
        If aLine = 0 Then bLine = DocClearLines + DocTopMargin

        xSetDocLine_ bLine, astr
        
        If aLine = 0 Then PrintPageFoot_P
        If aLine = 0 And (i < Control.Rows - 1) Then
            xPrintDoc_ owner
            curPageNo = curPageNo + 1
            xClearDoc_
            PrintPageHead_P
        End If

    Next i
    If aLine <> 0 Then PrintPageFoot_P
    
    xPrintDoc_ owner
    If UnlockAfterPrint Then docPrinter.WUnlock owner
End Sub

Public Function ReadLines() As Boolean
ReadLines = False
Dim i As Integer, k As Integer, cr As Integer
Dim astr As String
If GetLinesFromBuffer Then
    Control.clear
    Control.Cols = PairsCount
    If FormatString <> "" Then
        Control.Rows = ReceivedData.count + 1
        Control.FixedRows = 1
        Control.FormatString = FormatString
        For i = 0 To Control.Cols - 1
            Control.ColAlignmentHeader(i) = Control.ColAlignment(i)
        Next i
        
        cr = 1
    Else
        Control.Rows = ReceivedData.count
        Control.FixedRows = 0
        cr = 0
    End If
    
    For i = 1 To ReceivedData.count
        
        astr = ReceivedData(i)
        If (i > ExcludeFirstLines) And (i <= ReceivedData.count - ExcludeLastLines) Then
        
        If (Len(astr) > MinInLineLength) Then
            
            For k = 1 To PairsCount
                Control.TextMatrix(cr, k - 1) = Trim(Mid(astr, ColumnPairs(k).p1, ColumnPairs(k).p2 - ColumnPairs(k).p1 + 1))
            Next k
            cr = cr + 1
'            If FromColumn > 0 And FromColumn <= Len(astr) Then
'                If ToColumn > FromColumn Then
'                    astr = Mid(astr, FromColumn, ToColumn - FromColumn + 1)
'                Else
'                    astr = Right(astr, Len(astr) - FromColumn + 1)
'                End If
'            ElseIf FromColumn <= Len(astr) Then
'                If ToColumn > 0 Then
'                    astr = Left(astr, ToColumn)
'                End If
'            ElseIf FromColumn > Len(astr) Then
'                astr = ""
'            End If
'            Control.AddItem (astr)
        End If
        End If
    Next i
    Control.Rows = cr + 1
End If
ReadLines = True
End Function

Public Function IsVisible() As Boolean
    IsVisible = DisplayFlag
End Function

Public Sub ReadFromStruct(ByRef inPart As BufferPart)
Dim i As Long, k As Long
    Rows = inPart.Times + 1
    FixedRows = 1
    Cols = inPart.Struct(1).PartNum + 1
    For i = 1 To inPart.Times
        For k = 1 To inPart.Struct(i).PartNum
            Control.TextMatrix(i, k) = CStr(inPart.ByIndex(k, i).value)
        Next k
    Next i
End Sub

Public Property Get TextMatrix(Row As Integer, Col As Integer) As String
    TextMatrix = Control.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(Row As Integer, Col As Integer, aValue As String)
    Control.TextMatrix(Row, Col) = aValue
    PropertyChanged "TextMatrix"
End Property

Public Function GetCellText(aRow, aCol) As String
    GetCellText = Control.TextMatrix(aRow, aCol)
End Function

Public Sub SetCellText(aRow, aCol, aValue)
    Control.TextMatrix(aRow, aCol) = aValue
End Sub

Public Sub ClearLines()
    Dim i As Integer, k As Integer
    For i = Control.FixedRows To Control.Rows - 1
        For k = Control.FixedCols To Control.Cols - 1
            Control.TextMatrix(i, k) = ""
        Next
    Next
End Sub
Public Property Get FormatString() As String
    FormatString = Control.FormatString
End Property
Public Property Let FormatString(informatstring As String)
'    FormatString = informatstring
    Control.FormatString = informatstring
    
    PropertyChanged "FormatString"

End Property
Public Property Get DataSource() As ADODB.Recordset
    Set DataSource = Control.DataSource
End Property

Public Property Let DataSource(inDataSource)
    If Not (Control.DataSource Is Nothing) Then
        'Control.DataSource.Close
        Set Control.DataSource = Nothing
    End If
    Set Control.DataSource = inDataSource
    PropertyChanged "DataSource"
End Property

Public Property Let DataSourceAsText(inCmd)
    Set Control.DataSource = owner.GetADORecordset(inCmd)
    PropertyChanged "DataSource"
End Property

Public Property Get Rows() As Integer
    Rows = Control.Rows
End Property

Public Property Get FixedRows() As Integer
    FixedRows = Control.FixedRows
End Property

Public Property Let Rows(aValue As Integer)
    Control.Rows = aValue
    PropertyChanged "Rows"
End Property

Public Property Let FixedRows(aValue As Integer)
    Control.FixedRows = aValue
    PropertyChanged "FixedRows"
End Property

Public Property Get Cols() As Integer
    Cols = Control.Cols
End Property

Public Property Get FixedCols() As Integer
    FixedCols = Control.FixedCols
End Property

Public Property Let Cols(aValue As Integer)
    Control.Cols = aValue
    PropertyChanged "Cols"
End Property

Public Property Let FixedCols(aValue As Integer)
    Control.FixedCols = aValue
    PropertyChanged "FixedCols"
End Property

Public Property Get Text() As String
    Text = Control.Text
End Property

Public Property Get Col() As Integer
    Col = Control.Col
End Property

Public Property Get Row() As Integer
    Row = Control.Row
End Property

Public Property Get TextArray(cellindex As Integer) As String
    TextArray = Control.TextArray(cellindex)
End Property

'Public Property Get TextMatrix(rowindex As Integer, colindex As Integer) As String
'    TextMatrix = Control.TextMatrix(rowindex, colindex)
'End Property

Public Property Let Text(aValue As String)
    Control.Text = aValue
    PropertyChanged "Text"
End Property

Public Property Let Col(aValue As Integer)
    Control.Col = aValue
    PropertyChanged "Col"
End Property

Public Property Let Row(aValue As Integer)
    Control.Row = aValue
    PropertyChanged "Row"
End Property

Public Property Let TextArray(cellindex As Integer, aValue As String)
    Control.TextArray(cellindex) = aValue
    PropertyChanged "TextArray"
End Property

Private Sub Control_GotFocus()
    Set owner.ActiveSpread = Me
End Sub

'Public Property Let TextMatrix(rowindex As Integer, colindex As Integer, aValue As String)
'    Control.TextMatrix(rowindex, colindex) = aValue
'    PropertyChanged "TextMatrix"
'End Property

Private Sub UserControl_Resize()
Control.Left = 0
Control.Top = 0
Control.Width = Width
Control.Height = Height
End Sub


Private Sub ConvertToPairs(inputStr As String)
Dim V1 As Integer, V2 As Integer

Dim i As Integer, s As Integer

i = 0
Dim astr As String
astr = Trim(inputStr)

While Len(astr) > 0
    V1 = -1
    V2 = -1
    s = InStr(astr, ",")
    On Error GoTo invalid_v1
    If s > 0 Then
        V1 = Val(Left(astr, s - 1))
        astr = Trim(Right(astr, Len(astr) - s))
        s = InStr(astr, ",")
        On Error GoTo invalid_v2
        If s > 0 Then
            V2 = Val(Left(astr, s - 1))
            astr = Trim(Right(astr, Len(astr) - s))
        Else
            V2 = Val(astr)
            astr = ""
        End If
    Else
        V1 = Val(astr)
        astr = ""
    End If
    GoTo before_exit
    
invalid_v1:
        V1 = -1
        astr = Trim(Right(astr, Len(astr) - s))
        GoTo before_exit
invalid_v2:
        V2 = -1
        astr = Trim(Right(astr, Len(astr) - s))
        GoTo before_exit
before_exit:
    
    If V1 <= V2 And V1 > 0 And V2 > 0 Then
        i = i + 1
        If i <= UBound(ColumnPairs) Then
            ColumnPairs(i).p1 = V1
            ColumnPairs(i).p2 = V2
        End If
    End If
Wend
If i <= UBound(ColumnPairs) Then
    PairsCount = i
Else
    PairsCount = UBound(ColumnPairs)
End If
End Sub

Public Sub FinalizeEdit()

End Sub

Public Sub InitializeFromXML(inOwner As Form, inProcessControl As ScriptControl, _
    inNode As MSXML.IXMLElement)
    
    Dim astr As String
    
    Set owner = inOwner
    Set Font = Control.Font
    SprdNo = NodeIntegerFld(inNode, "SprdNo", gridModel)
    SprdName = "Spd" & StrPad_(CStr(SprdNo), 3, "0", "L") 'NodeStringFld(inNode, "Name", gridModel)
    
    Set ValidationControl = inProcessControl
    ValidationControl.AddObject SprdName, Me, True
    
    ScrLeft = NodeIntegerFld(inNode, "ScrX", gridModel)
    ScrWidth = NodeIntegerFld(inNode, "ScrWidth", gridModel)
    ScrTop = NodeIntegerFld(inNode, "ScrY", gridModel) * 290
    ScrHeight = NodeIntegerFld(inNode, "ScrHeight", gridModel) * 285
    
    LabelName = "GLabel" & StrPad_(CStr(SprdNo), 3, "0", "L")
        
    Set Prompt = Parent.Controls.Add("Vb.Label", LabelName)
    Prompt.BackColor = Parent.BackColor
    Prompt.AutoSize = False
    Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", gridModel)
        
    Parent.ScaleMode = vbCharacters
    Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", gridModel)
    Prompt.Width = NodeIntegerFld(inNode, "ScrPromptWidth", gridModel)
    Parent.ScaleMode = vbTwips
    Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", gridModel) * 290
    Prompt.Height = NodeIntegerFld(inNode, "ScrPromptHeight", gridModel) * 285
    DisplayFlag = NodeBooleanFld(inNode, "ScrDisplay", gridModel)
    
    TTabIndex = NodeIntegerFld(inNode, "TabIndex", gridModel)
    
    
    FormatString = NodeStringFld(inNode, "FormatString", gridModel)
    Control.FormatString = FormatString
    Control.Rows = NodeIntegerFld(inNode, "Rows", gridModel)
    Control.Cols = NodeIntegerFld(inNode, "Cols", gridModel)
    Control.FixedRows = NodeIntegerFld(inNode, "FixedRows", gridModel)
    Control.FixedCols = NodeIntegerFld(inNode, "FixedCols", gridModel)
    
    Control.AllowUserResizing = flexResizeColumns

    GetLinesFromBuffer = NodeBooleanFld(inNode, "GetLinesFromBuffer", gridModel)
    MinInLineLength = NodeIntegerFld(inNode, "MinInLineLength", gridModel)
    FromColumn = NodeIntegerFld(inNode, "FromColumn", gridModel)
    ToColumn = NodeIntegerFld(inNode, "ToColumn", gridModel)
    ExcludeFirstLines = NodeIntegerFld(inNode, "ExcludeFirstLines", gridModel)
    ExcludeLastLines = NodeIntegerFld(inNode, "ExcludeLastLines", gridModel)

    
    DocTopMargin = NodeIntegerFld(inNode, "DocTopMargin", gridModel)
    DocLeftMargin = NodeIntegerFld(inNode, "DocLeftMargin", gridModel)
    DocRightMargin = NodeIntegerFld(inNode, "DocRightMargin", gridModel)
    DocBottomMargin = NodeIntegerFld(inNode, "DocBottomMargin", gridModel)
    DocOrientation = NodeIntegerFld(inNode, "DocOrientation", gridModel)
    DocPageHead = NodeStringFld(inNode, "DocPageHead", gridModel)
    DocPageFoot = NodeStringFld(inNode, "DocPageFoot", gridModel)
    DocLines = NodeIntegerFld(inNode, "DocLines", gridModel)
    If DocLines > DocumentLines Then DocLines = DocumentLines
    DocColumns = NodeIntegerFld(inNode, "DocColumns", gridModel)
    
    astr = NodeStringFld(inNode, "ExamineLineScript", gridModel)
If astr <> "" Then
    ExamineLineFlag = True
    ValidationControl.AddCode "Public Sub " & SprdName & "_ExamineLine_Script " & vbCrLf & _
        astr & vbCrLf & "End Sub"
End If
Dim inputStr As String
    inputStr = NodeStringFld(inNode, "InputColumns", gridModel)
    ConvertToPairs (inputStr)
    
Dim bstr As String, aLine As String
    bstr = DocPageHead
    DocHeadLines = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        DocHeadLines = DocHeadLines + 1
    Loop
    bstr = DocPageFoot
    DocFootLines = 0
    Do Until bstr = ""
        aLine = ExtractFirstLineFromString(bstr)
        DocFootLines = DocFootLines + 1
    Loop
    DocClearLines = DocLines - DocTopMargin - DocBottomMargin - DocHeadLines - DocFootLines
    
End Sub


