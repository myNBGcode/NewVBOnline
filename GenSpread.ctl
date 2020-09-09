VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl GenSpread 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox EditText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "GenSpread.ctx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
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
      BackColorFixed  =   12632256
      ForeColorSel    =   -2147483637
      BackColorBkg    =   12632256
      BackColorUnpopulated=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New Greek"
         Size            =   9.75
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
Public SprdNo As Integer, SprdName As String, SprdName2 As String, name As String, LabelName As String
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
Private ERowsCount As Long
Public TTabIndex As Long
Public PrintTruncatedCells As Boolean
Public EnableLaserPrinter As Boolean
Private PrintTarget As Integer
Public AllowLowerCase As Boolean

Public Property Let Enabled(value As Boolean)
    Control.Enabled = value
End Property

Public Property Get Enabled() As Boolean
    Enabled = Control.Enabled
End Property

Public Sub UnLockPrinter(owner As Form)
    If cPassbookPrinter = 5 Or cPassbookPrinter = 6 Or cPassbookPrinter = 7 Then
        owner.SPCPanel.UnLockPrinter
    End If
End Sub

Public Sub PrintPageHead()
If cListToPassbook = 1 And PrintTarget = 0 Then PrintPageHead_P: Exit Sub
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
If cListToPassbook = 1 And PrintTarget = 0 Then PrintPageFoot_P: Exit Sub
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

Public Function GetJournalLine(inLine As Integer) As String
Dim astr As String, OutStr As String, aWidth As Integer, i, adiff As Integer
OutStr = ""
For i = 0 To Control.Cols - 1
    OutStr = OutStr & "[" & Trim(Control.TextMatrix(inLine, i)) & "]"
Next i
GetJournalLine = OutStr
End Function

Public Sub PrintToJournal()
Dim i As Integer, astr As String, bstr As String
If Control.Cols * Control.Rows = 0 Then Exit Sub
For i = 0 To Control.Rows - 1
    astr = GetJournalLine(i): bstr = astr
    bstr = Replace(bstr, "[", "")
    bstr = Replace(bstr, "]", "")
    If Trim(bstr) <> "" Then eJournalWrite astr
Next i
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
                astr = Left(astr, aWidth)
'                If PrintTruncatedCells Then astr = Left(astr, aWidth) Else astr = StrPad_("", aWidth, "*", "R")
            End If
        Case flexAlignRightTop, flexAlignRightCenter, flexAlignRightBottom
            If Len(astr) <= aWidth Then
                astr = StrPad_(astr, aWidth, " ", "L")
            Else
                astr = Right(astr, aWidth)
'                If PrintTruncatedCells Then astr = Right(astr, aWidth) Else astr = StrPad_("", aWidth, "*", "L")
            End If
        Case flexAlignCenterTop, flexAlignCenterCenter, flexAlignCenterBottom
            If Len(astr) <= aWidth Then
                adiff = (aWidth - Len(astr)) \ 2
                astr = StrPad_(astr, adiff + Len(astr), " ", "L")
                astr = StrPad_(astr, aWidth, " ", "R")
            Else
                astr = Left(astr, aWidth)
'                If PrintTruncatedCells Then astr = Left(astr, aWidth) Else astr = StrPad_("", aWidth, "*", "L")
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

eJournalWrite "ÅÊÔÕÐÙÓÇ ÓÔÏÉ×ÅÉÙÍ"
eJournalWrite "Ô ÙÑÁ ÓÕÍÁËËÁÃÇÓ: " & FormatDateTime(Time, vbShortTime)

If cListToPassbook = 1 And Not EnableLaserPrinter Then PrintTarget = 0: PrintLines_P: Exit Sub
Dim aFrm As SelectPrinterFrm
Set aFrm = New SelectPrinterFrm
Load aFrm
Dim res As Long
aFrm.Show vbModal, owner

Dim aPrinterName As String
aPrinterName = aFrm.SelectedPrinter

If UCase(aPrinterName) = "PASSBOOK" Then PrintTarget = 0: PrintLines_P: Exit Sub
PrintTarget = 1
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


'If DocOrientation = 1 Then
'    Printer.Orientation = vbPRORPortrait
'Else
'    Printer.Orientation = vbPRORLandscape
'End If
'Printer.ScaleMode = vbCharacters
'Printer.FontName = Control.FontName
'Printer.FontSize = Control.FontSize
'Printer.FontBold = Control.FontBold
'
'Dim aFontName As String, aFontSize As Integer, aFontBold As Boolean
'aFontName = owner.FontName
'aFontSize = owner.FontSize
'aFontBold = owner.FontBold
'
'owner.FontName = Control.FontName
'owner.FontSize = Control.FontSize
'owner.FontBold = Control.FontBold
'
'owner.ScaleMode = vbTwips
Dim aWidth, formcharwidth As Integer
    aWidth = TextWidth("ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ")
    formcharwidth = aWidth / 24
'
'owner.FontName = aFontName
'owner.FontSize = aFontSize
'owner.FontBold = aFontBold
    
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

        If aLine > 0 Then
            Printer.CurrentY = aLine + DocTopMargin + DocHeadLines
        End If
        If aLine = 0 Then
            Printer.CurrentY = DocClearLines + DocTopMargin + DocHeadLines
        End If

        Printer.Print astr
        
        If aLine = 0 Then
            PrintPageFoot
        End If
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
    If DocClearLines = 0 Then MsgBox "Äåí Ý÷åé ïñéóôåß ðåñéï÷Þ åêôýðùóçò", vbCritical, "ËÁÈÏÓ": Exit Sub
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
End Sub

Public Function ReadLines() As Boolean
ReadLines = False
Dim i As Integer, k As Integer, cr As Integer
Dim astr As String
If GetLinesFromBuffer Then
    Control.Clear
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
        End If
        End If
    Next i
    Control.Rows = cr + 1
End If
ReadLines = True
End Function

Public Function IsVisible(Optional processphase As Integer) As Boolean
    IsVisible = DisplayFlag
End Function

Public Sub SetDisplay(SetFlag)
    DisplayFlag = CBool(SetFlag)
End Sub

Public Sub ReadFromStruct(ByRef inPart As BufferPart, Optional KeyPart As String, Optional VisibleParts As String)
Dim i As Long, k As Long, l As Long, rowsFound As Long, astr As String, apart As String
Dim aList As Collection
    Control.Rows = inPart.times + 1
    Control.FixedRows = 1

    If Not IsMissing(VisibleParts) Then astr = VisibleParts Else astr = ""
    If astr <> "" Then
        Set aList = New Collection
        i = InStr(astr, ",")
        While i > 0
            apart = Left(astr, i - 1)
            If i < Len(astr) Then astr = Right(astr, Len(astr) - i) Else astr = ""
            If Trim(apart) <> "" Then aList.add Trim(apart)
            i = InStr(astr, ",")
        Wend
        Control.Cols = aList.count
    Else
        Control.Cols = inPart.Struct(1).PartNum + 1
    End If
    
    rowsFound = 0
    For i = 1 To inPart.times
        If Not IsMissing(KeyPart) And KeyPart <> "" Then
            If inPart.ByName(KeyPart, i).value = 0 Or Trim(inPart.ByName(KeyPart, i).value) = "" Then Exit For
        End If
        rowsFound = rowsFound + 1
        
        l = 0
        If IsMissing(VisibleParts) Or VisibleParts = "" Then
            For k = 1 To inPart.Struct(i).PartNum
                If Not IsMissing(VisibleParts) And VisibleParts <> "" Then
                    If InStr(UCase(VisibleParts), UCase(inPart.ByIndex(k, i).name) & ",") > 0 Then
                        Control.TextMatrix(0, l) = inPart.ByIndex(k, i).name
                        Control.TextMatrix(i, l) = IIf(inPart.ByIndex(k, i).datatype = ptDate, inPart.ByIndex(k, i).FormatedDate8, CStr(inPart.ByIndex(k, i).value))
                        l = l + 1
                    End If
                Else
                    Control.TextMatrix(0, l) = inPart.ByIndex(k, i).name
                    Control.TextMatrix(i, l) = IIf(inPart.ByIndex(k, i).datatype = ptDate, inPart.ByIndex(k, i).FormatedDate8, CStr(inPart.ByIndex(k, i).value))
                    l = l + 1
                End If
            Next k
        Else
            For k = 1 To aList.count
                astr = inPart.ByName(aList(k), i).name
                Control.TextMatrix(0, k - 1) = Trim(astr)
                astr = IIf(inPart.ByName(aList(k), i).datatype = ptDate, inPart.ByName(aList(k), i).FormatedDate8, inPart.ByName(aList(k), i).value)
                Control.TextMatrix(i, k - 1) = Trim(astr)
            Next k
        End If
    Next i
    If rowsFound > 0 Then Control.Rows = rowsFound + 1 Else Control.Rows = 0
    If Not (aList Is Nothing) Then
        For i = aList.count To 1 Step -1
            aList.Remove i
        Next i
        Set aList = Nothing
    End If
End Sub

Public Sub ReadFromStructVertical(ByRef inPart As BufferPart, Optional KeyPart As String, Optional VisibleParts As String)
Dim i As Long, k As Long, l As Long, colsFound As Long, astr As String, apart As String, maxLen As Integer, FormatString As String
Dim aList As Collection
    Control.Cols = inPart.times + 1
    Control.FixedCols = 1: Control.FixedRows = 1

    If Not IsMissing(VisibleParts) Then astr = VisibleParts Else astr = ""
    If astr <> "" Then
        Set aList = New Collection
        i = InStr(astr, ",")
        While i > 0
            apart = Left(astr, i - 1)
            If i < Len(astr) Then astr = Right(astr, Len(astr) - i) Else astr = ""
            If Trim(apart) <> "" Then aList.add Trim(apart)
            i = InStr(astr, ",")
        Wend
        Control.Rows = aList.count + 1
    Else
        Control.Rows = inPart.Struct(1).PartNum + 1
    End If
    
    colsFound = 0: maxLen = 0
    For i = 1 To inPart.times
        If Not IsMissing(KeyPart) And KeyPart <> "" Then
            If inPart.ByName(KeyPart, i).value = 0 Then Exit For
        End If
        colsFound = colsFound + 1
        
        l = 1
        If IsMissing(VisibleParts) Or VisibleParts = "" Then
            For k = 1 To inPart.Struct(i).PartNum
                If Not IsMissing(VisibleParts) And VisibleParts <> "" Then
                    If InStr(UCase(VisibleParts), UCase(inPart.ByIndex(k, i).name) & ",") > 0 Then
                        Control.TextMatrix(l, 0) = inPart.ByIndex(k, i).name
                        Control.TextMatrix(l, i) = IIf(inPart.ByIndex(k, i).datatype = ptDate, inPart.ByIndex(k, i).FormatedDate8, CStr(inPart.ByIndex(k, i).value))
                        If maxLen < Len(CStr(inPart.ByIndex(k, i).value)) Then maxLen = Len(CStr(inPart.ByIndex(k, i).value))
                        l = l + 1
                    End If
                Else
                    Control.TextMatrix(l, 0) = inPart.ByIndex(k, i).name
                    Control.TextMatrix(l, i) = IIf(inPart.ByIndex(k, i).datatype = ptDate, inPart.ByIndex(k, i).FormatedDate8, CStr(inPart.ByIndex(k, i).value))
                    If maxLen < Len(CStr(inPart.ByIndex(k, i).value)) Then maxLen = Len(CStr(inPart.ByIndex(k, i).value))
                    l = l + 1
                End If
            Next k
        Else
            For k = 1 To aList.count
                astr = inPart.ByName(aList(k), i).name
                Control.TextMatrix(k, 0) = Trim(astr)
                astr = IIf(inPart.ByName(aList(k), i).datatype = ptDate, inPart.ByName(aList(k), i).FormatedDate8, inPart.ByName(aList(k), i).value)
                Control.TextMatrix(k, i) = Trim(astr)
                If maxLen < Len(CStr(Trim(astr))) Then maxLen = Len(CStr(Trim(astr)))
            Next k
        End If
    Next i
    If maxLen > 0 And Control.Cols > 1 Then
        FormatString = "___________________"
        For i = 1 To Control.Cols - 1
            FormatString = FormatString & "|" & Left("ÓÔÏÉ×ÅÉÁ" & String(maxLen, "_"), maxLen)
        Next i
        Control.FixedCols = 1: Control.FixedRows = 1
        Control.FormatString = FormatString
    End If
    If colsFound > 0 Then Control.Cols = colsFound + 1 Else Control.Cols = 0
    If Control.Cols * Control.Rows > 0 Then
        For i = 1 To Control.Rows - 1
            For k = 1 To Control.Cols - 1
                Control.Row = i: Control.col = k: Control.CellAlignment = flexAlignLeftCenter
            Next k
        Next i
    End If
    If Not (aList Is Nothing) Then
        For i = aList.count To 1 Step -1
            aList.Remove i
        Next i
        Set aList = Nothing
    End If
End Sub

Public Sub ReplaceRowText(inRow As Integer, initialStr As String, newStr As String)
Dim i As Long
    If Control.Cols > Control.FixedCols And inRow < Control.Rows Then
        For i = Control.FixedCols To Control.Cols - 1
            If UCase(Control.TextMatrix(inRow, i)) = initialStr Then Control.TextMatrix(inRow, i) = newStr
        Next i
    End If
End Sub

Public Sub ReplaceColText(inCol As Integer, initialStr As String, newStr As String)
Dim i As Long
    If Control.Rows > Control.FixedRows And inCol < Control.Cols Then
        For i = Control.FixedRows To Control.Rows - 1
            If UCase(Control.TextMatrix(i, inCol)) = initialStr Then Control.TextMatrix(i, inCol) = newStr
        Next i
    End If
End Sub

Public Property Let VerticalFormatString(invalue As String)
Dim i As Long, k As Long, astr As String, apart As String, maxLen As Integer, FormatString As String

    astr = invalue
    k = 1
    If astr <> "" Then
        i = InStr(astr, "|")
        While i > 0
            apart = Left(astr, i - 1)
            If i < Len(astr) Then astr = Right(astr, Len(astr) - i) Else astr = ""
            If Trim(apart) <> "" Then Control.TextMatrix(k, 0) = Trim(apart)
            i = InStr(astr, "|"): k = k + 1
        Wend
    End If
End Property

Public Property Get TextMatrix(Row, col)
    TextMatrix = Control.TextMatrix(Row, col)
End Property

Public Property Let TextMatrix(Row, col, aValue)
    On Error Resume Next: Control.TextMatrix(Row, col) = aValue
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

Public Sub AddItem(Text, position)
    Control.AddItem CStr(Text), CLng(position)
End Sub

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
    'Control.TabStop = (Control.Rows * Control.Cols > 0)
    PropertyChanged "Cols"
End Property

Public Property Let FixedCols(aValue As Integer)
    Control.FixedCols = aValue
    PropertyChanged "FixedCols"
End Property

Public Property Get Text() As String
    Text = Control.Text
End Property

Public Property Get col() As Integer
    col = Control.col
End Property

Public Property Get Row() As Integer
    Row = Control.Row
End Property

Public Property Get TextArray(cellindex As Integer) As String
    TextArray = Control.TextArray(cellindex)
End Property

Public Property Let Text(aValue As String)
    Control.Text = aValue
    PropertyChanged "Text"
End Property

Public Property Let col(aValue As Integer)
    On Error Resume Next: Control.col = aValue
    PropertyChanged "Col"
End Property

Public Property Let Row(aValue As Integer)
    On Error Resume Next: Control.Row = aValue
    PropertyChanged "Row"
End Property

Public Property Let TextArray(cellindex As Integer, aValue As String)
    Control.TextArray(cellindex) = aValue
    PropertyChanged "TextArray"
End Property

Private Sub Control_DblClick()
    Dim aFlag As Boolean
    aFlag = False
    With owner
        Control.Enabled = False
        .HandleEvent SprdName & "DBLCLICK", Row, col, aFlag
        If Not aFlag Then
            .HandleEvent SprdName2 & "DBLCLICK", Row, col, aFlag
        End If
        Control.Enabled = True
    End With
End Sub

Private Sub Control_EnterCell()
    If Control.Rows * Control.Cols = 0 Then Exit Sub
    If Control.Row >= Control.FixedRows And Control.col >= Control.FixedCols Then Control.CellBackColor = RGB(192, 192, 192)
End Sub

Private Sub Control_GotFocus()
    Control_EnterCell
    Set owner.ActiveSpread = Me
End Sub
Public Property Let EditRowsCount(value As Long)
    ERowsCount = value
End Property
Public Property Get EditRowsCount() As Long
    EditRowsCount = ERowsCount
End Property

Private Sub Control_KeyDown(KeyCode As Integer, Shift As Integer)
    If Control.Rows * Control.Cols = 0 Then
        If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        If Control.Row = Control.Rows - 1 Then
            SendKeys "{TAB}"
        Else
            Control_LeaveCell
            Control.Row = Control.Row + 1
            Control_EnterCell
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If Control.Row < Control.FixedRows Or Control.col < Control.FixedCols Then Exit Sub
        If (Control.Row + 1 <= ERowsCount + Control.FixedRows) Then
            Dim tTop, tLeft, tWidth, tHeight As Long
            tTop = Control.Top + Control.CellTop - 5
            tLeft = Control.Left + Control.CellLeft
            tWidth = Control.CellWidth - 25
            tHeight = Control.CellHeight - 25
          
            EditText.Enabled = True
            EditText.Top = tTop: EditText.Left = tLeft: EditText.width = tWidth: EditText.height = tHeight
            EditText.Text = RTrim(Control.TextMatrix(Control.Row, Control.col))
            EditText.SelStart = 0
            EditText.SelLength = Len(EditText.Text)
            EditText.Visible = True
            EditText.SetFocus
        End If
    End If
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
    If KeyAscii > 32 And _
    (KeyAscii <> vbKeyUp And KeyAscii <> vbKeyDown And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab) Then
        If Control.Rows * Control.Cols = 0 Then Exit Sub
        If Control.Row < Control.FixedRows Or Control.col < Control.FixedCols Then Exit Sub
        If (Control.Row + 1 <= ERowsCount + Control.FixedRows) Then
            Dim tTop, tLeft, tWidth, tHeight As Long
            tTop = Control.Top + Control.CellTop - 5
            tLeft = Control.Left + Control.CellLeft
            tWidth = Control.CellWidth - 25
            tHeight = Control.CellHeight - 25
          
            EditText.Enabled = True
            EditText.Top = tTop: EditText.Left = tLeft: EditText.width = tWidth: EditText.height = tHeight
            EditText.Text = RTrim(Control.TextMatrix(Control.Row, Control.col))
'            If EditText.Text = "" Then
                EditText.Text = Chr(KeyAscii): EditText.SelStart = 1: EditText.SelLength = 0
'            Else
'                EditText.SelStart = 0
'                EditText.SelLength = Len(EditText.Text)
'            End If
            EditText.Visible = True
            EditText.SetFocus
'            SendKeys Chr$(KeyAscii)
        End If
    End If
End Sub

Private Sub Control_LeaveCell()
    If Control.Rows * Control.Cols = 0 Then Exit Sub
    If Control.Row >= Control.FixedRows And Control.col >= Control.FixedCols Then Control.CellBackColor = RGB(255, 255, 255)
     Control.CellBackColor = RGB(255, 255, 255)
End Sub

Private Sub Control_LostFocus()
    Control_LeaveCell
End Sub

Private Sub Control_SelChange()
    If Control.Rows * Control.Cols = 0 Then Exit Sub
    
    Control.BackColor = RGB(255, 255, 255)
    If Control.Row >= Control.FixedRows And Control.col >= Control.FixedCols Then Control.CellBackColor = RGB(192, 192, 192)
End Sub

Private Sub EditText_Change()
Dim astr As String
    astr = EditText.Text
    If Not AllowLowerCase Then
        If astr <> UCase(astr) Then
            Dim aselpos As Integer, asellength As Integer
            aselpos = EditText.SelStart
            asellength = EditText.SelLength
            
            astr = UCase(astr)
            EditText.Text = astr
            EditText.SelStart = aselpos
            EditText.SelLength = asellength
        End If
    End If
    Control.TextMatrix(Control.Row, Control.col) = EditText.Text
End Sub

Private Sub EditText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       EditText.Enabled = False
       EditText.Visible = False
       KeyCode = 0
    End If
End Sub

Private Sub EditText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
       EditText.Enabled = False
       EditText.Visible = False
       KeyAscii = 0
    End If
End Sub
Private Sub EditText_LostFocus()
    EditText.Enabled = False
    EditText.Visible = False
End Sub

Public Property Let SelectionMode(value As Integer)
    Control.SelectionMode = value
End Property

Public Property Get SelectionMode() As Integer
    SelectionMode = Control.SelectionMode
End Property


'Public Property Let TextMatrix(rowindex As Integer, colindex As Integer, aValue As String)
'    Control.TextMatrix(rowindex, colindex) = aValue
'    PropertyChanged "TextMatrix"
'End Property

Private Sub UserControl_Resize()
Control.Left = 0
Control.Top = 0
Control.width = width
Control.height = height
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
    Set Font = Control.Font
    SprdNo = NodeIntegerFld(inNode, "SprdNo", gridModel)
    SprdName = "Spd" & StrPad_(CStr(SprdNo), 3, "0", "L") 'NodeStringFld(inNode, "Name", gridModel)
    SprdName2 = UCase(NodeStringFld(inNode, "NAME", gridModel))
    If InStr(ReservedControlPrefixes, "," & Left(SprdName2, 3) & ",") > 0 Then SprdName2 = ""
    name = IIf(SprdName2 <> "", SprdName2, SprdName)
    
    Set ValidationControl = inProcessControl
    ValidationControl.AddObject SprdName, Me, True
    If Trim(SprdName2) <> "" And UCase(Trim(SprdName2)) <> UCase(Trim(SprdName)) Then
        On Error GoTo FldRegistrationError
        ValidationControl.ExecuteStatement "Set " & SprdName2 & "=" & SprdName
        GoTo FldRegistrationOk
FldRegistrationError:
        MsgBox "ËÜèïò êáôá ôç äÞëùóç ôïõ ðåäßïõ: " & SprdName & ":" & SprdName2
FldRegistrationOk:
    End If
    
    ScrLeft = NodeIntegerFld(inNode, "ScrX", gridModel)
    ScrWidth = (NodeIntegerFld(inNode, "ScrWidth", gridModel) + 1)
    ScrTop = NodeIntegerFld(inNode, "ScrY", gridModel) * 290
    ScrHeight = NodeIntegerFld(inNode, "ScrHeight", gridModel) * 285
    
    TTabIndex = NodeIntegerFld(inNode, "TabIndex", gridModel)
    
    LabelName = "GLabel" & StrPad_(CStr(SprdNo), 3, "0", "L")
        
    Set Prompt = parent.Controls.add("Vb.Label", LabelName)
    Prompt.BackColor = parent.BackColor
    Prompt.AutoSize = False
    Prompt.Caption = NodeStringFld(inNode, "ScrPrompt", gridModel)
        
    parent.ScaleMode = vbCharacters
    Prompt.Left = NodeIntegerFld(inNode, "ScrPromptX", gridModel)
    Prompt.width = NodeIntegerFld(inNode, "ScrPromptWidth", gridModel)
    parent.ScaleMode = vbTwips
    Prompt.Top = NodeIntegerFld(inNode, "ScrPromptY", gridModel) * 290
    Prompt.height = NodeIntegerFld(inNode, "ScrPromptHeight", gridModel) * 285
    DisplayFlag = NodeBooleanFld(inNode, "ScrDisplay", gridModel)
    
    
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
    
'    DocPageHead = NodeStringFld(inNode, "DocPageHead", gridModel)
'    DocPageFoot = NodeStringFld(inNode, "DocPageFoot", gridModel)
    PageHead = NodeStringFld(inNode, "DocPageHead", listModel)
    PageFoot = NodeStringFld(inNode, "DocPageFoot", listModel)
    
'Dim bstr As String, aLine As String
'    bstr = DocPageHead
'    DocHeadLines = 0
'    Do Until bstr = ""
'        aLine = ExtractFirstLineFromString(bstr)
'        DocHeadLines = DocHeadLines + 1
'    Loop
'    bstr = DocPageFoot
'    DocFootLines = 0
'    Do Until bstr = ""
'        aLine = ExtractFirstLineFromString(bstr)
'        DocFootLines = DocFootLines + 1
'    Loop
'    DocClearLines = DocLines - DocTopMargin - DocBottomMargin - DocHeadLines - DocFootLines
    
End Sub


Public Function GetControl()
    Set GetControl = Control
End Function

Public Function GetEditText()
    Set GetEditText = EditText
End Function

Public Function TranslateToProperties(inPhase) As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("GRID")
    Set attr = XML.createAttribute("NO")
    attr.nodeValue = UCase(Me.SprdNo)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("NAME")
    attr.nodeValue = UCase(Me.SprdName)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FULLNAME")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("ENABLED")
    attr.nodeValue = UCase(Me.Enabled)
    elm.setAttributeNode attr
    'Set attr = xml.createAttribute("VISIBLE")
    'attr.nodeValue = UCase(Me.IsVisible(CInt(inPhase))) '(Owner.ProcessPhase))
    'Elm.setAttributeNode attr
    Set attr = XML.createAttribute("TEXT")
    attr.nodeValue = UCase(Me.Text)
    elm.setAttributeNode attr
    
    Set attr = XML.createAttribute("COLS")
    attr.nodeValue = UCase(Me.Cols)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("ROWS")
    attr.nodeValue = UCase(Me.Rows)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FORMATSTRING")
    attr.nodeValue = UCase(Me.FormatString)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("SELECTIONMODE")
    attr.nodeValue = UCase(Me.SelectionMode)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FIXEDROWS")
    attr.nodeValue = UCase(Me.FixedRows)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("FIXEDCOLS")
    attr.nodeValue = UCase(Me.FixedCols)
    elm.setAttributeNode attr
    
    Dim i As Integer
    Dim j As Integer
'    For i = 0 To Rows - 1
'        For j = 0 To Cols - 1
'            Set attr = xml.createAttribute("ROW" + CStr(i) + CStr(j))
'            attr.nodeValue = UCase(Me.TextMatrix(i, j))
'            Elm.setAttributeNode attr
'        Next
'    Next
                    
    Dim rowelm As IXMLDOMElement
    
    For i = 0 To Rows - 1
        Set rowelm = XML.createElement("ROW")
        elm.appendChild rowelm
        'Set attr = xml.createAttribute("NO")
        'attr.nodeValue = i
        'rowElm.setAttributeNode attr
         
        For j = 0 To Cols - 1
            Dim aelm As IXMLDOMElement
            Set aelm = XML.createElement("COL")
            aelm.Text = Me.TextMatrix(i, j)
            rowelm.appendChild aelm
            'Set attr = xml.createAttribute("COL" + CStr(j))
            'attr.nodeValue = Me.TextMatrix(i, j)
            'rowElm.setAttributeNode attr
        Next
    Next
              
                    
    Set TranslateToProperties = elm
End Function

Sub SetXMLValue(elm As IXMLDOMElement, inPhase) '(CtrlName, attrName)
    Dim aattr As IXMLDOMAttribute
    If inPhase = 0 Then inPhase = 1
    For Each aattr In elm.Attributes
        Select Case aattr.baseName
            Case "FIXEDROWS"
                Me.Rows = aattr.value + 1: Me.FixedRows = aattr.value
            Case "FIXEDCOLS"
                Me.FixedCols = aattr.value
            Case "FORMATSTRING"
                Me.FormatString = aattr.value
            Case "SELECTIONMODE"
                Me.SelectionMode = aattr.value
            
            End Select
    Next aattr
    
    If elm.SelectNodes("./ROW").length > 0 Then
        Me.Rows = elm.SelectNodes("./ROW").length + Me.FixedRows
        Me.Cols = elm.selectSingleNode("./ROW").SelectNodes("./COL").length + Me.FixedCols
        'Me.ClearLines
        Dim aRow As IXMLDOMElement
        Dim aCol As IXMLDOMElement
        Dim i As Integer
        Dim j As Integer
        
        For i = 0 To elm.SelectNodes("./ROW").length - 1
           Set aRow = elm.SelectNodes("./ROW").item(i)
            For j = 0 To aRow.SelectNodes("./COL").length - 1
                 Me.TextMatrix(i + FixedRows, j + FixedCols) = aRow.SelectNodes("./COL").item(j).Text
            Next
        Next
    Else
        Me.Rows = 0
        Me.Cols = 0
    End If
    
    If inPhase = 0 Then inPhase = 1
    For Each aattr In elm.Attributes
        Select Case aattr.baseName
            Case "ROWS"
                Me.Rows = aattr.value
            Case "COLS"
                Me.Cols = aattr.value
            End Select
    Next aattr
End Sub
