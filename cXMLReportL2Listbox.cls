VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLReportL2Listbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement
Public ESignature As IXMLDOMAttribute

Private eSignSocket As cXmlESignatureSocketHandler

Private PageLines As Integer, PageColumns As Integer
Private PageHeaderLines As Integer, PageHeaderNode As IXMLDOMElement
Private PageFooterLines As Integer, PageFooterNode As IXMLDOMElement
Private curPageNo As Integer, curPageLine As Integer
Const MaxLinesCount = 200

Const A4Lines = 54
Const A5Lines = 45

Private PageLinedata(MaxLinesCount) As String

Private unlockattribute As IXMLDOMAttribute

Public PrintPromptMessage As String
Public TotalPages As Integer
Private printPages As Integer
Private PrintCondensed As Boolean

Private Sub Class_Initialize()
    Set eSignSocket = New cXmlESignatureSocketHandler
End Sub
Private Sub Class_Terminate()
    Set content = Nothing
    Set ESignature = Nothing
    Set eSignSocket = Nothing
End Sub

Public Sub PageHeader(elm As IXMLDOMElement)
    If elm Is Nothing Then Exit Sub
    PageHeaderLines = elm.SelectNodes(".//line").length
    Set PageHeaderNode = elm
End Sub

Public Sub PageFooter(elm As IXMLDOMElement)
    If elm Is Nothing Then Exit Sub
    PageFooterLines = elm.SelectNodes(".//line").length
    Set PageFooterNode = elm
End Sub

Private Sub PrintLine(Node As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    Dim colNode As IXMLDOMElement
    Dim Counter As Integer, i As Integer, astr As String
    Dim colitems() As String, FormatString As String
    FormatString = ""
    Set aattr = Node.Attributes.getNamedItem("format")
    If Not (aattr Is Nothing) Then FormatString = aattr.Text
        
    Counter = Node.SelectNodes("./col").length
    If Counter > 0 Then
        ReDim colitems(Counter): i = 1
        For Each colNode In Node.SelectNodes("./col")
           
          If UCase(Trim(colNode.Text)) = "%PG%" Then
                colitems(i - 1) = curPageNo
          ElseIf UCase(Trim(colNode.Text)) = "%GG%" Then
                colitems(i - 1) = TotalPages
            Else
                colitems(i - 1) = colNode.Text
            End If
            i = i + 1
            Set aattr = colNode.Attributes.getNamedItem("format")
            If Not (aattr Is Nothing) Then FormatString = FormatString & aattr.Value
        Next colNode
        
        astr = gFormat_(FormatString, colitems)
    Else
        astr = FormatString
    End If
    
    PageLinedata(curPageLine) = astr
    curPageLine = curPageLine + 1
End Sub

Private Sub PrintLinePassbook(Node As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
    Dim colNode As IXMLDOMElement
    Dim Counter As Integer, i As Integer, astr As String
    Dim colitems() As String, FormatString As String
    FormatString = ""
    Set aattr = Node.Attributes.getNamedItem("format")
    If Not (aattr Is Nothing) Then FormatString = aattr.Text
        
    Counter = Node.SelectNodes("./col").length
    If Counter > 0 Then
        ReDim colitems(Counter): i = 1
        For Each colNode In Node.SelectNodes("./col")
            If UCase(Trim(colNode.Text)) = "%PG%" Then
                colitems(i - 1) = curPageNo
            ElseIf UCase(Trim(colNode.Text)) = "%GG%" Then
              colitems(i - 1) = TotalPages
            Else
                colitems(i - 1) = colNode.Text
            End If
            i = i + 1
            Set aattr = colNode.Attributes.getNamedItem("format")
            If Not (aattr Is Nothing) Then FormatString = FormatString & aattr.Value
        Next colNode
        astr = gFormat_(FormatString, colitems)
    Else
        astr = FormatString
    End If
    
    PageLinedata(curPageLine) = astr
    curPageLine = curPageLine + 1
End Sub

Private Sub PrintPageHeader()
    Dim Node As IXMLDOMElement
    If PageHeaderNode Is Nothing Then Exit Sub
    For Each Node In PageHeaderNode.SelectNodes("./line")
        PrintLine Node
    Next Node
End Sub

Private Sub PrintPageFooter()
    Dim Node As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    
    If PageFooterNode Is Nothing Then Exit Sub
    
    Set aattr = PageFooterNode.Attributes.getNamedItem("lastpage")
    If aattr Is Nothing Then
        For Each Node In PageFooterNode.SelectNodes("./line")
            PrintLine Node
        Next Node
    End If

End Sub

Private Sub PrintLastPageFooter()
    Dim Node As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    
    If PageFooterNode Is Nothing Then Exit Sub
    
    Set aattr = PageFooterNode.Attributes.getNamedItem("lastpage")
    If Not (aattr Is Nothing) Then
        For Each Node In PageFooterNode.SelectNodes("./line")
            PrintLine Node
        Next Node
    End If
        
End Sub


Private Sub PrintPageHeaderPassbook()
    Dim Node As IXMLDOMElement
    If PageHeaderNode Is Nothing Then Exit Sub
    For Each Node In PageHeaderNode.SelectNodes("./line")
        PrintLinePassbook Node
    Next Node
End Sub

Private Sub PrintPageFooterPassbook()
    Dim Node As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    
    If PageFooterNode Is Nothing Then Exit Sub
    Set aattr = PageFooterNode.Attributes.getNamedItem("lastpage")
    If aattr Is Nothing Then
        For Each Node In PageFooterNode.SelectNodes("./line")
            PrintLinePassbook Node
        Next Node
    End If
End Sub
Private Sub PrintLastPageFooterPassbook()
    Dim Node As IXMLDOMElement
    Dim aattr As IXMLDOMAttribute
    
    If PageFooterNode Is Nothing Then Exit Sub
    Set aattr = PageFooterNode.Attributes.getNamedItem("lastpage")
    If Not (aattr Is Nothing) Then
        For Each Node In PageFooterNode.SelectNodes("./line")
            PrintLinePassbook Node
        Next Node
    End If
End Sub

Private Function SelectPrinter(elm As IXMLDOMElement) As String
Dim aFrm As SelectPrinterFrm
Dim aPrinterName As String
Dim aattr As IXMLDOMAttribute
Dim cattr As IXMLDOMAttribute
Dim battr As IXMLDOMAttribute
    
    Set aattr = elm.Attributes.getNamedItem("printertype")
    Set cattr = elm.Attributes.getNamedItem("printercopies")
    Set battr = elm.Attributes.getNamedItem("cancelprint")
    If aattr Is Nothing Then
        Set aattr = elm.Attributes.getNamedItem("printername")
        If aattr Is Nothing Then
    
            Set aFrm = New SelectPrinterFrm
            
            If battr Is Nothing Then
            Else
                If battr.Text = "true" Or battr.Text = "True" Or battr.Text = "TRUE" Then
                    aFrm.Cancel_CMD.Visible = True
                End If
            End If
            
            Load aFrm
            aFrm.Show vbModal
            aPrinterName = aFrm.SelectedPrinter
        Else
            aPrinterName = aattr.Text
        End If
        
        Dim x As Printer
            For Each x In Printers
                If aPrinterName = x.DeviceName Then _
                    Set Printer = x: Exit For
            Next
        Set aFrm = Nothing
        If aPrinterName = "" Then Exit Function
        SelectPrinter = aPrinterName
        If UCase(aPrinterName) = "PASSBOOK" And cPassbookPrinter = 5 Then
            SelectPrinter = "Passbook"
        Else
            Printer.ScaleMode = vbCharacters
            Printer.FontName = "Courier New Greek"
            Printer.FontSize = 9
            
            Set aattr = elm.Attributes.getNamedItem("orientation")
            If aattr Is Nothing Then
            Else
                If aattr.Text = "portrait" Then
                    Printer.Orientation = vbPRORPortrait
                ElseIf aattr.Text = "landscape" Then
                    Printer.Orientation = vbPRORLandscape
                End If
            End If
            Printer.ScaleMode = vbCharacters
            Printer.FontName = "Courier New Greek"
            Printer.FontSize = 9
            
            If cattr Is Nothing Then
            Else
                Dim IsPdf
                Dim IsXps
                Dim IsNote
                Dim strPdf As String
                Dim strXPS As String
                Dim strOneNote As String
                Dim PrinterNameUCase As String
                                
                PrinterNameUCase = UCase(aPrinterName)
                
                strPdf = "PDF"
                strXPS = "XPS"
                strOneNote = "ONENOTE"
                
                IsPdf = InStr(PrinterNameUCase, strPdf)
                IsXps = InStr(PrinterNameUCase, strXPS)
                IsNote = InStr(PrinterNameUCase, strOneNote)

                If IsPdf > 0 Or IsXps > 0 Or IsNote > 0 Then
                    MsgBox ("Εκτύπωση σε pdf,xps,OneNote - 1 αρχείο")
                Else
                    Printer.copies = cattr.Text
                End If
            End If
       
        End If
    ElseIf UCase(aattr.Text) = UCase("passbook") Then
        SelectPrinter = "Passbook"
    End If
    
    If Not (content Is Nothing) Then
        
        Set aattr = content.Attributes.getNamedItem("printername")
        If aattr Is Nothing Then
            Set aattr = content.ownerDocument.createAttribute("printername")
            aattr.Value = SelectPrinter
            content.Attributes.setNamedItem aattr
        Else
            aattr.Value = SelectPrinter
        End If
    End If
End Function

Private Sub PrintPageLineData()
    Dim i As Integer
    For i = 1 To PageLines
        Printer.CurrentX = 0: Printer.CurrentY = i
        If PageLinedata(i) <> "" Then Printer.Print PageLinedata(i)
    Next i
End Sub

Private Sub PrintPageLineDataPassbook()
    Dim i As Integer
    For i = 1 To PageLines
        xSetDocLine_ i, PageLinedata(i)
    Next i
End Sub

Private Sub PrintLineList(Node As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
   
    Dim Value As String
    
    Set aattr = Node.Attributes.getNamedItem("value")
    If Not (aattr Is Nothing) Then Value = aattr.Text
    
    PageLinedata(curPageLine) = Value
    curPageLine = curPageLine + 1
End Sub
Private Sub PrintLinePassBookList(Node As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
   
    Dim Value As String
    
    Set aattr = Node.Attributes.getNamedItem("value")
    If Not (aattr Is Nothing) Then Value = aattr.Text
    
    PageLinedata(curPageLine) = Value
    curPageLine = curPageLine + 1
End Sub


Public Sub PrintLines(elm As IXMLDOMElement)
    curPageNo = 1: curPageLine = 1:
    Dim Node As IXMLDOMElement
    Dim pageNode As IXMLDOMElement
    Dim i As Integer
    
    TotalPages = 0
    For Each pageNode In elm.SelectNodes(".//page")
        TotalPages = TotalPages + pageNode.SelectNodes(".//line").length
    Next
    If PageLines <> 0 Then
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ PageLines) + 1
    ElseIf (PrintCondensed) Then
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ A5Lines) + 1
    Else
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ 54) + 1
    End If

    For i = 1 To MaxLinesCount
        PageLinedata(i) = ""
    Next i
    
    curPageLine = 0

    For Each pageNode In elm.SelectNodes(".//page")
        If curPageLine <> 0 Then 'για καθε σελιδα εκτος της πρωτης
            PrintPageFooter
            PrintPageLineData
            curPageNo = curPageNo + 1: curPageLine = 1
            For i = 1 To MaxLinesCount
              PageLinedata(i) = ""
            Next i
            Printer.NewPage
        End If
        curPageLine = 1
        For Each Node In pageNode.SelectNodes(".//lines/item")
            If curPageLine = 1 Then PrintPageHeader
            PrintLineList Node
            If curPageLine + PageFooterLines > PageLines Then
                PrintPageFooter
                
                PrintPageLineData
                curPageNo = curPageNo + 1: curPageLine = 1
                If printPages <> -1 And printPages <> 0 And curPageNo > printPages Then
                   For i = 1 To MaxLinesCount
                       PageLinedata(i) = ""
                   Next i
                   Exit For
                Else
                   Printer.NewPage
                   For i = 1 To MaxLinesCount
                       PageLinedata(i) = ""
                   Next i
                End If
            End If
        Next Node

    Next pageNode
    
    If curPageLine = 1 Then
    ElseIf curPageLine > 1 Then
        PrintPageFooter
        'last page footer
        curPageLine = PageLines - PageFooterLines + 1
        PrintLastPageFooter
        
        PrintPageLineData
        
        curPageNo = curPageNo + 1: curPageLine = 1
        Printer.NewPage
    End If
    
    Printer.EndDoc
End Sub

Public Sub PrintLinesPassbook(elm As IXMLDOMElement)
    Dim i As Integer
    Dim attr As IXMLDOMAttribute
    
    xClearDoc_
    For i = 1 To MaxLinesCount
        PageLinedata(i) = ""
    Next i
    
    If gPanel Is Nothing Then Set gPanel = New GlobalSPCPanel
    
    Set attr = elm.Attributes.getNamedItem("printpromptmessage")
    
    If Not (attr Is Nothing) Then
        PrintPromptMessage = attr.Value
    Else
        PrintPromptMessage = "ΕΚΤΥΠΩΣΗ ΠΑΡΑΣΤΑΤΙΚΟΥ..."
    End If
    
    curPageNo = 1: curPageLine = 1:
    Dim Node As IXMLDOMElement
    Dim pageNode As IXMLDOMElement
    
    Dim OCRFlag As Boolean
    Dim reportelm As IXMLDOMElement
    Dim OCRAttr As IXMLDOMAttribute
    OCRFlag = False
    Set reportelm = elm.selectSingleNode("/report")
    If Not (reportelm Is Nothing) Then
       Set OCRAttr = reportelm.Attributes.getNamedItem("ocr")
       If Not (OCRAttr Is Nothing) Then
          If UCase(OCRAttr.Value) = "TRUE" Then
            OCRFlag = True
          Else
            OCRFlag = False
          End If
       End If
    End If
    
    
    TotalPages = 0
    For Each pageNode In elm.SelectNodes(".//page")
        TotalPages = TotalPages + pageNode.SelectNodes(".//line").length
    Next
    If PageLines <> 0 Then
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ PageLines) + 1
    ElseIf (PrintCondensed) Then
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ A5Lines) + 1
    Else
        TotalPages = ((TotalPages + PageFooterLines + PageHeaderLines) \ 54) + 1
    End If
    curPageLine = 0
    For Each pageNode In elm.SelectNodes(".//page")
        If curPageLine <> 0 Then 'για καθε σελιδα εκτος της πρωτης
            PrintPageFooterPassbook
            PrintPageLineDataPassbook
            curPageNo = curPageNo + 1: curPageLine = 1
            xPrintDoc_ Nothing, PrintPromptMessage, , PrintCondensed: xClearDoc_:
            For i = 1 To MaxLinesCount
              PageLinedata(i) = ""
            Next i
        End If
        curPageLine = 1
        For Each Node In pageNode.SelectNodes(".//lines/item")
            If curPageLine = 1 Then PrintPageHeaderPassbook
            PrintLinePassBookList Node
            If curPageLine + PageFooterLines > PageLines Then
                PrintPageFooterPassbook
                
                PrintPageLineDataPassbook
                curPageNo = curPageNo + 1: curPageLine = 1
                If printPages <> -1 And printPages <> 0 And curPageNo > printPages + 1 Then
                    For i = 1 To MaxLinesCount
                        PageLinedata(i) = ""
                    Next
                    Exit For
                Else
                    xPrintDoc_ Nothing, PrintPromptMessage, OCRFlag, PrintCondensed: xClearDoc_:
                    For i = 1 To MaxLinesCount
                        PageLinedata(i) = ""
                    Next
                End If
            End If
        Next Node

    Next pageNode
    If curPageLine = 1 Then
    ElseIf curPageLine > 1 Then
        PrintPageFooterPassbook
        
        'last page footer
        curPageLine = PageLines - PageFooterLines + 1
        PrintLastPageFooterPassbook
        
        PrintPageLineDataPassbook
        curPageNo = curPageNo + 1: curPageLine = 1
        xPrintDoc_ Nothing, PrintPromptMessage, OCRFlag, PrintCondensed: xClearDoc_:
        For i = 1 To MaxLinesCount
            PageLinedata(i) = ""
        Next i
    End If
    
    If unlockattribute Is Nothing Then
        Set gPanel = Nothing
    ElseIf unlockattribute.Text = "true" Then
        Set gPanel = Nothing
    ElseIf unlockattribute.Text = "false" Then
    
    End If
End Sub


Public Function LoadXML(Value As String) As String
    Dim reportDoc As New MSXML2.DOMDocument30
    reportDoc.preserveWhiteSpace = True
    reportDoc.LoadXML Value
    If reportDoc.documentElement Is Nothing Then Exit Function
        
    If ESignature Is Nothing Or HasPad = "" Then
    ElseIf ESignature.Value = "true" And Trim(HasPad) = "1" Then
            
        If eSignSocket.TrySendData(reportDoc) Then
            LoadXML = content.XML
            Exit Function
        End If
    End If
    
    Dim i As Integer
    For i = 1 To 200
        PageLinedata(i) = ""
    Next i

    Dim aPrinterName As String, aPrinterNameAttr As IXMLDOMAttribute
    Set aPrinterNameAttr = reportDoc.documentElement.Attributes.getNamedItem("printer")
    If aPrinterNameAttr Is Nothing Then
        aPrinterName = ""
    Else
        aPrinterName = aPrinterNameAttr.Value
    End If
    
    Dim aprinterPagesAttr As IXMLDOMAttribute
    Set aprinterPagesAttr = reportDoc.documentElement.Attributes.getNamedItem("printpages")
    If aprinterPagesAttr Is Nothing Then
       printPages = -1
    Else
       printPages = CInt(aprinterPagesAttr.Text)
    End If
    Dim condensed As IXMLDOMAttribute
    Set condensed = reportDoc.documentElement.Attributes.getNamedItem("condensed")
    If (condensed Is Nothing) Then
        PrintCondensed = False
    ElseIf ((Not condensed Is Nothing) And condensed.Value = "-1") Then
        PrintCondensed = True
    Else
        PrintCondensed = False
    End If
    
    Set unlockattribute = reportDoc.documentElement.Attributes.getNamedItem("unlock")
        
    If aPrinterName = "" Then aPrinterName = SelectPrinter(reportDoc.documentElement)
    If aPrinterName <> "" Then
        If UCase(aPrinterName) = UCase("Passbook") Then PageLines = 54
        If UCase(aPrinterName) = UCase("Passbook") And PrintCondensed Then PageLines = A5Lines
        
        Dim aattr As IXMLDOMAttribute
        Set aattr = reportDoc.documentElement.Attributes.getNamedItem("pagelines")
        If aattr Is Nothing Then
        Else
            PageLines = aattr.Text
        End If
        
        If UCase(aPrinterName) = UCase("Passbook") And PrintCondensed = False And PageLines > 54 Then PageLines = 54
        If UCase(aPrinterName) = UCase("Passbook") And PrintCondensed = True And PageLines > A5Lines Then PageLines = A5Lines
        
        PageHeader reportDoc.documentElement.selectSingleNode("./head")
        PageFooter reportDoc.documentElement.selectSingleNode("./foot")
        
        If UCase(aPrinterName) = UCase("Passbook") And cPassbookPrinter = 5 Then
            PrintLinesPassbook reportDoc.documentElement
        Else
            PrintLines reportDoc.documentElement
        End If
        
    End If
    LoadXML = content.XML
End Function

Public Function XML() As String
    If content Is Nothing Then
        XML = ""
    Else
        XML = LoadXML(content.XML)
    End If
End Function

