VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl GenRichTextBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin RichTextLib.RichTextBox vcontrol 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      TextRTF         =   $"GenRichTextBox.ctx":0000
   End
End
Attribute VB_Name = "GenRichTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As Form
Private DisplayFlag(10) As Boolean
Private RichTextBoxName As String
Public ScrLeft As Long, ScrTop As Long, ScrWidth As Long, ScrHeight As Long
Private ValidationControl As ScriptControl
Private rtfText As String
Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const QuarterInch As Long = 360

Public Property Get Control()
    Set Control = vControl
End Property

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Private Sub UserControl_Resize()
    vControl.Left = 0: vControl.Top = 0: vControl.width = width: vControl.height = height
End Sub

Public Sub Initialize(inOwner As Form, inProcessControl As ScriptControl, name As String, _
    wLeft As Long, wTop As Long, wWidth As Long, wHeight As Long)
Dim i As Integer
    Set owner = inOwner
    Set ValidationControl = inProcessControl
    RichTextBoxName = name
    ValidationControl.AddObject name, Me, True
    
    For i = 1 To 10
        DisplayFlag(i) = True
    Next i
 
    ScrLeft = wLeft
    ScrWidth = wWidth
    ScrTop = wTop * 290
    ScrHeight = wHeight * 285
        
End Sub


Private Function ProcessBookmarkTree(richtext As String, Node) As String
    Dim bookmarkname As String, bkmkstart As String, bkmkend As String, rplc As String, rslt As String
    Dim i As Long, k As Long, StartPos As Long, endpos As Long
    Dim childnode As MSXML2.IXMLDOMNode, childsfound As Boolean
    bookmarkname = Node.baseName
    bkmkstart = "{\*\bkmkstart " + bookmarkname + "}"
    i = InStr(1, richtext, bkmkstart)
    If i = 0 Then
        bkmkstart = "{\*\bkmkstart " + bookmarkname + vbCrLf + "}"
        i = InStr(1, richtext, bkmkstart)
    End If
    bkmkend = "{\*\bkmkend " + bookmarkname + "}"
    k = InStr(1, richtext, bkmkend)
    If k = 0 Then
        bkmkend = "{\*\bkmkend " + bookmarkname + vbCrLf + "}"
        k = InStr(1, richtext, bkmkend)
    End If
    
    StartPos = i + Len(bkmkstart)
    endpos = k
    
    If StartPos < endpos And StartPos > 1 Then
        rplc = Mid(richtext, StartPos, endpos - StartPos)
        rslt = rplc: childsfound = False
        For Each childnode In Node.childNodes
            If childnode.nodeType = NODE_ELEMENT Then
                rslt = ProcessBookmarkTree(rslt, childnode): childsfound = True
            End If
        Next childnode
        If Not childsfound Then
            rslt = Node.Text
        End If
        
        Dim aattr As IXMLDOMAttribute
        Set aattr = Node.Attributes.getNamedItem("action")
        If aattr Is Nothing Then
        
            ProcessBookmarkTree = Replace(richtext, bkmkstart & rplc & bkmkend, bkmkstart & rslt & bkmkend)
        Else
            If aattr.value = "insertbefore" Then
                ProcessBookmarkTree = Replace(richtext, bkmkstart & rplc & bkmkend, rslt & bkmkstart & rplc & bkmkend)
            ElseIf aattr.value = "insertafter" Then
                ProcessBookmarkTree = Replace(richtext, bkmkstart & rplc & bkmkend, bkmkstart & rplc & bkmkend & rslt)
            ElseIf aattr.value = "replace" Then
                ProcessBookmarkTree = Replace(richtext, bkmkstart & rplc & bkmkend, bkmkstart & rslt & bkmkend)
            ElseIf aattr.value = "remove" Then
                ProcessBookmarkTree = Replace(richtext, bkmkstart & rplc & bkmkend, bkmkstart & bkmkend)
            End If
        End If
    Else
        ProcessBookmarkTree = richtext
    End If
End Function

Public Function processdocument(filename, Data) As Boolean
    processdocument = False
    Dim astring As String
    On Error GoTo fileopenerror
    Open ReadDir & "\Reports\" & filename For Input As #1
    On Error GoTo filereaderror
    astring = input$(LOF(1), 1)
    Close #1
    On Error GoTo fileprocesserror
    rtfText = ProcessBookmarkTree(astring, Data)
    On Error GoTo fileopenerror
    On Error GoTo fileshowerror
    vControl.TextRTF = rtfText
    'vControl.SaveFile "c:\test1.rtf"
    processdocument = True
    Exit Function
fileopenerror:
    MsgBox "Πρόβλημα στο άνοιγμα του αρχείου: " & ReadDir & "\Messages\" & filename & " " & Err.Number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
filereaderror:
    MsgBox "Πρόβλημα στο διάβασμα του αρχείου: " & ReadDir & "\Messages\" & filename & " " & Err.Number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
fileprocesserror:
    MsgBox "Πρόβλημα επεξεργασία του αρχείου: " & ReadDir & "\Messages\" & filename & " " & Err.Number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
fileshowerror:
    MsgBox "Πρόβλημα προβολή του αρχείου: " & ReadDir & "\Messages\" & filename & " " & Err.Number & " " & Err.description, vbCritical, "Λάθος"
    Exit Function
End Function

Public Sub processBookmark(bookmarkname, bookmarkvalue)
    Dim bkmkstart As String, bkmkend As String, rplc As String, rslt As String
    Dim i As Long, k As Long, StartPos As Long, endpos As Long
    bkmkstart = "{\*\bkmkstart " + bookmarkname + "}"
    i = InStr(1, rtfText, bkmkstart)
    If i = 0 Then
        bkmkstart = "{\*\bkmkstart " + bookmarkname + vbCrLf + "}"
        i = InStr(1, rtfText, bkmkstart)
    End If
    bkmkend = "{\*\bkmkend " + bookmarkname + "}"
    k = InStr(1, rtfText, bkmkend)
    If k = 0 Then
        bkmkend = "{\*\bkmkend " + bookmarkname + vbCrLf + "}"
        k = InStr(1, rtfText, bkmkend)
    End If
    
    StartPos = i + Len(bkmkstart)
    endpos = k
    
    If StartPos < endpos And StartPos > 1 Then
        rplc = Mid(rtfText, StartPos, endpos - StartPos)
        rslt = bookmarkvalue
        
        rtfText = Replace(rtfText, bkmkstart & rplc & bkmkend, bkmkstart & rslt & bkmkend)
        vControl.TextRTF = rtfText
    End If
    
    Dim PrintableWidth As Long
    Dim PrintableHeight As Long
    'WYSIWYG_RTF VControl, AnInch / 2, AnInch / 2, AnInch / 2, AnInch / 2, PrintableWidth, PrintableHeight
    'Me.width = PrintableWidth + 200
    'Me.height = PrintableHeight + 800
End Sub

Public Function PrintDocument(Optional copies) As Boolean
    Dim aFrm As SelectPrinterFrm
    PrintDocument = False
    Set aFrm = New SelectPrinterFrm
    aFrm.excludepassbook = True
    Load aFrm
    Dim res As Long
    aFrm.Show vbModal, owner

    Dim aPrinterName As String
    aPrinterName = aFrm.SelectedPrinter

    Dim x As Printer
        For Each x In Printers
            If aPrinterName = x.DeviceName Then _
                Set Printer = x: Exit For
        Next
    Set aFrm = Nothing
    If aPrinterName = "" Then Exit Function
    
    vControl.SelStart = 0
    vControl.SelLength = Len(vControl.Text)
    If IsMissing(copies) Then copies = 1
    Printer.PaperSize = vbPRPSA4
    Dim i As Integer
    For i = 1 To copies
        PrintRTF vControl, QuarterInch * 2, QuarterInch * 2, QuarterInch * 2, QuarterInch * 2

        'VControl.SelPrint Printer.hdc
    Next i
    PrintDocument = True
End Function

