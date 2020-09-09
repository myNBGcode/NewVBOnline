VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl L2Grid 
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
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "L2Grid.ctx":0000
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Control 
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   50
      FixedCols       =   0
      ForeColorSel    =   -2147483637
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   50
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "L2Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As L2Form
Public name As String
Public tVisible As Boolean, tLeft As Long, tTop As Long, tWidth As Long, tHeight As Long
Public TTabIndex As Integer, tTabStop As Boolean
Public Caption As String

Private ERowsCount As Integer
Private AllowLowerCase As Boolean
Private onClick As String
Private onRowColChange As String
Private onDblClick As String

Private Sub Control_Click()
   If onClick <> "" Then
        owner.Enabled = False
        owner.owner.DocumentManager.XmlObjectList.item(onClick).XML
        owner.Enabled = True
   End If
End Sub

Private Sub Control_DblClick()
   If onDblClick = "." Then Exit Sub
   If onDblClick <> "" Then
        owner.Enabled = False
        owner.KeyPreview = False
        Dim aname As String
        aname = onDblClick
        onDblClick = "."
        Dim ajob As cXMLDocumentJob
        Set ajob = owner.owner.DocumentManager.XmlObjectList.item(aname)
        If ajob Is Nothing Then
            MsgBox "Δεν βρέθηκε το job: " & aname, vbCritical, "Λάθος..."
            onDblClick = aname
            Exit Sub
        End If
        ajob.XML
        
        On Error Resume Next
        If onDblClick = "." Then
            onDblClick = aname
        End If
        If Not ajob.exitformflag Then owner.Enabled = True: owner.KeyPreview = True
        
    End If
End Sub

Private Sub Control_EnterCell()
    If Control.Rows * Control.Cols = 0 Then Exit Sub
    If Control.Row >= Control.FixedRows And Control.col >= Control.FixedCols Then Control.CellBackColor = RGB(192, 192, 192)
End Sub

Private Sub Control_GotFocus()
    Control_EnterCell
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
            
            EditText.Text = Chr(KeyAscii): EditText.SelStart = 1: EditText.SelLength = 0
'            If EditText.Text = "" Then
'                EditText.Text = Chr(KeyAscii): EditText.SelStart = 1: EditText.SelLength = 0
'            Else
'                EditText.SelStart = 0
'                EditText.SelLength = Len(EditText.Text)
'            End If

            EditText.Visible = True
            EditText.SetFocus
          
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

Private Sub Control_RowColChange()
   If onRowColChange <> "" Then
        owner.Enabled = False
        owner.owner.DocumentManager.XmlObjectList.item(onRowColChange).XML
        owner.Enabled = True
   End If
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

Private Sub UserControl_Resize()
Control.Left = 0
Control.Top = 0
Control.width = width
Control.height = height
End Sub


Public Function IXMLDOMElementView() As IXMLDOMElement

Dim XML As DOMDocument30
Set XML = New DOMDocument30

Dim elm As IXMLDOMElement
Dim attr As IXMLDOMAttribute

    Set elm = XML.createElement("grid")
    Set attr = XML.createAttribute("name")
    attr.nodeValue = UCase(Me.name)
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("enabled")
    attr.nodeValue = Control.Enabled
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("visible")
    attr.nodeValue = tVisible
    elm.setAttributeNode attr
    
    Set attr = XML.createAttribute("col")
    attr.nodeValue = Control.col
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("row")
    attr.nodeValue = Control.Row
    elm.setAttributeNode attr
    
    Set attr = XML.createAttribute("cols")
    attr.nodeValue = Control.Cols
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("rows")
    attr.nodeValue = Control.Rows
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("formatstring")
    attr.nodeValue = Control.FormatString
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("selectionmode")
    attr.nodeValue = Control.SelectionMode
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("fixedrows")
    attr.nodeValue = Control.FixedRows
    elm.setAttributeNode attr
    Set attr = XML.createAttribute("fixedcols")
    attr.nodeValue = Control.FixedCols
    elm.setAttributeNode attr
    
    Set attr = XML.createAttribute("caption")
    attr.nodeValue = Me.Caption
    elm.setAttributeNode attr
    
    Dim i As Integer
    Dim j As Integer
    Dim selectedrow As IXMLDOMElement
                    
    Dim rowelm As IXMLDOMElement
    'If FormatString <> "" And Control.FixedRows >= 1 Then
        For i = 0 To Control.Rows - 1
            Set rowelm = XML.createElement("row")
            elm.appendChild rowelm
    
            Set attr = XML.createAttribute("row")
            attr.nodeValue = i
            rowelm.setAttributeNode attr
             
            For j = 0 To Control.Cols - 1
                Dim aelm As IXMLDOMElement
                Set aelm = XML.createElement("col")
                aelm.Text = Control.TextMatrix(i, j)
                rowelm.appendChild aelm
            
                Set attr = XML.createAttribute("col")
                attr.nodeValue = j
                aelm.setAttributeNode attr
                
            Next
            If Control.Row = i Then Set selectedrow = rowelm
        Next
    'End If
    If Not (selectedrow Is Nothing) Then
        Set selectedrow = selectedrow.cloneNode(True)
        Set rowelm = XML.createElement("selectedrow")
        elm.appendChild rowelm
        
        Dim child As IXMLDOMNode
        For Each child In selectedrow.childNodes
            rowelm.appendChild child
        Next child
        
    End If
    
                    
    Set IXMLDOMElementView = elm
End Function

Sub LoadFromIXMLDOMElement(elm As IXMLDOMElement)
    Dim aattr As IXMLDOMAttribute
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
                    tTabStop = True
                ElseIf UCase(aattr.value) = bvFalse Then
                    tTabStop = False
                End If
            Case "ENABLED"
                If UCase(aattr.value) = bvTrue Then
                    Enabled = True
                ElseIf UCase(aattr.value) = bvFalse Then
                    Enabled = False
                End If
            Case "TABINDEX"
                TTabIndex = aattr.value
            Case "VISIBLE"
                If UCase(aattr.value) = bvFalse Then
                    Me.tVisible = False
                ElseIf UCase(aattr.value) = bvTrue Then
                    Me.tVisible = True
                End If
            Case "FIXEDROWS"
                Control.Rows = aattr.value + 1: Control.FixedRows = aattr.value
            Case "FIXEDCOLS"
                Control.FixedCols = aattr.value
            Case "FORMATSTRING"
                Control.FormatString = aattr.value
                
            Case "SELECTIONMODE"
                Control.SelectionMode = aattr.value
            Case "ROWS"
                If aattr.value = "NaN" Then
                    Control.Rows = 0
                Else
                    Control.Rows = aattr.value
                End If
            Case "COLS"
                Control.Cols = aattr.value
            Case "ONCLICK"
                onClick = aattr.value
            Case "ONROWCOLCHANGE"
                onRowColChange = aattr.value
            Case "CAPTION"
                Caption = aattr.value
            Case "EDITROWSCOUNT"
                EditRowsCount = aattr.value
            Case "ONDBLCLICK"
                onDblClick = aattr.value
            End Select
    Next aattr
    
    If elm.SelectNodes("./row").length > 0 Then
        Control.Rows = elm.SelectNodes("./row").length + Control.FixedRows
        Control.Cols = elm.selectSingleNode("./row").SelectNodes("./col").length '+ Control.FixedCols 9060newver
        'Me.ClearLines
        Dim aRow As IXMLDOMElement
        Dim aCol As IXMLDOMElement
        Dim i As Integer
        Dim j As Integer
        Dim formatattr As IXMLDOMAttribute
        For i = 0 To elm.SelectNodes("./row").length - 1
           Set aRow = elm.SelectNodes("./row").item(i)
            For j = 0 To aRow.SelectNodes("./col").length - 1
                 'Control.TextMatrix(i + Control.FixedRows, j + Control.FixedCols) = aRow.SelectNodes("./col").item(j).Text 9060newver
                 Set formatattr = aRow.SelectNodes("./col").item(j).Attributes.getNamedItem("format")
                 If Not (formatattr Is Nothing) Then
                    Dim colItem(1) As String
                    colItem(0) = aRow.SelectNodes("./col").item(j).Text
                    Control.TextMatrix(i + Control.FixedRows, j) = gFormatType_(formatattr.Text, colItem)
                 Else
                    Control.TextMatrix(i + Control.FixedRows, j) = aRow.SelectNodes("./col").item(j).Text
                 End If
            Next
        Next
    Else
        'Control.Rows = 0
        'Control.Cols = 0
    End If
    
    For Each aattr In elm.Attributes
        Select Case UCase(aattr.baseName)
            Case "ROWS"
                Control.Rows = aattr.value
            Case "COLS"
                Control.Cols = aattr.value
            End Select
    Next aattr

    Dim textmatrixelm As IXMLDOMElement
    For Each textmatrixelm In elm.SelectNodes("./textmatrix")
    
        Dim rowattr As IXMLDOMAttribute, colattr As IXMLDOMAttribute
        Set rowattr = textmatrixelm.Attributes.getNamedItem("row")
        Set colattr = textmatrixelm.Attributes.getNamedItem("col")
        Dim Row As Long, col As Long
        If rowattr Is Nothing Then Row = 0 Else Row = rowattr.value
        If colattr Is Nothing Then col = 0 Else col = colattr.value
        Control.TextMatrix(Row, col) = textmatrixelm.Text
    
    Next textmatrixelm
    
End Sub


Public Sub CreateFromIXMLDOMElement(inOwner As L2Form, inNode As MSXML2.IXMLDOMElement)

    Set owner = inOwner
    Dim aattr As IXMLDOMAttribute
    Set aattr = inNode.Attributes.getNamedItem("name")
    If Not (aattr Is Nothing) Then name = aattr.value
    
    LoadFromIXMLDOMElement inNode
    Control.AllowUserResizing = flexResizeColumns
        
End Sub

Public Sub CleanUp()
    Set owner = Nothing
End Sub

