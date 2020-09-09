VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLRecordsetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DOM As New DOMDocument30
Public FieldsExplicit As Boolean
Public name As String
Public ActiveRow As Long
Public ErrorCount As Long
Public ResultCount As Long
Public RowCount As Long

Dim aXMapper 'As XMLMapperInProc.XMLMapperLinkIn
Dim aSMapper 'As soapclient
'Dim Fields As XMLFields
Dim afield As XMLField

Public Property Get RecordCount() As Long
    RecordCount = RowCount
End Property

Public Property Get Eof() As Boolean
    Eof = (ActiveRow = -1)
End Property

Public Property Get Bof() As Boolean
    Bof = (ActiveRow = -1)
End Property

Public Sub prepare(inSMapper, inXMapper)
    Set aSMapper = inSMapper
    Set aXMapper = inXMapper
    
    'Fields.XMLFields("Test").Value = 5 '***
    'Fields.XMLFields("Name").Value = 5
End Sub

'Public Sub ReadXMLMapperX(aSqlCmd)
'Dim list 'As DOMDocument
'    If Not (aXMapper Is Nothing) Then
'        aXMapper.GetFullTaggedRowsXML CStr(aSqlCmd), list
'        SetNodes list.documentElement.childNodes
'    End If
'End Sub

Public Sub ReadXMLMapperX(ConnectionString As Variant, SqlCmd As Variant)
Dim list
    If Not (aXMapper Is Nothing) Then
        aXMapper.GetFullTaggedRowsXML ConnectionString, SqlCmd, list
        SetNodes2 list
    End If
End Sub

Public Sub ReadXMLMapperS(aSqlCmd)
Dim list
    If Not (aSMapper Is Nothing) Then
        aSMapper.GetFullTaggedRowsXML CStr(aSqlCmd), list
        SetNodes2 list
    End If

End Sub

Public Sub Read(aSqlCmd)
    ReadXMLMapperS aSqlCmd
End Sub

Public Sub ExecInsert(aSqlCmd)
Dim list
    If Not (aSMapper Is Nothing) Then
        aSMapper.InsertRows CStr(aSqlCmd), DOM.XML, list
        SetNodes2 list
    End If
End Sub

'Public Sub ExecInsertX(aSqlCmd)
'Dim list 'As IXMLDOMNodeList
'    If Not (aXMapper Is Nothing) Then
'        aXMapper.InsertRows CStr(aSqlCmd), DOM.childNodes, list
'        SetNodes list.documentElement.childNodes
'    End If
'End Sub

Public Sub ExecInsertX(ConnectionString, aSqlCmd)
Dim list
    If Not (aXMapper Is Nothing) Then
        aXMapper.InsertRows ConnectionString, aSqlCmd, DOM.XML, list
        SetNodes2 list
    End If
End Sub
Public Sub ExecUpdate(aSqlCmd)
Dim list
'        Set list = UpdateRows(CStr(aSqlCmd), Dom.childNodes)
        SetNodes2 list
End Sub

Public Sub ExecUpdateX(ConnectionString, aSqlCmd)
Dim list 'As IXMLDOMNodeList
    If Not (aXMapper Is Nothing) Then
        aXMapper.UpdateRows ConnectionString, aSqlCmd, DOM.XML, list
        SetNodes2 list
    End If
End Sub

Public Function SetNodes2(list)
Dim anode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement
Dim aelm As IXMLDOMElement, bElm As IXMLDOMElement
Dim ares As Boolean
    RowCount = 0: ErrorCount = 0: ResultCount = 0
    ares = DOM.LoadXml(list)
    If Not ares Then
'        Open "c:\view11.txt" For Output As #1
'        Print #1, list
'        Close #1
        Do
            If DOM.parseError.errorCode <> -1072896760 Then Exit Do

            list = Left(list, DOM.parseError.filepos) & Right(list, Len(list) - DOM.parseError.filepos - 1)
'            Open "c:\view11.txt" For Output As #1
'            Print #1, list
'            Close #1
            ares = DOM.LoadXml(list)
        Loop
    End If
    ares = DOM.LoadXml(list)
    RowCount = DOM.documentElement.SelectNodes(".//ROW").length '   list.documentElement.getElementsByTagName("Row").length
    ErrorCount = DOM.documentElement.SelectNodes(".//ERROR").length

    If RowCount > 0 Then ActiveRow = 1 Else ActiveRow = -1
End Function

Public Function AddRow(Optional NewRowsCount) As Long
Dim anode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement
Dim i As Long
    If DOM Is Nothing Then Set DOM = New DOMDocument30
    If DOM.documentElement Is Nothing Then DOM.appendChild DOM.createElement("RESULT"): DOM.documentElement.appendChild DOM.createElement("ROWS")
    Set anode = DOM.documentElement.selectSingleNode("ROWS")
    Set bNode = anode.selectSingleNode("ROW")
    If IsMissing(NewRowsCount) Then NewRowsCount = 1
    If Not (bNode Is Nothing) Then
        For i = 1 To NewRowsCount
            Set cnode = bNode.cloneNode(True)
            anode.appendChild cnode
            ClearRow cnode
            RowCount = RowCount + 1
        Next i
    Else
        For i = 1 To NewRowsCount
            Set cnode = DOM.createElement("ROW")
            anode.appendChild cnode
            RowCount = 1
        Next i
    End If
    ActiveRow = RowCount
End Function

Public Function AddField(fieldname As String, Optional value As String)
Dim anode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement, dnode As IXMLDOMElement
Dim i As Long
    
    If DOM Is Nothing Then Set DOM = New DOMDocument30
    If DOM.documentElement Is Nothing Then DOM.appendChild DOM.createElement("RESULT"): DOM.documentElement.appendChild DOM.createElement("ROWS")
    Set anode = DOM.documentElement
    Set anode = anode.selectSingleNode("ROWS")
    
    Set bNode = anode.selectSingleNode("ROW")
    If Not (bNode Is Nothing) Then
        If (Not (bNode.selectSingleNode(UCase(fieldname)) Is Nothing)) Then
            MsgBox "To πεδίο " & UCase(fieldname) & " υπάρχει ήδη."
            Exit Function
        End If
    End If
    
    Set cnode = DOM.createElement(UCase(fieldname))
    If Not IsMissing(value) Then cnode.Text = value
        
    If (bNode Is Nothing) Then anode.appendChild DOM.createElement("ROW"): RowCount = 1: ActiveRow = RowCount
    For Each dnode In anode.childNodes
        If dnode.nodename = "ROW" Then dnode.appendChild cnode.cloneNode(True)
    Next dnode

End Function
Public Function AddNew(Optional NewRowsCount) As Long
AddNew = AddRow(NewRowsCount)
End Function

Public Function DropField(inFieldName As String) As Long
Dim anode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement
Dim i As Long, k As Long
    Set anode = DOM.documentElement '.selectSingleNode("ROWS")
    If (Not (anode Is Nothing)) And (anode.childNodes.length > 0) Then
        For Each bNode In anode.childNodes
            If bNode.nodename = "ROW" Then
                Set cnode = bNode.selectSingleNode(UCase(inFieldName))
                If Not cnode Is Nothing Then bNode.removeChild cnode
            End If
        Next bNode
    End If
End Function

Public Function DeleteRow(Optional DelRowNum As Integer) As Integer
Dim anode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement
Dim i As Long, k As Long
    DeleteRow = 0
    If IsMissing(DelRowNum) Then DelRowNum = 1
    If DOM Is Nothing Then Exit Function
    If DOM.documentElement Is Nothing Then Exit Function
    Set anode = DOM.documentElement.selectSingleNode("ROWS")
    If anode Is Nothing Then Exit Function
    i = 0:
    Do
    Set bNode = anode.childNodes(ActiveRow - 1)
    anode.removeChild bNode
    RowCount = RowCount - 1
    If ActiveRow > RowCount Then ActiveRow = RowCount
    i = i + 1
    If i >= DelRowNum Then Exit Do
    If RowCount = 0 Then Exit Do
    Loop
    DeleteRow = i
End Function
Public Function Delete(Optional DelRowNum As Integer) As Integer
    Delete = DeleteRow(DelRowNum)
End Function

Public Function ClearRow(inRow As IXMLDOMElement)
Dim anode As IXMLDOMElement
    For Each anode In inRow.childNodes
        anode.Text = ""
    Next anode
End Function

Public Function ErrDescription() As String
Dim anode As IXMLDOMElement
    Set anode = DOM.documentElement.selectSingleNode("Error")
    If anode Is Nothing Then
        ErrDescription = ""
    Else
        ErrDescription = anode.selectSingleNode("DESCRIPTION").Text
    End If
End Function

Public Sub SaveAsFile(inFileName As String)
    'DOM.save inFileName
    SaveXmlFile inFileName, DOM
End Sub

Public Function MoveFirst() As Long
    If RowCount > 0 Then ActiveRow = 1 Else ActiveRow = -1
    MoveFirst = ActiveRow
End Function

Public Function MoveLast() As Long
    If RowCount > 0 Then ActiveRow = RowCount Else ActiveRow = -1
    MoveLast = ActiveRow
End Function

Public Function MoveNext() As Long
    If ActiveRow < RowCount Then
        ActiveRow = ActiveRow + 1
        MoveNext = ActiveRow
    Else
        ActiveRow = -1: MoveNext = -1
    End If
End Function

Public Function MovePrevious() As Long
    If ActiveRow > 1 Then
        ActiveRow = ActiveRow - 1
        MovePrevious = ActiveRow
    Else
        MovePrevious = -1: ActiveRow = -1
    End If
End Function

Public Function AsStringByNo(inFieldNo As Long) As String
Dim anode As IXMLDOMElement
    If ActiveRow > 0 Then
        AsStringByNo = "": On Error Resume Next:
        Set anode = DOM.documentElement.childNodes(ActiveRow - 1).childNodes(inFieldNo)
        AsStringByNo = anode.Text
    Else
        AsStringByNo = ""
    End If
End Function

Public Property Get AsString(inFieldName As String) As String
Dim anode As IXMLDOMElement
    If ActiveRow > 0 Then
        AsString = "": On Error Resume Next:
        Set anode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
        If anode Is Nothing Then
            If FieldsExplicit Then MsgBox "Δεν υπάρχει το πεδίο: " & inFieldName
            AsString = ""
        Else
            AsString = anode.Text
        End If
    Else
        AsString = ""
    End If
End Property

Public Property Let AsString(inFieldName As String, value As String)
Dim anode As IXMLDOMElement
    If ActiveRow > 0 Then
        On Error Resume Next:
        Set anode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
        If anode Is Nothing Then
            If FieldsExplicit Then
                MsgBox "Δεν υπάρχει το πεδίο: " & inFieldName: Exit Property
            Else
                AddField (inFieldName)
                Set anode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
            End If
        End If
        anode.Text = value
    Else
    End If
End Property

Public Property Get fields(inFieldName As String) As XMLField
Dim anode As IXMLDOMElement
Dim bNode As IXMLDOMElement
On Error Resume Next:
Set anode = DOM.documentElement.selectSingleNode("ROWS")
Set anode = anode.childNodes(ActiveRow - 1)
Set bNode = anode.selectSingleNode(UCase(inFieldName))
If bNode Is Nothing Then
    'MsgBox "Δεν υπάρχει το πεδίο: " & inFieldName
    AddField (inFieldName)
    Set fields = fields(inFieldName): Exit Property
    'Set bNode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
    'Set Fields = aField
    'Set Fields.bNode = bNode
End If
'Else
Set fields = afield
Set fields.bNode = bNode

End Property

Public Property Get AsInteger(inFieldName As String) As Long
Dim astr As String
    astr = AsString(inFieldName): AsInteger = 0
    On Error Resume Next
    AsInteger = CLng(astr)
End Property

Public Property Get AsDouble(inFieldName As String) As Long
Dim astr As String
    astr = AsString(inFieldName): AsDouble = 0
    On Error Resume Next
    AsDouble = CDbl(astr)
End Property

Public Property Get AsDateU8(inFieldName As String) As String
Dim astr As String, apos As Integer, bpos As Integer
    astr = AsString(inFieldName): AsDateU8 = ""
    If astr = "01011900" Then Exit Property
    On Error Resume Next
    apos = InStr(1, astr, "/")
    bpos = InStr(apos + 1, astr, "/")
    AsDateU8 = Right("0" & Left(astr, apos - 1), 2) & Right("0" & Mid(astr, apos + 1, bpos - apos - 1), 2) & Right(astr, 4)
End Property

Public Property Let AsDateU8(inFieldName As String, value As String)
    AsString(inFieldName) = Left(value, 2) & "-" & Mid(value, 3, 2) & "-" & Right(value, 4)
End Property

Private Sub ResetRecordset()
    RowCount = -1
    ActiveRow = -1
End Sub

Private Sub Class_Initialize()
    FieldsExplicit = False
    ResetRecordset
    
    Set afield = New XMLField
'    Set Fields = New Collection
End Sub

Private Sub Class_Terminate()
'    ClearFields
    Set aSMapper = Nothing
    Set aXMapper = Nothing
    Set afield = Nothing
End Sub

Private Function ReportError(ByRef DOM As DOMDocument30, Optional addinfo As String)
Dim elm As IXMLDOMElement, Fld_lm As IXMLDOMElement

    Set elm = DOM.createElement("Error")
    DOM.documentElement.appendChild elm
    Set Fld_lm = DOM.createElement("Description"): Fld_lm.Text = Err.description
    elm.appendChild Fld_lm
    Set Fld_lm = DOM.createElement("ErrNo"): Fld_lm.Text = Err.Number
    elm.appendChild Fld_lm
    If Not IsMissing(addinfo) Then
        Set Fld_lm = DOM.createElement("AddInfo"): Fld_lm.Text = addinfo
        elm.appendChild Fld_lm
    End If
End Function

Private Function ReportConnectionError(DOM As DOMDocument30)
Dim elm As IXMLDOMElement, Fld_lm As IXMLDOMElement

    Set elm = DOM.createElement("Error")
    DOM.documentElement.appendChild elm
    Set Fld_lm = DOM.createElement("Description"): Fld_lm.Text = Err.description
    elm.appendChild Fld_lm
    Set Fld_lm = DOM.createElement("ErrNo"): Fld_lm.Text = Err.Number
    elm.appendChild Fld_lm
End Function


'Public Function UpdateRows(SqlStatement As String, Rows As IXMLDOMNodeList) As IXMLDOMNodeList
'
'Dim MainDB As New ADODB.Connection
'
'Dim ars As New ADODB.Recordset
'Dim Dom As New DOMDocument30
'Dim RootElm As IXMLDOMElement, Elm As IXMLDOMElement, Fld_lm As IXMLDOMElement
'Dim i As Long, k As Long
'Dim addinfo As String
''ares
'
'    MainDB.Open "File Name=" & "c:\VBTrade.udl" ', "sa", "sp750"
'
'    Set Dom.documentElement = Dom.createElement("ROWS")
''    If Failed Then
''        ReportConnectionError Dom
''    Else
'        On Error GoTo GenReportError
'        ars.Open SqlStatement, MainDB, adOpenKeyset, adLockOptimistic
'        If Rows.Length = 1 Then
'            Set RootElm = Rows.Item(0)
'        End If
'        If RootElm.childNodes.Length > 0 Then
'            ars.MoveFirst
'            For i = 0 To RootElm.childNodes.Length - 1
'
'                If RootElm.childNodes.Item(i).nodeName = "Row" Then
'                    Set Elm = RootElm.childNodes.Item(i)
'                    For k = 0 To Elm.childNodes.Length - 1
'                        addinfo = "Field Name: " & Elm.childNodes(k).nodeName & " Field Value: " & Elm.childNodes(k).Text
'
'                        If Trim(Elm.childNodes(k).Text) <> "" Then
'                            ars.Fields(Elm.childNodes(k).nodeName).Value = Elm.childNodes(k).Text
'                        Else
'                            ars.Fields(Elm.childNodes(k).nodeName).Value = Null
'                        End If
'
'                        addinfo = ""
'                    Next k
'                    addinfo = "Row: " & i
'                    ars.Update
'                    addinfo = ""
'                End If
'                Set Elm = Dom.createElement("RowUpdated")
'                Dom.documentElement.appendChild Elm
'                Elm.Text = CStr(i)
'
'                addinfo = "Row: " & i
'                ars.MoveNext
'                addinfo = ""
'            Next i
'        End If
''    End If
'    Set UpdateRows = Dom.childNodes
'Exit Function
'GenReportError:
'    ReportError Dom, addinfo
'    Set UpdateRows = Dom.childNodes
'End Function
'
'Public Sub DropFields()
'Dim i As Integer
'    While Fields.Count > 0
'        Fields.Remove 1
'    Wend
'End Sub
'
'Public Function DefineField(inFldName) As RSViewField
'Dim field As RSViewField
'    Set field = Fields(UCase(inFldName))
'    If field Is Nothing Then
'        Set field = New RSViewField
'        field.name = UCase(inFldName)
'        Fields.Add field, UCase(inFldName)
'    Else
'        MsgBox "Το Πεδίο " & inFldName & " υπάρχει ήδη"
'    End If
'    Set DefineField = field
'End Function
'

'Public Sub MapToFields()
'Dim node As IXMLDOMElement
'Dim field As RSViewField
'    If Dom Is Nothing Then Exit Sub
'    If Dom.documentElement Is Nothing Then Exit Sub
'
'    If Dom.documentElement.childNodes.length = 0 Then Exit Sub
'    If Fields.Count = 0 Then
'        For Each node In Dom.documentElement.childNodes(ActiveRow - 1).childNodes
'            Set field = New RSViewField
'            field.name = UCase(node.nodeName)
'            field.Value = node.Text
'            Fields.Add field, UCase(node.nodeName)
'        Next
'    Else
'        On Error Resume Next
'        For Each node In Dom.documentElement.childNodes(ActiveRow - 1).childNodes
'            Set field = Fields(UCase(node.nodeName))
'            If field Is Nothing Then MsgBox "Το Πεδίο " & node.nodeName & " δεν έχει δηλωθεί" _
'            Else field.Value = node.Text
'        Next
'    End If
'End Sub
'
'
'
'Public Function RowCount() As Long
'    If Dom Is Nothing Then RowCount = 0: Exit Function
'    If Dom.documentElement Is Nothing Then RowCount = 0: Exit Function
'    RowCount = Dom.documentElement.childNodes.length
'End Function



'
'Public Function SetNodes(list As IXMLDOMNodeList)
'Dim aNode As IXMLDOMElement, bNode As IXMLDOMElement, cnode As IXMLDOMElement
'Dim aElm As IXMLDOMElement, bElm As IXMLDOMElement
'    Set DOM.documentElement = DOM.createElement(list.Item(0).nodeName)
'    RowCount = 0: ErrorCount = 0: ResultCount = 0
'    For Each bNode In list.Item(0).childNodes
'        Set aElm = DOM.createElement(bNode.nodeName)
'
'        DOM.documentElement.appendChild aElm
'        If bNode.nodeName = "Row" Or bNode.nodeName = "Error" Then
'            If bNode.nodeName = "Row" Then RowCount = RowCount + 1
'            If bNode.nodeName = "Error" Then ErrorCount = ErrorCount + 1
'            If Not (bNode.childNodes Is Nothing) Then
'                For Each cnode In bNode.childNodes
'                    Set bElm = DOM.createElement(UCase(cnode.baseName))
'                    bElm.Text = cnode.Text
'                    aElm.appendChild bElm
'                Next cnode
'            Else
'                aElm.Text = bNode.Text
'            End If
'        ElseIf bNode.nodeName = "RowAdded" Or bNode.nodeName = "RowUpdated" Then
'            ResultCount = ResultCount + 1
'            aElm.Text = bNode.Text
'        End If
'    Next bNode
'    If RowCount > 0 Then ActiveRow = 1 Else ActiveRow = -1
'End Function
'



'Public Property Let Fields(inFieldName As String, inValue As XMLField)
'Dim aNode As IXMLDOMElement
'    If ActiveRow > 0 Then
'        On Error Resume Next:
'        Set aNode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
'        If aNode Is Nothing Then
'            If FieldsExplicit Then
'                MsgBox "Δεν υπάρχει το πεδίο: " & inFieldName: Exit Property
'            Else
'                AddField (inFieldName)
'                Set aNode = DOM.documentElement.childNodes(ActiveRow - 1).selectSingleNode(UCase(inFieldName))
'            End If
'        End If
'        aNode.Text = inValue.Value
'        aField.Name = inFieldName
'        aField.Value = inValue.Value
'    Else
'    End If
'End Property


