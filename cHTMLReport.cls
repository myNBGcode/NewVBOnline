VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTMLReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public controlOwner As GenBrowser
Public Control As WebBrowser
Attribute Control.VB_VarHelpID = -1

Public Sub PrepareFromGenBrowser(htmlControl As GenBrowser)
    Set controlOwner = htmlControl
    Set Control = htmlControl.Control
End Sub

Private Sub processreporttablecell(data As IXMLDOMElement, Row As HTMLTableRow, ByRef curcell As Long)
    Dim posattr As IXMLDOMAttribute
    Dim abspos As Long
    Dim workcell As HTMLTableCell
    abspos = curcell
    Set posattr = data.Attributes.getNamedItem("pos")
    If posattr Is Nothing Then
    Else
        If posattr.Value = "first" Then
            abspos = 0
        ElseIf posattr.Value = "last" Then
            abspos = Row.cells.Length - 1
        Else
            abspos = posattr.Value
        End If
    End If
    On Error GoTo erSelectingCell
    Set workcell = Row.cells(abspos)
    On Error GoTo erProcessingCell
    processsinglenode data, workcell
    curcell = abspos
    Exit Sub
erSelectingCell:
    MsgBox "Λάθος στην επιλογή πεδίου πίνακα, θέση " & curcell & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erProcessingCell:
    MsgBox "Λάθος στην επεξεργασία πεδίου πίνακα, θέση " & curcell & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
End Sub
    
Private Sub processreporttablecopyrow(oldrow As HTMLTableRow, newrow As HTMLTableRow)
    Dim i As Integer
    On Error GoTo erMergingRowAttributes
    newrow.mergeAttributes oldrow
    For i = 0 To oldrow.cells.Length - 1
        newrow.insertCell
        On Error GoTo erMergingCellAttributes
        newrow.cells(i).mergeAttributes oldrow.cells(i)
        On Error GoTo erCopyingCellHTML
        newrow.cells(i).innerHTML = oldrow.cells(i).innerHTML
    Next i
    Exit Sub
erMergingRowAttributes:
    MsgBox "Λάθος στην αντιγραφή ιδιοτήτων γραμμής πίνακα" & " " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erMergingCellAttributes:
    MsgBox "Λάθος στην αντιγραφή ιδιοτήτων πεδίου πίνακα" & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erCopyingCellHTML:
    MsgBox "Λάθος στην αντιγραφή HTML κώδικα πεδίου πίνακα" & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
End Sub

Private Sub processreporttablerow(data As IXMLDOMElement, table As HTMLTable)
    Dim abspos As Long, action As String
    Dim workrow As HTMLTableRow, oldrow As HTMLTableRow
    action = "update"
    Dim actionattr As IXMLDOMAttribute, posattr As IXMLDOMAttribute
    Set actionattr = data.Attributes.getNamedItem("action")
    Set posattr = data.Attributes.getNamedItem("pos")
    
    If posattr Is Nothing Then
        abspos = 0
    Else
        If posattr.Value = "first" Then
            abspos = 0
        ElseIf posattr.Value = "last" Then
            abspos = table.Rows.Length - 1
        Else
            abspos = posattr.Value
        End If
    End If
    If actionattr Is Nothing Then
    Else
        action = actionattr.Value
    End If
    On Error GoTo erSelectingRow
    Set oldrow = table.Rows(abspos)
    If action = "insertfirst" Then
        On Error GoTo erInsertingRow
        table.insertRow 0
        Set workrow = table.Rows(0)
        processreporttablecopyrow oldrow, workrow
        
    ElseIf action = "insertlast" Then
        On Error GoTo erInsertingRow
        table.insertRow table.Rows.Length
        Set workrow = table.Rows(table.Rows.Length - 1)
        processreporttablecopyrow oldrow, workrow
    
    ElseIf action = "insertbefore" Then
        On Error GoTo erInsertingRow
        table.insertRow abspos - 1
        Set workrow = table.Rows(table.Rows.Length - 1)
        processreporttablecopyrow oldrow, workrow
    
    ElseIf action = "insertafter" Then
        On Error GoTo erInsertingRow
        table.insertRow abspos + 1
        Set workrow = table.Rows(table.Rows.Length - 1)
        processreporttablecopyrow oldrow, workrow
    
    ElseIf action = "remove" Then
        table.DeleteRow abspos
    ElseIf action = "update" Then
        Set workrow = oldrow
    End If
    
    If action <> "remove" Then
        Dim cellnode As IXMLDOMNode, curcell As Long
        curcell = 0
        For Each cellnode In data.childNodes
            If cellnode.nodeType = NODE_ELEMENT Then
                If cellnode.baseName = "td" Then
                    On Error GoTo erProcessingCell
                    processreporttablecell cellnode, workrow, curcell
                    curcell = curcell + 1
                End If
            Else
            
            End If
        
        Next cellnode
    End If
    Exit Sub

erSelectingRow:
    MsgBox "Λάθος στην επιλογή γραμμής πίνακα " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erInsertingRow:
    MsgBox "Λάθος στην εισαγωγή γραμμής πίνακα " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erRemovingRow:
    MsgBox "Λάθος στη διαγραφή γραμμής πίνακα " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub
erProcessingCell:
    MsgBox "Λάθος στην επεξεργασία πεδίου πίνακα " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub

End Sub

Private Sub processreporttable(data As IXMLDOMElement, Report)
    Dim nameAttr As IXMLDOMAttribute
    Set nameAttr = data.Attributes.getNamedItem("name")
    If nameAttr Is Nothing Then
    Else
        Dim table As HTMLTable
        Set table = Report.All(nameAttr.Value)
        If table Is Nothing Then
        
        Else
            Dim rownode As IXMLDOMNode
            For Each rownode In data.childNodes
                If rownode.nodeType = NODE_ELEMENT Then
                    If rownode.baseName = "tr" Then
                        
                        processreporttablerow rownode, table
                        
                    End If
                End If
            Next rownode
        
        End If
    
    End If
    Exit Sub
erSelectingCell:
    MsgBox "Λάθος στην επεξεργασία πεδίου πίνακα " & " " & Err.Number & " " & Err.Description, vbCritical, "Λάθος"
    Exit Sub

End Sub


Private Sub processsinglenode(data As IXMLDOMNode, Report)
    Dim anode As IXMLDOMNode, foundflag As Boolean
    foundflag = False
    For Each anode In data.childNodes
        foundflag = True
        If anode.nodeType = NODE_ELEMENT Then
            If anode.baseName = "table" Then
                processreporttable anode, Report
            Else
            
            End If
        ElseIf anode.nodeType = NODE_TEXT Then
            Report.innerHTML = anode.Text
        End If
    Next anode
    If Not foundflag Then
        If data.nodeType = NODE_ELEMENT Then
            Report.innerHTML = data.Text
        End If
    End If
End Sub

Public Sub ProcessReport(filename As String, data As IXMLDOMNode)
    controlOwner.navigate ReadDir & "Reports\" & filename
    While controlOwner.inprogress
        
        DoEvents
    Wend
    processsinglenode data, controlOwner.Control.document
End Sub
