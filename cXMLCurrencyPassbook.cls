VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLCurrencyPassbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public workDocument As MSXML2.DOMDocument30
Public Manager As cXMLDocumentManager
Public content As IXMLDOMElement
Public StaticLink As IXMLDOMElement

Dim currentline As Integer

Private Sub clearDoc()
    Dim i As Integer
    For i = 0 To DocumentLines - 1: DocLines(i) = String(255, " "): Next i
End Sub

Public Sub ProcessLineList(list As Collection)
    Dim linedata
    
    For Each linedata In list
    
        If currentline = 0 Then currentline = 1
        
        If currentline = 1 Then
            Dim nplinedata As String
            nplinedata = String(34, " ") & "EK MET " 'todo ypoloipo
            DocLines(currentline + 2) = " " & CStr(nplinedata)
            currentline = 2
        End If
        
        DocLines(currentline + 2) = " " & CStr(linedata)
        If currentline = 20 Then
            LastDocLine = currentline + 2
            'PrintDocLines_ Manager.activeform, "Εισαγωγή Βιβλιαρίου"
            xPrintDoc_ Nothing, "Εισαγωγή Βιβλιαρίου": xClearDoc_:
            clearDoc
            currentline = 1
        Else
            currentline = currentline + 1
        End If
    Next linedata
End Sub

Public Function LoadXml(inDocument As String) As String

    Dim ruleattr As IXMLDOMAttribute
    Dim rule As String
    Set ruleattr = StaticLink.Attributes.getNamedItem("selectrule")
    If ruleattr Is Nothing Then
        rule = "//BIBL_LINE[./BIBL_EIDOS != '']"
    Else
        rule = ruleattr.Text
    End If


    Set workDocument = New MSXML2.DOMDocument30
    workDocument.LoadXml inDocument
    
    Dim firstline As IXMLDOMAttribute
    If Not workDocument.documentElement Is Nothing Then
        Set firstline = workDocument.documentElement.Attributes.getNamedItem("firstline")
    End If
    If firstline Is Nothing Then currentline = 1 Else currentline = firstline.Text
    
    LoadXml = inDocument
    
    Dim list As IXMLDOMNodeList
    Dim entries As New Collection
    Set list = workDocument.SelectNodes(rule)
    
    If (list.length > 0) Then
        Dim Node As IXMLDOMNode
        For Each Node In list
            Dim entry As cCurrencyPassbookLine
            Set entry = New cCurrencyPassbookLine
            entry.PrepareEntry Node
            entries.add entry
        Next Node
        
        'Set CurrencyMessageForm.MessageDocument = workDocument
        'CurrencyMessageForm.Show vbModal, ActiveL2TrnHandler.activeform
    
        LoadXml = workDocument.XML
    End If
    
    For Each entry In entries
        Dim entrylines As Collection
        Set entrylines = entry.Lines
        ProcessLineList entrylines
    Next entry
    
    LastDocLine = currentline + 2
    'PrintDocLines_ Manager.activeform, "Εισαγωγή Βιβλιαρίου"
    xPrintDoc_ Nothing, "Εισαγωγή Βιβλιαρίου": xClearDoc_:
    
    Set list = workDocument.SelectNodes(rule)
    If list.length > 0 Then
    
    Else
    
    End If
    
End Function

