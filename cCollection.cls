VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Collection_ As New Collection
Private Names As New msxml2.DOMDocument30

Private Sub Class_Initialize()
    Dim rootnode As IXMLDOMElement
    Set rootnode = Names.createElement("root")
    Names.appendChild rootnode
    Set rootnode = Nothing
End Sub

Public Sub add(item, Key As String)
    Dim newnode As IXMLDOMElement
    On Error GoTo InvalidNodeError
    Set newnode = Names.createElement(UCase(Key))
    Names.documentElement.appendChild newnode
    Collection_.add item, Key
    Set newnode = Nothing
    Exit Sub
InvalidNodeError:
    MsgBox "Λάθος στη δημιουργία Στοιχείου. Κλειδί: " & UCase(Key) & vbCrLf & Err.Number & " " & Err.description, vbOKOnly, "Λάθος Αρχικοποίησης(cCollection.Add)"
End Sub

Public Sub addunique(item, Key As String)
    Dim a As IXMLDOMNode
    Set a = Names.selectSingleNode("//" & UCase(Key))
    If a Is Nothing Then
        Dim newnode As IXMLDOMElement
        On Error GoTo InvalidNodeError
        Set newnode = Names.createElement(UCase(Key))
        Names.documentElement.appendChild newnode
        Collection_.add item, Key
        Set newnode = Nothing
    End If
    Exit Sub
    
InvalidNodeError:
    MsgBox "Λάθος στη δημιουργία Στοιχείου. Κλειδί: " & UCase(Key) & vbCrLf & Err.Number & " " & Err.description, vbOKOnly, "Λάθος Αρχικοποίησης(cCollection.Add)"

End Sub

Public Property Get item(name As String)
    If IsNumeric(name) Then
        Set item = Collection_.item(name)
    Else
        If Names.selectSingleNode("//root/" & UCase(name)) Is Nothing Then
            Set item = Nothing
        Else
            Set item = Collection_.item(name)
        End If
    End If
End Property

Public Sub update(value, name As String)
    If Not item(name) Is Nothing Then
        Remove (name)
        add value, name
    End If
End Sub

Public Sub Remove(name As String)
    If IsNumeric(name) Then
        Collection_.Remove name
    Else
        Dim elm As IXMLDOMElement
        Set elm = Names.selectSingleNode("//root/" & UCase(name))
        If elm Is Nothing Then
        Else
            Collection_.Remove name
            Names.documentElement.removeChild elm
        End If
        Set elm = Nothing
    End If
End Sub

Public Function count() As Long
    count = Collection_.count
End Function
