VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLAppicationCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Collection_ As New Collection

Public Sub add(item As String, Key As String)
    On Error GoTo InvalidDocError
    Collection_.add item, Key
    Exit Sub
InvalidDocError:
    MsgBox "Λάθος στη δημιουργία Στοιχείου. Κλειδί: " & UCase(Key) & vbCrLf & Err.Number & " " & Err.description, vbOKOnly, "Λάθος Αρχικοποίησης(cCollection.Add)"
End Sub

Public Property Get item(name As String)
    Set item = Collection_.item(name)
End Property

Public Function count() As Long
    count = Collection_.count
End Function

