Attribute VB_Name = "WosaXFS"
Option Explicit

Public DocumentHandle As Long
Public JournalHandle As Long

Public Function WCleanup()
Dim Result As Long

DocumentHandle = 0
JournalHandle = 0

Result = GKWFSCleanUp()

WCleanup = Result
End Function

Public Function WClose(pWhere As String)

Dim Result As Long

If pWhere = "DOCUMENT" Then
    Result = GKWFSClose(DocumentHandle)
    DocumentHandle = 0
End If
If pWhere = "PASSBOOK" Then
    Result = GKWFSClose(DocumentHandle)
    DocumentHandle = 0
End If

'If pWhere = "JOURNAL" Then
'    Result = GKWFSClose(JournalHandle)
'    JournalHandle = 0
'End If

'MsgBox "WFSClose : " & Result

WClose = Result
End Function

Public Function WLock(pTimeOut As Long, pWhere As String) As Long
 
Dim Result As Long


If pWhere = "DOCUMENT" Then
    Result = GKWFSLock(pTimeOut, DocumentHandle)
End If

WLock = Result
End Function
Public Function WOpen(pTimeOut As Long, pWhere As String) As Long
 
Dim Result As Long
Dim Name As String

Result = GKWFSOpen("Document", pTimeOut, DocumentHandle)

WOpen = Result
End Function
Public Function WPrint(pData As String, _
                       pLenData As Long, _
                       pTimeOut As Long, _
                       pWhere As String) As Long
Dim Result As Long
 

Dim Data As String

Data = pData + Chr$(13) + Chr$(10)

pLenData = Len(pData) + 2
Data = Data & String(255 - pLenData, 0)

Result = GKWFSPrint(Data, pLenData, pTimeOut, DocumentHandle)


'MsgBox "WFSPrint : " & Result

WPrint = Result
End Function

Public Function WPrintForm(pForm As String, _
                            pMedia As String, _
                            pFields As String, _
                            pFormFeed As Long, _
                            pTimeOut As Long, _
                            pWhere As String) As Long
Dim Result As Long
Dim TimeOut As Long
Dim FormFeed As Long
Dim Form As String
Dim Media As String
Dim Fields As String
 


FormFeed = pFormFeed
TimeOut = pTimeOut
Form = pForm
Media = pMedia
Fields = pFields

Result = GKWFSPrintForm(Form, Media, Fields, FormFeed, TimeOut, DocumentHandle)
'MsgBox "WFSPrintForm : " & Result

WPrintForm = Result
End Function

Public Function WStart() As Long
Dim Versions As Long
Dim Result As Long
Versions = 0

DocumentHandle = 0
JournalHandle = 0

Result = GKWFSStartUp(Versions)
'MsgBox "WFSStartUp : " & Result

WStart = Result
End Function
Public Function WUnlock(pWhere As String)
Dim Result As Long

Result = GKWFSUnlock(DocumentHandle)

'MsgBox "WFSUnlock : " & Result

WUnlock = Result
End Function

Public Function WGetStatus(pTimeOut As Long, pWhere As String)
Dim Result As Long

Dim pDevice As Long
Dim pMedia As Long
Dim pPaper As Long

WGetStatus = 0
Result = 0
pDevice = 0
pMedia = 0
pPaper = 0

'εκτελείται δύο φορές γιατί όταν παρουσιάζεται λάθος και
'διορθώνεται, την πρώτη φορά εμφανίζει λάθος και
'την δεύτερη το σωστό
Result = GKWFSGetInfo(pDevice, pMedia, pPaper, pTimeOut, DocumentHandle)
Result = GKWFSGetInfo(pDevice, pMedia, pPaper, pTimeOut, DocumentHandle)

If Result <> 0 Then
    WGetStatus = -1
    Exit Function
End If


If pDevice <> 0 Then
    WGetStatus = WGetStatus + pDevice
End If

If pMedia <> 0 Then
    WGetStatus = WGetStatus + 10 * pMedia
End If

If pPaper <> 0 Then
    WGetStatus = WGetStatus + 100 * pPaper
End If

End Function

