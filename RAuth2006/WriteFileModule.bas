Attribute VB_Name = "WriteFileModule"
Option Explicit

Public Sub SaveTextFile(filename As String, contents As String)
    Dim homedir As String
    homedir = NetworkHomeDir()

    On Error GoTo returnfalse

    If homedir <> "" Then
        Open homedir & "\" & filename For Output As #1
        Print #1, contents
        Close #1
    Else
        GoTo returnfalse
    End If

    Exit Sub
returnfalse:
    MsgBox "Η δημιουργία του αρχείου " & homedir & "\" & filename & " απέτυχε.", vbInformation

End Sub

Public Function ChkXmlFileExist(ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Dim homedir As String
    homedir = NetworkHomeDir()
    If homedir <> "" Then
         sFileName = homedir & "\" & sFileName
    End If
    
    Open sFileName For Input As #1
    Close #1
    
    ChkXmlFileExist = True
    Exit Function
returnfalse:
    ChkXmlFileExist = False
End Function

Public Function ChkXmlFileExistRemote(ByVal sFileName As String, ByVal ComputerName As String) As Boolean
On Error GoTo returnfalse
    Dim homedir As String
    homedir = WorkDir + "\" + ComputerName
    If homedir <> "" Then
         sFileName = homedir & "\" & sFileName
    End If
    
    Open sFileName For Input As #1
    Close #1
    
    ChkXmlFileExistRemote = True
    Exit Function
returnfalse:
    ChkXmlFileExistRemote = False
End Function

Public Sub SaveXmlFile(filename As String, xmldoc)
    Dim homedir As String
    homedir = NetworkHomeDir()
    On Error GoTo returnfalse
    If homedir <> "" Then
        xmldoc.Save homedir & "\" & filename
    Else
        GoTo returnfalse
    End If
    Exit Sub
returnfalse:
    MsgBox "Η δημιουργία του αρχείου " & filename & " απέτυχε.", vbInformation

End Sub

Public Function LoadXmlFile(location As String, filename As String) As MSXML2.DOMDocument30

    Dim homedir As String
   
    homedir = location 'NetworkHomeDir()
   
    On Error GoTo returnfalse
    Dim newdoc As New MSXML2.DOMDocument30
    
    If homedir <> "" Then
        newdoc.Load homedir & "\" & filename
    Else
        GoTo returnfalse
    End If
    Set LoadXmlFile = newdoc
    Exit Function
returnfalse:
    MsgBox "Η ανάκτηση του αρχείου " & filename & " απέτυχε.", vbInformation

End Function

Public Function LoadXmlFileRemote(filename As String, ComputerName As String) As MSXML2.DOMDocument30

    Dim homedir As String
   
    homedir = WorkDir + "\" + ComputerName
   
    On Error GoTo returnfalse
    Dim newdoc As New MSXML2.DOMDocument30
    
    If homedir <> "" Then
        newdoc.Load homedir & "\" & filename
    Else
        GoTo returnfalse
    End If
    Set LoadXmlFileRemote = newdoc
    Exit Function
returnfalse:
    MsgBox "Η ανάκτηση του αρχείου " & filename & " απέτυχε.", vbInformation

End Function


Public Function NetworkHomeDir() As String
    On Error GoTo creationError
    Dim fso
    Dim folder
    Dim foldername As String
    foldername = WorkDir & MachineName
    Set fso = CreateObject("Scripting.FileSystemObject")
    foldername = "\" & Replace(foldername, "\\", "\")
    If fso.FolderExists(foldername) Then
        Set folder = fso.GetFolder(foldername)
        NetworkHomeDir = foldername
        
        Set folder = Nothing
        Set fso = Nothing
        
        Exit Function
    Else
        Set folder = fso.CreateFolder(foldername)
        NetworkHomeDir = foldername
        
        Set folder = Nothing
        Set fso = Nothing
        
        Exit Function
    End If
    
    Set folder = Nothing
    Set fso = Nothing
    Exit Function
creationError:
     MsgBox "Η δημιουργία του  " & foldername & " απέτυχε.", vbInformation
End Function

Public Function RauthDir() As String
    On Error GoTo creationError
    Dim fso
    Dim folder
    Dim foldername As String
    foldername = WorkDir & "USERS"
    Set fso = CreateObject("Scripting.FileSystemObject")
    foldername = "\" & Replace(foldername, "\\", "\")
    If fso.FolderExists(foldername) Then
        Set folder = fso.GetFolder(foldername)
        RauthDir = foldername
        
        Set folder = Nothing
        Set fso = Nothing
        
        Exit Function
    Else
        Set folder = fso.CreateFolder(foldername)
        RauthDir = foldername
        
        Set folder = Nothing
        Set fso = Nothing
        
        Exit Function
    End If
    
    Set folder = Nothing
    Set fso = Nothing
    Exit Function
creationError:
     MsgBox "Η δημιουργία του  " & foldername & " απέτυχε.", vbInformation
End Function

Public Sub SaveXmlFileNew(location As String, filename As String, xmldoc)
    Dim homedir As String
    homedir = location
    On Error GoTo returnfalse
    If homedir <> "" Then
        xmldoc.Save homedir & "\" & filename
    Else
        GoTo returnfalse
    End If
    Exit Sub
returnfalse:
    MsgBox "Η δημιουργία του αρχείου " & filename & " απέτυχε.", vbInformation

End Sub

Public Function ChkXmlFileExistNew(ByVal location As String, ByVal sFileName As String) As Boolean
On Error GoTo returnfalse
    Dim homedir As String
    homedir = location 'NetworkHomeDir()
    If homedir <> "" Then
         sFileName = homedir & "\" & sFileName
    End If
    
    Open sFileName For Input As #1
    Close #1
    
    ChkXmlFileExistNew = True
    Exit Function
returnfalse:
    ChkXmlFileExistNew = False
End Function
