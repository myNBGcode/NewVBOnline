Attribute VB_Name = "MainModule"
Option Explicit

Dim ToDos(256) As Byte
Dim ToWin(256) As Byte

Public Sub Main()
Dim astr As String
Dim args() As String
     
    Dim i As Integer
    For i = 0 To 255
        ToDos(i) = i
        ToWin(i) = i
    Next i
    
    Dim set1 As String
    Dim set2 As String
    Dim set3 As String
    set1 = "ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÓÔÕÖ×ØÙáâãäåæçèéêëìíîïğñóòôõö÷ø"
    set2 = "ùÜİŞúßüıûş¢¸¹º¼¾¿"
    set3 = "ÚÛ"
    
    Dim startpos As Integer
    startpos = &H80
    For i = 1 To Len(set1)
        ToWin(startpos + i - 1) = Asc(Mid(set1, i, 1))
        ToDos(Asc(Mid(set1, i, 1))) = i + startpos - 1
    Next i
    startpos = &HE0
    For i = 1 To Len(set2)
        ToWin(startpos + i - 1) = Asc(Mid(set2, i, 1))
        ToDos(Asc(Mid(set2, i, 1))) = i + startpos - 1
    Next i
    startpos = &HF4
    For i = 1 To Len(set3)
        ToWin(startpos + i - 1) = Asc(Mid(set3, i, 1))
        ToDos(Asc(Mid(set3, i, 1))) = i + startpos - 1
    Next i
    
    
    astr = Command
    args = Split(astr, " ")
    
    Dim conversion As String
    Dim infilename As String
    Dim outfilename As String
    
    Dim arg
    For Each arg In args
        arg = Trim(UCase(arg)) & String(10, " ")
        If Left(arg, 2) = "C=" Then
            conversion = UCase(Trim(Right(arg, Len(arg) - 2)))
        ElseIf Left(arg, 2) = "I=" Then
            infilename = UCase(Trim(Right(arg, Len(arg) - 2)))
        ElseIf Left(arg, 2) = "O=" Then
            outfilename = UCase(Trim(Right(arg, Len(arg) - 2)))
        End If
    Next arg
    
    If (infilename <> "" And outfilename <> "" And conversion <> "") Then
        Dim fs As Scripting.FileSystemObject
        Set fs = CreateObject("Scripting.FileSystemObject")
        Dim inf As TextStream, outf As TextStream
        Set inf = fs.OpenTextFile(infilename, ForReading, False, TristateUseDefault)
        Set outf = fs.OpenTextFile(outfilename, ForWriting, True, TristateUseDefault)
        
        Dim aline As String
        While Not inf.AtEndOfStream
            aline = inf.ReadLine
            Dim newline As String
            newline = ""
            If UCase(Trim(conversion)) = "TODOS" Then
                While Len(aline) > 0
                    newline = newline & Chr(ToDos(Asc(Left(aline, 1))))
                    aline = Replace(aline, Left(aline, 1), "", 1, 1, vbBinaryCompare)
                Wend
                aline = newline
            ElseIf UCase(Trim(conversion)) = "TOWIN" Then
                While Len(aline) > 0
                    newline = newline & Chr(ToWin(Asc(Left(aline, 1))))
                    aline = Replace(aline, Left(aline, 1), "", 1, 1, vbBinaryCompare)
                Wend
                aline = newline
            End If
            outf.WriteLine aline
        Wend
        
        inf.Close
        outf.Close
        
    Else
        MsgBox "ChangeCodePage C=<conversion> I=<inputfilename> O=<outputfilename>" & vbCrLf & _
            "conversion: TOWIN, TODOS"
    End If




End Sub
