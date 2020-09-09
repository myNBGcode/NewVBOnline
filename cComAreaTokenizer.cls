VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComAreaTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public InitialData As String
Public Data As String

Private occurs1flag As Boolean

'ExtractXXXXXX         κάνει trim το data μετα το extraction
'ExtractXXXXXXnotrim   δεν κάνει trim το data μετα το extraction
'πουθενά αλλού δεν γίνεται trim

Public Function PrepareForFile(filename As String) As Boolean
    Dim structurecode As String
    structurecode = ""
    On Error Resume Next
    Close #1
    On Error GoTo GenError
    Dim s As String
    Open ReadDir & ComAreaDir & filename & ".txt" For Input As #1
    Do While Not Eof(1)
        Line Input #1, s
        structurecode = structurecode & s
    Loop
    Close #1
    
    InitialData = structurecode
    Data = structurecode
    PrepareForFile = True
    Exit Function
GenError:
    LogMsgbox "Δεν βρέθηκε το αρχείο της δομής: " & filename & " Err: " & CStr(Err.Number) & " - " & Err.description, vbCritical, "ΛΑΘΟΣ"
    Exit Function
End Function

Public Sub prepare(data_ As String)
    InitialData = data_
    Data = data_
End Sub

Public Function DetectToken(token As String) As Boolean
    If Len(token) <= Len(Data) Then
        If UCase(Left(Data, Len(token))) = UCase(token) Then
            DetectToken = True
        Else
            DetectToken = False
        End If
    Else
        DetectToken = False
    End If
End Function

Public Function DetectChar(charList As String) As Boolean
    If Len(Data) >= 1 Then
        If InStr(1, charList, Left(Data, 1)) > 0 Then
            DetectChar = True
        Else
            DetectChar = False
        End If
    Else
        DetectChar = False
    End If
End Function

Private Function ExtractCharnotrim(count As Integer) As String
    If Len(Data) < count Then
        Err.Raise 1200, "ComAreaTokenizer", "Δεν βρέθηκαν " & count & " χαρακτήρες στο:" & Data
    ElseIf Len(Data) = count Then
        ExtractCharnotrim = Data: Data = ""
    Else
        ExtractCharnotrim = Left(Data, count): Data = Right(Data, Len(Data) - count)
    End If
End Function

Public Function ExtractToken(token As String) As String
    ExtractToken = token
    
    If Len(token) = Len(Data) Then
        Data = ""
    Else
        If Not DetectToken(token) Then
            Err.Raise 1000, "ComAreaTokenizer", "Δεν βρέθηκε το στοιχείο " & token & " στο:" & Data
        End If
        Data = Trim(Right(Trim(Data), Len(Trim(Data)) - Len(token)))
    End If
End Function

Private Function ExtractUntilToken(token As String) As String
    Dim res As String, init As String
    init = Data
    res = ""
    While Data <> ""
        If DetectToken(token) Then
            ExtractUntilToken = res
            Data = Trim(Data)
            Exit Function
        Else
            res = res & ExtractCharnotrim(1)
        End If
    Wend
    Err.Raise 1100, "ComAreaTokenizer", "Δεν βρέθηκε το στοιχείο " & token & " στο:" & init
End Function

'----------------------------------------------------------------------------------------------------------
'**********************************************************************************************************
'----------------------------------------------------------------------------------------------------------

Private Function IsComment() As Boolean
    IsComment = DetectToken("/*")
End Function

Private Function ExtractComment() As String
    ExtractToken "/*"
    ExtractComment = ExtractUntilToken("*/")
    ExtractToken "*/"
End Function

Public Sub RemoveComments()
    While IsComment
        ExtractComment
    Wend
End Sub

Public Function IsLiteral() As Boolean
    IsLiteral = DetectToken("'")
End Function

Public Function ExtractLiteral() As String
    Dim res As String
    res = ""
    Do While IsLiteral
        res = res & ExtractCharnotrim(1)
        Do While Data <> ""
            If Left(Data, 1) = "'" Then
                res = res & ExtractCharnotrim(1)
                Exit Do
            Else
                res = res & ExtractCharnotrim(1)
            End If
        Loop
    Loop
    Data = Trim(Data)
    If res <> "" Then
    If Left(res, 1) = "'" Then
        If Len(res) > 1 Then res = Right(res, Len(res) - 1) Else res = ""
    End If
    End If
    If res <> "" Then
    If Right(res, 1) = "'" Then
        If Len(res) > 1 Then res = Left(res, Len(res) - 1) Else res = ""
    End If
    End If
    res = Replace(res, "''", "'")
    ExtractLiteral = res
End Function

Public Function IsDigit() As Boolean
    IsDigit = DetectChar("0123456789")
End Function

Public Function ExtractInteger() As String
    Dim res As String
    res = ""
    Do While IsDigit
        res = res & ExtractCharnotrim(1)
    Loop
    Data = Trim(Data)
    ExtractInteger = res
End Function

Public Function ExtractNumber() As String
    Dim res As String
    res = ""
    If DetectChar("0123456789+-.") Then
        res = res & ExtractCharnotrim(1)
        Do While DetectChar("0123456789.")
            res = res & ExtractCharnotrim(1)
            If Right(res, 1) = "." Then Exit Do
        Loop
        Do While IsDigit
            res = res & ExtractCharnotrim(1)
        Loop
    End If
    Data = Trim(Data)
    ExtractNumber = res
End Function

Public Function IsEndofField() As Boolean
    IsEndofField = DetectToken(",")
End Function

Public Function IsEndofStruct() As Boolean
    IsEndofStruct = DetectToken(";")
End Function

Public Function IsInclude() As Boolean
    IsInclude = DetectToken("%INCLUDE")
End Function

Public Function ExtractInclude() As String
    Dim res As String
    ExtractToken "%INCLUDE"
    ExtractInclude = ExtractUntilToken(";")
    ExtractToken ";"
    Data = Trim(Data)
End Function

Public Function ExtractName() As String
    Dim res As String
    res = ""
    Do While Data <> ""
        If InStr(1, " ,;()", Left(Data, 1)) > 0 Then
            ExtractName = res: Exit Do
        Else
            res = res & Left(Data, 1)
            If Len(Data) = 1 Then Data = "" Else Data = Right(Data, Len(Data) - 1)
        End If
    Loop
    Data = Trim(Data)
End Function

Private Function IsCharacter() As Boolean
    IsCharacter = DetectToken("CHARACTER") 'CHARACTER
End Function

Private Function IsChar() As Boolean
    IsChar = DetectToken("CHAR") 'CHARACTER
End Function

Private Function IsFixedDec() As Boolean
    IsFixedDec = DetectToken("FIXED DEC") And Not DetectToken("FIXED DECIMAL") 'FIXED DEC
End Function

Private Function IsFixedDecimal() As Boolean
    IsFixedDecimal = DetectToken("FIXED DECIMAL") 'FIXED DECIMAL
End Function

Private Function IsDecFixed() As Boolean
    IsDecFixed = DetectToken("DEC FIXED") 'DECIMAL FIXED
End Function

Private Function IsDecimalFixed() As Boolean
    IsDecimalFixed = DetectToken("DECIMAL FIXED") 'DEC FIXED
End Function

Private Function IsBinFixed() As Boolean
    IsBinFixed = DetectToken("BIN FIXED") 'BIN FIXED
End Function

Private Function IsBinaryFixed() As Boolean
    IsBinaryFixed = DetectToken("BINARY FIXED") 'BINARY FIXED
End Function

Private Function IsFixedBin() As Boolean
    IsFixedBin = DetectToken("FIXED BIN") And Not DetectToken("FIXED BINARY") 'FIXED BIN
End Function

Private Function IsFixedBinary() As Boolean
    IsFixedBinary = DetectToken("FIXED BINARY") 'FIXED BINARY
End Function

Public Function IsTypeDefinition() As Boolean
    IsTypeDefinition = IsChar Or IsFixedDec Or IsFixedDecimal Or IsDecFixed Or IsDecimalFixed Or IsBinFixed Or IsBinaryFixed Or IsFixedBin Or IsFixedBinary
End Function

Public Function ExtractTypeDefinition() As String
    If IsChar Then
        ExtractTypeDefinition = ExtractToken("CHAR")
    ElseIf IsFixedDec Then
        ExtractTypeDefinition = ExtractToken("FIXED DEC")
    ElseIf IsFixedDecimal Then
        ExtractTypeDefinition = ExtractToken("FIXED DECIMAL")
    ElseIf IsDecFixed Then
        ExtractTypeDefinition = ExtractToken("DEC FIXED")
    ElseIf IsDecimalFixed Then
        ExtractTypeDefinition = ExtractToken("DECIMAL FIXED")
    ElseIf IsBinFixed Then
        ExtractTypeDefinition = ExtractToken("BIN FIXED")
    ElseIf IsBinaryFixed Then
        ExtractTypeDefinition = ExtractToken("BINARY FIXED")
    ElseIf IsFixedBin Then
        ExtractTypeDefinition = ExtractToken("FIXED BIN")
    ElseIf IsFixedBinary Then
        ExtractTypeDefinition = ExtractToken("FIXED BINARY")
    End If
End Function

Public Function ExtractTypeSize() As Integer
    If DetectToken("(") Then
        ExtractToken "("
        ExtractTypeSize = ExtractInteger
        If DetectToken(",") Then
            ExtractToken ","
            ExtractInteger
        End If
        ExtractToken ")"
    Else
        ExtractTypeSize = 1
    End If
End Function

Public Function ExtractOccurs() As Integer
    occurs1flag = False
    If DetectToken("(") Then
        ExtractToken "("
        ExtractOccurs = CInt(ExtractInteger)
        ExtractToken ")"
        If ExtractOccurs = 1 Then occurs1flag = True
    Else
        ExtractOccurs = 1
    End If
End Function

Public Function IsPIC() As Boolean
    IsPIC = DetectToken("PIC")
End Function

Public Function ExtractPIC() As String
    Dim res As String
    res = ExtractToken("PIC")
    res = ExtractLiteral
    ExtractPIC = res
End Function

Public Function IsInit() As Boolean
    IsInit = DetectToken("INIT")
End Function

Public Function ExtractInit() As String
    If IsInit() Then
        ExtractToken "INIT"
        RemoveComments
        ExtractToken "("
        If IsLiteral Then
            Dim res As String
            ExtractInit = ExtractLiteral
        Else
            ExtractInit = ExtractNumber
        End If
        ExtractToken ")"
        RemoveComments
    End If
End Function

Public Function IsBased() As Boolean
    IsBased = DetectToken("BASED")
End Function

Public Function ExtractBased() As String
    Dim res As String
    ExtractToken "BASED"
    ExtractToken "("
    ExtractBased = ExtractName
    ExtractToken ")"
End Function

Public Function IsPTR() As Boolean
    IsPTR = DetectToken("PTR")
End Function

Public Function IsPOINTER() As Boolean
    IsPOINTER = DetectToken("POINTER")
End Function

Public Function IsSubStructure() As Boolean
    Dim back As String
    back = Data
    RemoveComments
    ExtractInteger
    ExtractName
    ExtractOccurs
    IsSubStructure = DetectToken(",")
    Data = back
End Function


Public Function ExtractField() As CComAreaField
    Set ExtractField = New CComAreaField
    RemoveComments
    ExtractField.Level = ExtractInteger
    ExtractField.name = ExtractName
    ExtractField.occurs = ExtractOccurs
    ExtractField.occurs1flag = occurs1flag
    
    If IsTypeDefinition Then
        ExtractField.datatype = ExtractTypeDefinition
        ExtractField.length = ExtractTypeSize
    End If
    If IsPIC Then ExtractField.pic = ExtractPIC: ExtractField.datatype = "CHAR":  ExtractField.length = Len(ExtractField.pic)
    If IsInit Then
        ExtractField.init = ExtractInit
    End If
    If DetectToken(",") Then ExtractToken ",": ExtractField.endofstruct = False
    If DetectToken(";") Then ExtractToken ";": ExtractField.endofstruct = True
    RemoveComments
End Function

Public Function ExtractSubStruct() As CComAreaField
    Set ExtractSubStruct = New CComAreaField
    RemoveComments
    ExtractSubStruct.Level = ExtractInteger
    ExtractSubStruct.name = ExtractName
    ExtractSubStruct.occurs = ExtractOccurs
    ExtractSubStruct.occurs1flag = occurs1flag
    If IsBased Then ExtractBased
    
    ExtractSubStruct.datatype = "SUBVIEW"
    ExtractSubStruct.endofstruct = False
    ExtractToken ","
    
    RemoveComments
    Set ExtractSubStruct.fields = New Collection
    Dim afielddesc As CComAreaField
    Dim initiallength As Integer, backupdata As String
    initiallength = 0

    While Data <> "" Or initiallength = Len(Data)
        initiallength = Len(Data)
        If IsEndofStruct Then
            ExtractSubStruct.endofstruct = False
            ExtractToken ";"
            Exit Function
        End If
        RemoveComments
        While Me.IsInclude
            Dim ainclude As String
            ainclude = ExtractInclude

            If ainclude <> "" Then
                Dim includetokenizer As New cComAreaTokenizer
                If Not includetokenizer.PrepareForFile(ainclude) Then

                Else
                    Data = Trim(includetokenizer.Data) & Data
                End If
                Data = Replace(Data, Chr(9), " ")
            Else

            End If
            RemoveComments
        Wend
        RemoveComments
        backupdata = Data
        
        If IsSubStructure Then
            Set afielddesc = Me.ExtractSubStruct
            If afielddesc.Level <= ExtractSubStruct.Level Then
                Data = backupdata:
                Exit Function
            End If
                           
            ExtractSubStruct.fields.add afielddesc
            If afielddesc.endofstruct Then
                Exit Function
            End If
        Else
            Set afielddesc = ExtractField
            If afielddesc.Level <= ExtractSubStruct.Level Then
                Data = backupdata:
                Exit Function
            End If
            ExtractSubStruct.fields.add afielddesc
            If afielddesc.endofstruct Then
                Exit Function
            End If
        End If
    Wend
    
End Function

'-----------------------------------------------------------------------------------------------------------
Public Function DetectStructDef(StructName As String) As Boolean
    Dim res As String
    Dim alevel As Integer
    Data = Trim(Data)
    On Error GoTo StructBuildFailed
    Do While Data <> ""
        RemoveComments
        ExtractToken "DCL"
        Dim backdata As String
        backdata = Data
        If IsDigit Then
            alevel = ExtractInteger
            res = ExtractName
            If res = StructName Then
                Data = backdata
                DetectStructDef = True
                Exit Function
            End If
        Else
            res = ExtractName
            If res = StructName Then
            
            Else
                If IsPTR Then
                    ExtractToken ("PTR")
                    ExtractToken ";"
                ElseIf IsPOINTER Then
                    ExtractToken ("POINTER")
                    ExtractToken ";"
                End If
            End If
        End If
    Loop
    Exit Function
StructBuildFailed:
    Err.Raise Err.Number, "CComAreaTokenizer.DetectStructDef", Err.description
    Exit Function
End Function

Public Sub ResolveOccurs1Flag(ProcessStruct As CComAreaField)
    Dim afield As CComAreaField
    Dim bfield As CComAreaField
    Dim foundflag As Boolean
    foundflag = False
    ProcessStruct.ResolveDataLength
    For Each afield In ProcessStruct.fields
        If afield.fields Is Nothing Then
        Else
            For Each bfield In afield.fields
                If bfield.occurs1flag Then
                    Dim aNum As Long
                    'aNum = (16250 - ProcessStruct.datalength) \ bfield.datalength
                    aNum = (32500 - ProcessStruct.DataLength) \ bfield.DataLength
                    bfield.occurs = bfield.occurs + aNum
                    ProcessStruct.ResolveDataLength
                    foundflag = True
                    Exit For
                End If
            Next bfield
            If foundflag Then Exit For
        End If
    Next afield
End Sub

Public Function ProcessStruct(StructureName As String) As CComAreaField
    On Error GoTo errInvalidStructureName
    If Not DetectStructDef(StructureName) Then
        GoTo errInvalidStructureName
    Else
        Set ProcessStruct = ExtractSubStruct
        ResolveOccurs1Flag ProcessStruct
    End If
    Exit Function
errInvalidStructureName:
    Set ProcessStruct = Nothing
    If Err.Number = 0 Then
        LogMsgbox "Λάθος Structure: " & StructureName, vbCritical, "Λάθος"
    Else
        LogMsgbox "Λάθος Structure: " & StructureName & vbCrLf & "Λάθος: " & Err.Number & vbCrLf & Err.description, vbCritical, "Λάθος"
    End If
    Exit Function
End Function

