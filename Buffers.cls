VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private BufferList() As Buffer
Private ComAreaSubViews As New Collection
Public BufferNum As Integer, name As String

Public Property Get ByIndex(Index) As Buffer
    Set ByIndex = BufferList(Index)
End Property

Public Function Exists(aname) As Boolean
Dim foundflag As Boolean, i As Integer, bufferpos As Integer
    Exists = False
    foundflag = False
    If BufferNum > 0 Then
        For i = 1 To BufferNum
            If Not (BufferList(i) Is Nothing) Then
                If BufferList(i).name = UCase(aname) Then
                    bufferpos = i: foundflag = True: Exit For
                End If
            End If
        Next i
    End If
    Exists = foundflag
End Function

Public Property Get ByName(aname) As Buffer
Dim foundflag As Boolean, i As Integer, bufferpos As Integer
    Set ByName = Nothing
    foundflag = False
    If BufferNum > 0 Then
        For i = 1 To BufferNum
            If Not (BufferList(i) Is Nothing) Then
                If BufferList(i).name = UCase(aname) Then
                    bufferpos = i: foundflag = True: Exit For
                End If
            End If
        Next i
    End If
    If Not foundflag Then LogMsgbox "ΛΑΘΟΣ η δομή: " & aname & " δεν υπάρχει", vbCritical, "ΛΑΘΟΣ": Exit Property
    Set ByName = BufferList(bufferpos)
End Property

Public Function AddBuffer(aStructName As String) As Integer
Dim foundflag As Boolean, i As Integer, bufferpos As Integer
    
    foundflag = False
    If BufferNum > 0 Then
        For i = 1 To BufferNum
            If Not (BufferList(i) Is Nothing) Then
                If BufferList(i).name = UCase(aStructName) Then
                    bufferpos = i: foundflag = True: Exit For
                End If
            End If
        Next i
        If foundflag Then
            LogMsgbox "ΛΑΘΟΣ η δομή: " & aStructName & " υπάρχει ήδη", vbCritical, "ΛΑΘΟΣ"
            Exit Function
        End If
    Else
        ReDim Preserve BufferList(1 To 1): BufferNum = 1: foundflag = True: bufferpos = 1
    End If
    If Not foundflag Then ReDim Preserve BufferList(1 To BufferNum + 1): bufferpos = BufferNum + 1: BufferNum = BufferNum + 1
    Set BufferList(bufferpos) = New Buffer
    AddBuffer = bufferpos

End Function

Public Function FreeBuffer(aStructName As String) As Integer
Dim foundflag As Boolean, i As Integer, bufferpos As Integer
    foundflag = False
    If BufferNum > 0 Then
        For i = 1 To BufferNum
            If Not (BufferList(i) Is Nothing) Then
                If BufferList(i).name = UCase(aStructName) Then
                    bufferpos = i: foundflag = True: Exit For
                End If
            End If
        Next i
    End If
    If Not foundflag Then LogMsgbox "ΛΑΘΟΣ η δομή: " & aStructName & " δεν υπάρχει", vbCritical, "ΛΑΘΟΣ": Exit Function
    BufferList(bufferpos).Clear
    
    Set BufferList(bufferpos) = Nothing
    If BufferNum > 1 Then
        While BufferList(BufferNum) Is Nothing And BufferNum > 1
            ReDim Preserve BufferList(1 To BufferNum - 1): BufferNum = BufferNum - 1
        Wend
    End If
End Function


Private Function BuildBufferFromComAreaField(aBuffer As Buffer, afield As CComAreaField)

Dim SubStruct As Buffer
Dim DataPart As String
Dim bufferpos As Integer
    
Dim aname As String
Dim aType As Integer
Dim alength As Long
Dim aoccur As Long
Dim newbuffer As Buffer
Dim bfield As CComAreaField
    
    Set newbuffer = Nothing
    aname = afield.name
    If afield.datatype = "CHAR" Then
        alength = afield.length: aType = ptChar
        If afield.init = "" Then
            DataPart = AsciiToEbcdic_(String(alength, " "))
        Else
            DataPart = Left(AsciiToEbcdic_(afield.init) & AsciiToEbcdic_(String(alength, " ")), alength)
        End If
    ElseIf afield.datatype = "FIXED DEC" Or afield.datatype = "DEC FIXED" Or afield.datatype = "DECIMAL FIXED" Or afield.datatype = "FIXED DECIMAL" Then
        alength = afield.length: aType = ptDecimal
        DataPart = DecimalToHPS_(0, alength): alength = Len(DataPart)
    ElseIf afield.datatype = "FIXED BIN" Or afield.datatype = "BIN FIXED" Then
        alength = afield.length
        If alength = 31 Then
            alength = 4: aType = ptInt
            DataPart = IntToHps_(0)
        ElseIf alength = 15 Then
            alength = 2: aType = ptSmall
            DataPart = SmallToHps_(0)
        Else
            Err.Raise 1300, "cComAreaTokenizer", "Λάθος μέγεθος " & alength & " για το τύπο FIXED BIN"
        End If
    ElseIf afield.datatype = "SUBVIEW" Then
        'If Not (abuffer Is Nothing) Then aname = abuffer.name & "_" & aname
        If Not Exists(aname) Then
            Set newbuffer = New Buffer
            
            newbuffer.name = UCase(aname)
            newbuffer.StructID = UCase(aname)
            newbuffer.BuffType = aname
            newbuffer.LastLevel = True
            
            For Each bfield In afield.fields
                BuildBufferFromComAreaField newbuffer, bfield
            Next
            Set bfield = Nothing
            
            newbuffer.ClearData_ = newbuffer.data_
            
            Dim VName As String
            VName = aname
            VName = Replace(VName, ":", "_")
            VName = Replace(VName, " ", "_")
            If IsNumeric(Left(VName, 1)) Then VName = "_" & VName
            If aBuffer Is Nothing Then
                LoadStructXMLDoc newbuffer
                ReDim Preserve BufferList(1 To BufferNum + 1): bufferpos = BufferNum + 1: BufferNum = BufferNum + 1
                Set BufferList(bufferpos) = newbuffer
            Else
                ComAreaSubViews.add newbuffer
            End If
            
        Else
            Set newbuffer = ByName(aname)
        End If
        Set BuildBufferFromComAreaField = newbuffer
        
        alength = newbuffer.length
        DataPart = newbuffer.ClearData_
        aType = ptStruct
        
    End If
    aoccur = afield.occurs
    
    If Not aBuffer Is Nothing Then
        Set BuildBufferFromComAreaField = aBuffer.NewPart(Me, aname, aType, aoccur, alength, DataPart, aname, aname, newbuffer)
    End If
    
End Function


Public Function DefineComArea(structurecode As String, StructureName As String, hidden As Boolean) As Buffer
    Dim atokenizer As New cComAreaTokenizer
    atokenizer.prepare structurecode
    
    Dim afield As CComAreaField
    
    Set afield = atokenizer.ProcessStruct(StructureName)
    
    If afield Is Nothing Then
        Set DefineComArea = Nothing
    Else
        Set DefineComArea = BuildBufferFromComAreaField(Nothing, afield)
    End If
    
    afield.CleanUp
    
    Exit Function
    
errInvalidStructureName:
    LogMsgbox "Λάθος Structure: " & StructureName, vbCritical, "Λάθος"
    Exit Function
    
End Function

Public Function DefineBuffer(aStructName As String, astructid As String, aStructDesc As String, Optional astructtype As String, Optional aLastLevel As Boolean) As Integer
Dim StructDesc() As String, SubStruct As Buffer
Dim inDesc As String, DescPart As String, partname As String, PartTypeName As String, DataPart As String
Dim PartLenghtString  As String, PartTimesString As String
Dim parttype As Integer, PartLength As Long, PartTimes As Long, StructType As String
Dim apos As Long, bufferpos As Integer, foundflag As Boolean
Dim i As Long, k As Long
    
    DefineBuffer = -1
    foundflag = False
    If BufferNum > 0 Then
        For i = 1 To BufferNum
            If Not (BufferList(i) Is Nothing) Then
                If BufferList(i).name = UCase(aStructName) Then
                    bufferpos = i: foundflag = True: Exit For
                End If
            End If
        Next i
        If foundflag Then
            LogMsgbox "ΛΑΘΟΣ η δομή: " & aStructName & " υπάρχει ήδη", vbCritical, "ΛΑΘΟΣ"
            Exit Function
        End If
    Else
        ReDim Preserve BufferList(1 To 1): BufferNum = 1: foundflag = True: bufferpos = 1
    End If
    If Not foundflag Then ReDim Preserve BufferList(1 To BufferNum + 1): bufferpos = BufferNum + 1: BufferNum = BufferNum + 1
    
    If IsMissing(astructtype) Or astructtype = "" Then astructtype = UCase(aStructName)
    If IsMissing(aLastLevel) Then aLastLevel = True
    Set BufferList(bufferpos) = New Buffer
    
    BufferList(bufferpos).name = UCase(aStructName)
    BufferList(bufferpos).StructID = UCase(astructid)
    BufferList(bufferpos).BuffType = astructtype
    BufferList(bufferpos).Definition = aStructDesc
    BufferList(bufferpos).LastLevel = aLastLevel
    
    DefineBuffer = bufferpos
    
    inDesc = aStructDesc
    While inDesc <> ""
        apos = InStr(1, inDesc, ",")
        If apos = 0 Then apos = Len(inDesc) + 1
        DescPart = Trim(Left(inDesc, apos - 1))
        If apos > Len(inDesc) Then inDesc = "" Else inDesc = Right(inDesc, Len(inDesc) - apos)
        apos = InStr(1, DescPart, " ")
        If apos = 0 Then
            LogMsgbox "ΛΑΘΟΣ στον ορισμό δομής: " & aStructName, vbCritical, "ΛΑΘΟΣ"
            Exit Function
        End If
        partname = Left(DescPart, apos - 1)
        PartTypeName = UCase(Right(DescPart, Len(DescPart) - apos))
        Set SubStruct = Nothing
        
        If Left(PartTypeName, 3) = "INT" Then
            PartTypeName = Right(PartTypeName, Len(PartTypeName) - 3)
            PartLength = 4: parttype = ptInt
            DataPart = IntToHps_(0)
        ElseIf Left(PartTypeName, 9) = UCase("timestamp") Then
            PartTypeName = Right(PartTypeName, Len(PartTypeName) - 9)
            PartLength = 12: parttype = ptTimeStamp
            DataPart = String(12, Chr(0))
        ElseIf Left(PartTypeName, 7) = UCase("picture") Then
            PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - 7))
            apos = InStr(1, PartTypeName, " ")
            If apos = 0 Then apos = Len(PartTypeName) + 1
            PartLength = CInt(Trim(Left(PartTypeName, apos - 1))): parttype = ptPicture
            If apos > Len(PartTypeName) Then PartTypeName = "" Else PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - apos))
            DataPart = AsciiToEbcdic_(String(PartLength, " "))
        ElseIf Left(PartTypeName, 4) = "TIME" Then
            PartTypeName = Right(PartTypeName, Len(PartTypeName) - 4)
            If OpenCobolServer Then
                PartLength = 12: parttype = ptTime
                DataPart = AsciiToEbcdic_("00:00:00:000")
            Else
                PartLength = 4: parttype = ptTime
                DataPart = IntToHps_(0)
            End If
        ElseIf Left(PartTypeName, 2) = "FD" Then
            PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - 2))
            apos = InStr(1, PartTypeName, " ")
            If apos = 0 Then apos = Len(PartTypeName) + 1
            PartLength = CInt(Trim(Left(PartTypeName, apos - 1))): parttype = ptDecimal
            If apos > Len(PartTypeName) Then PartTypeName = "" Else PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - apos))
            DataPart = DecimalToHPS_(0, PartLength): PartLength = Len(DataPart)
        ElseIf Left(PartTypeName, 5) = "SMALL" Then
            PartTypeName = Right(PartTypeName, Len(PartTypeName) - 5)
            PartLength = 2: parttype = ptSmall
            DataPart = SmallToHps_(0)
        ElseIf Left(PartTypeName, 4) = "CHAR" Then
            PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - 4))
            apos = InStr(1, PartTypeName, " ")
            If apos = 0 Then apos = Len(PartTypeName) + 1
            PartLength = CInt(Trim(Left(PartTypeName, apos - 1))): parttype = ptChar
            If apos > Len(PartTypeName) Then PartTypeName = "" Else PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - apos))
            DataPart = AsciiToEbcdic_(String(PartLength, " "))
        ElseIf Left(PartTypeName, 7) = "VARCHAR" Then
            PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - 7))
            apos = InStr(1, PartTypeName, " ")
            If apos = 0 Then apos = Len(PartTypeName) + 1
            PartLength = CInt(Trim(Left(PartTypeName, apos - 1))): parttype = ptVarchar
            If apos > Len(PartTypeName) Then PartTypeName = "" Else PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - apos))
            DataPart = VarCharToHps_(String(PartLength, " "), PartLength): PartLength = Len(DataPart)
        ElseIf Left(PartTypeName, 4) = "DATE" Then
            PartTypeName = Right(PartTypeName, Len(PartTypeName) - 4)
            If OpenCobolServer Then
                PartLength = 10: parttype = ptDate
                DataPart = AsciiToEbcdic_("0001-01-01")
            Else
                PartLength = 4: parttype = ptDate
                DataPart = Chr(0) & Chr(0) & Chr(0) & Chr(0)
            End If
        ElseIf Left(PartTypeName, 6) = "STRUCT" Then
            PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - 6))
            
            apos = InStr(1, PartTypeName, " ")
            If apos = 0 Then apos = Len(PartTypeName) + 1
            StructType = Trim(Left(PartTypeName, apos - 1))
            If apos > Len(PartTypeName) Then PartTypeName = "" Else PartTypeName = Trim(Right(PartTypeName, Len(PartTypeName) - apos))
            Set SubStruct = ByName(StructType)
            
            With SubStruct
                PartLength = .length
                DataPart = .ClearData_
            End With
            'Το DataPart θα δημιουργηθεί απο τα substructures
            parttype = ptStruct
        End If
            
        If PartTypeName = "" Then PartTimes = 1 Else PartTimes = CInt(PartTypeName)
        
        BufferList(bufferpos).NewPart Me, partname, parttype, PartTimes, PartLength, DataPart, aStructName, StructType, SubStruct
    Wend
    BufferList(bufferpos).ClearData_ = BufferList(bufferpos).data_
    
    Dim VName As String
    VName = BufferList(bufferpos).name
    VName = Replace(VName, ":", "_")
    VName = Replace(VName, " ", "_")
    If IsNumeric(Left(VName, 1)) Then VName = "_" & VName
    
    LoadStructXMLDoc BufferList(bufferpos)
    'BufferList(BufferPos).xmlDoc.Save "C:\temp\views\" & VName & ".xml"
End Function

Private Sub LoadStructXMLDoc(aStruct As Buffer)
    
    Dim VName As String
    VName = aStruct.name
    VName = Replace(VName, ":", "_")
    VName = Replace(VName, " ", "_")
    If IsNumeric(Left(VName, 1)) Then VName = "_" & VName
'    aStruct.xmlDoc.loadXML "<" & VName & ">" & aStruct.xmlStr & "</" & VName & ">"
'    If aStruct.xmlDoc.parseError.errorCode <> 0 Then MsgBox "ΛΑΘΟΣ ΣΤΗ ΔΗΜΙΟΥΡΓΙΑ ΤΟΥ " & VName & " XML"
    
    aStruct.xmlDocV2.LoadXml "<" & VName & ">" & aStruct.xmlStrV2 & "</" & VName & ">"
    If aStruct.xmlDocV2.parseError.errorCode <> 0 Then LogMsgbox "ΛΑΘΟΣ ΣΤΗ ΔΗΜΙΟΥΡΓΙΑ ΤΟΥ " & VName & " XML Ver 2"
    Set aStruct.localnode = aStruct.xmlDocV2.documentElement
'    aStruct.xmlDocV2.Save "c:\tmp3\" & VName & ".xml"
'    BufferList(BufferPos).xmlDoc.Save "C:\temp\views\" & VName & ".xml"
    
End Sub

Public Sub ClearAll()
    Dim i As Long
    i = 1
    While i <= BufferNum
        If Not (BufferList(i) Is Nothing) Then BufferList(i).Clear: Set BufferList(i) = Nothing
        i = i + 1
    Wend
    ReDim BufferList(0): BufferNum = 0
    
    While ComAreaSubViews.count > 0
        'ComAreaSubViews(ComAreaSubViews.count).Clear
        'ComAreaSubViews.Remove ComAreaSubViews.count
        ComAreaSubViews(1).Clear
        ComAreaSubViews.Remove 1
    Wend
End Sub

Public Sub SetPart(aStruct As String, apart As String, position As Long, avalue)
    On Error Resume Next
    ByName(aStruct).ByName(apart).value = avalue
End Sub

Public Function GetPart(aStruct As String, apart As String, position As Long)
    On Error Resume Next
    GetPart = ByName(aStruct).ByName(apart).value
End Function

Public Function GetInPart(aStruct As String, apart As String, position As Long)
Dim aSlice As Long
    On Error Resume Next
    With ByName(aStruct).ByName(apart)
        aSlice = Len(.data) \ .Times
        GetInPart = Mid(.data, (position - 1) * aSlice + 1, aSlice)
    End With
End Function

Public Sub SetInPart(aStruct As String, apart As String, avalue As String)
    On Error Resume Next
    With ByName(aStruct).ByName(apart)
        .data = Left(avalue & String(Len(.data), " "), Len(.data))
    End With
End Sub

Public Function GetIn(aStruct As String) As String
    On Error Resume Next
    GetIn = ByName(aStruct).data
End Function

Public Sub SetIn(aStruct As String, avalue As String)
    On Error Resume Next
    ByName(aStruct).data = avalue
End Sub

