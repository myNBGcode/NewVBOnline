VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public name As String, datatype As Integer, times As Double, owner As Buffers
Public nStartPos_ As Long, nLength_ As Long, ownerBuffer As Buffer
Public SubStruct As Buffer
Public localnode As MSXML2.IXMLDOMElement


Public Sub ClearData()
Dim astr As String
    
    astr = Mid(ownerBuffer.ClearData_, nStartPos_, nLength_)
    
    ownerBuffer.data_ = IIf(nStartPos_ > 1, Left(ownerBuffer.data_, nStartPos_ - 1), "") & astr & _
        IIf(ownerBuffer.length - nStartPos_ - nLength_ > 0, Right(ownerBuffer.data_, ownerBuffer.length - nStartPos_ + 1 - nLength_), "")
    
End Sub

Public Property Get value(Optional position As Long)
Dim aSlice As Long, astr As String
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1

    aSlice = nLength_ \ times
    
    astr = Mid(ownerBuffer.data_, nStartPos_ + (position - 1) * aSlice, aSlice)
    If datatype = ptChar Then
        value = EbcdicToAscii_(astr)
    ElseIf datatype = ptPicture Then
        value = EbcdicToAscii_(astr)
    ElseIf datatype = ptVarchar Then
        value = EbcdicToAscii_(Right(astr, aSlice - 2))
    ElseIf datatype = ptInt Then
        value = HpsToInt_(IIf(astr <> "", astr, String(4, Chr(0))))
    ElseIf datatype = ptTime Then
        If OpenCobolServer Then
            Dim ahh As Long
            Dim amm As Long
            Dim ass As Long
            Dim amss As Long
            
            astr = EbcdicToAscii_(astr)
            ahh = Left(astr, 2)
            amm = Mid(astr, 4, 2)
            ass = Mid(astr, 7, 2)
            amss = Right(astr, 2)
            value = ((ahh * 60 + amm) * 60 + ass) * 1000 + amss
        Else
            value = HpsToInt_(IIf(astr <> "", astr, String(4, Chr(0))))
        End If
    ElseIf datatype = ptSmall Then
        value = HpsToSmall_(IIf(astr <> "", astr, String(4, Chr(0))))
    ElseIf datatype = ptDecimal Then
        value = HPSToDecimal_(IIf(astr <> "", astr, String(4, Chr(0))))
    ElseIf datatype = ptTimeStamp Then
        value = HpsToInt_(Left(astr, 4)) & "," & HpsToInt_(Mid(astr, 5, 4)) & "," & HpsToInt_(Right(astr, 4))
    ElseIf datatype = ptDate Then
        If OpenCobolServer Then
            value = EbcdicToAscii_(astr)
            value = DateSerial(Left(value, 4), Mid(value, 6, 2), Right(value, 2))
        Else
            value = DateAdd("y", IIf(HpsToInt_(astr) >= 36526, HpsToInt_(astr) - 36526, 0), #1/1/100#)
        End If
    End If
End Property

Public Property Get FormatedDate8(Optional position As Long)
Dim aSlice As Long, value As Date, astr As String
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1

    aSlice = nLength_ \ times
    
    astr = Mid(ownerBuffer.data_, nStartPos_ + (position - 1) * aSlice, aSlice)
    If datatype = ptDate Then
        'value = DateAdd("y", IIf(HpsToInt_(astr) >= 36526, HpsToInt_(astr) - 36526, 0), #1/1/100#)
        If OpenCobolServer Then
            astr = EbcdicToAscii_(astr)
            value = DateSerial(Left(astr, 4), Mid(astr, 6, 2), Right(astr, 2))
        Else
            value = DateAdd("y", IIf(HpsToInt_(astr) >= 36526, HpsToInt_(astr) - 36526, 0), #1/1/100#)
        End If
    End If
    If value > DateSerial(1900, 1, 1) Then
        astr = Right("00" & CStr(Day(value)), 2) & Right("00" & CStr(Month(value)), 2) & Right("0000" & CStr(Year(value)), 4)
        FormatedDate8 = format(astr, "00/00/0000")
    Else
        FormatedDate8 = ""
    End If
End Property

Public Property Let value(Optional position As Long, aValue)
    Dim astr As String, aSlice As Long, astart As Long
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1
    
    aSlice = nLength_ \ times
    
    If datatype = ptChar Then
        astr = AsciiToEbcdic_(Left(CStr(aValue) & String(aSlice, " "), aSlice))
    ElseIf datatype = ptPicture Then
        astr = AsciiToEbcdic_(Left(CStr(aValue) & String(aSlice, " "), aSlice))
    ElseIf datatype = ptVarchar Then
        astr = VarCharToHps_(CStr(aValue), aSlice - 2)
    ElseIf datatype = ptInt Then
        astr = IntToHps_(CLng(aValue))
        If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
    ElseIf datatype = ptTime Then
        If OpenCobolServer Then
            Dim ahh As Long
            Dim amm As Long
            Dim ass As Long
            Dim amss As Long
            
            amss = aValue Mod 1000
            value = aValue \ 1000
            ass = aValue Mod 60
            value = aValue \ 60
            amm = aValue Mod 60
            value = aValue \ 60
            ahh = aValue Mod 60
            
            astr = Right("00" & ahh, 2) & ":" & Right("00" & amm, 2) & ":" & Right("00" & ass, 2) & ":" & Right("000" & amss, 3)
            astr = AsciiToEbcdic_(astr)
        Else
            astr = IntToHps_(CLng(aValue))
            If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
        End If
    ElseIf datatype = ptSmall Then
        astr = SmallToHps_(CLng(aValue))
        If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
    ElseIf datatype = ptDecimal Then
        astr = DecimalToHPS_(CDbl(IIf(Trim(aValue) <> "", aValue, 0)), aSlice * 2 - 1)
        If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
    ElseIf datatype = ptTimeStamp Then
        astr = String(12, Chr(0))
        Dim splitStr As Variant
        splitStr = Split(aValue, ",")
        If (UBound(splitStr) = 2) Then
          astr = IntToHps_(CLng(splitStr(0))) & IntToHps_(CLng(splitStr(1))) & IntToHps_(CLng(splitStr(2)))
        End If
    ElseIf datatype = ptDate Then
        If OpenCobolServer Then
            If Not IsDate(aValue) Then
                astr = Right(aValue, 4) + "-" + Mid(aValue, 3, 2) + "-" + Mid(aValue, 1, 2)
            Else
                astr = Right("0000" & Year(aValue), 4) & "-" & Right("00" & Month(aValue), 2) & "-" & Right("00" & Day(aValue), 2)
            End If
            If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
            astr = AsciiToEbcdic_(astr)
        Else
            If Not IsDate(aValue) Then
                aValue = Mid(aValue, 1, 2) + "/" + Mid(aValue, 3, 2) + "/" + Right(aValue, 4)
            End If
            If aValue = "01010001" Or aValue = "01/01/0001" Or aValue = "0001-01-01" Or aValue = "01.01.0001" Or aValue = "0001.01.01" Then
                astr = IntToHps_(367)
            Else
                If CDate(aValue) = DateSerial(1900, 1, 1) Then
                    astr = IntToHps_(0)
'                ElseIf CDate(aValue) = DateSerial(1, 1, 1) Then
'                    astr = IntToHps_(367)
                Else
                    astr = IntToHps_(36526 + DateDiff("y", #1/1/100#, CDate(aValue)))
                End If
            End If
            If Len(astr) > nLength_ Then astr = Left(astr, nLength_)
        End If
    End If
    astart = nStartPos_ + (position - 1) * aSlice
    If astart - 1 + aSlice = ownerBuffer.length Then
        ownerBuffer.data_ = Left(ownerBuffer.data_, ownerBuffer.length - aSlice) + astr
    Else
        ownerBuffer.data_ = IIf(astart > 1, Left(ownerBuffer.data_, astart - 1), "") & astr & _
            IIf(ownerBuffer.length - astart + 1 - aSlice > 0, Right(ownerBuffer.data_, ownerBuffer.length - astart + 1 - aSlice), "")
    End If
    'astr = value
End Property


'Public Property Get Length(aValue) As Long
'    If dataType = ptStruct Then Length = SubStruct.Length Else Length = Length_
'End Property
'
Public Property Get Data(Optional position) As String
Dim i As Long, aSlice As Long
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1
    aSlice = nLength_ \ times
    
    Data = Mid(ownerBuffer.data_, nStartPos_ + (position - 1) * aSlice, aSlice)
    
End Property

Public Property Let Data(Optional position, invalue As String)
Dim i As Long, aSlice As Long, astart As Long
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1
    aSlice = nLength_ \ times
    
    astart = nStartPos_ + (position - 1) * aSlice
    
        Dim aLeftPart As String, aRightPart As String
        If astart > 1 Then aLeftPart = Left(ownerBuffer.data_, astart - 1) Else aLeftPart = ""
        If astart + aSlice - 1 < ownerBuffer.length Then aRightPart = Right(ownerBuffer.data_, ownerBuffer.length - (astart + aSlice - 1)) Else aRightPart = ""
        ownerBuffer.data_ = aLeftPart & invalue & aRightPart
    
'    If astart + aSlice = ownerBuffer.Length Then
'        ownerBuffer.Data_ = Left(ownerBuffer.Data_, ownerBuffer.Length - aSlice) + inValue
'    Else
'        ownerBuffer.Data_ = IIf(astart > 1, Left(ownerBuffer.Data_, astart - 1), "") & inValue & _
'            IIf(ownerBuffer.Length - astart - aSlice > 0, Right(ownerBuffer.Data_, ownerBuffer.Length - astart + 1 - aSlice), "")
'    End If
'
''    ownerBuffer.Data_ = IIf(astart > 1, Left(ownerBuffer.Data_, astart - 1), "") & inValue & _
''        IIf(ownerBuffer.Length - astart - aSlice > 0, Right(ownerBuffer.Data_, ownerBuffer.Length - astart + 1 - aSlice), "")
'
End Property

Public Property Get ByName(aname, position) As BufferPart
Dim apart As BufferPart, aSlice As Long
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1
    aSlice = nLength_ \ times
    Set apart = SubStruct.ByName(aname)
    If apart Is Nothing Then
        Set ByName = Nothing
        Exit Property
    Else
    Set ByName = New BufferPart
    With ByName
        Set .owner = owner
        Set .ownerBuffer = ownerBuffer
        Set .SubStruct = apart.SubStruct
        .name = apart.name
        .datatype = apart.datatype
        .times = apart.times
        .nLength_ = apart.nLength_
        .nStartPos_ = nStartPos_ + (position - 1) * aSlice + apart.nStartPos_ - 1
    End With
    End If
End Property

Public Property Get ByIndex(Index, position) As BufferPart
Dim apart As BufferPart, aSlice As Long
    If IsMissing(position) Then position = 1
    If position = 0 Then position = 1
    aSlice = nLength_ \ times
    Set apart = SubStruct.ByIndex(Index)
    
    Set ByIndex = New BufferPart
    With ByIndex
        Set .owner = owner
        Set .ownerBuffer = ownerBuffer
        Set .SubStruct = apart.SubStruct
        .name = apart.name
        .datatype = apart.datatype
        .times = apart.times
        .nLength_ = apart.nLength_
        .nStartPos_ = nStartPos_ + (position - 1) * aSlice + apart.nStartPos_ - 1
    End With
End Property

Public Sub Report(frmOwner)
    Dim astr As String, aSlice As Long, astart As Long, i As Long, k As Long, apart As BufferPart, copypart As BufferPart
    
'    For i = 1 To Times
'        aSlice = nLength_ \ Times
'
'        astr = Mid(ownerBuffer.Data_, nStartPos_ + (i - 1) * aSlice, aSlice)
'        If dataType = ptChar Then
'            Value = EbcdicToAscii_(astr)
'            If Trim(Value) <> "" Then eJournalWriteFld frmOwner, 9999, Name & ":", Value
'        ElseIf dataType = ptPicture Then
'            Value = EbcdicToAscii_(astr)
'            If Trim(Value) <> "" Then eJournalWriteFld frmOwner, 9999, Name & ":", Value
'        ElseIf dataType = ptVarchar Then
'            Value = EbcdicToAscii_(Right(astr, aSlice - 2))
'            If Trim(Value) <> "" Then eJournalWriteFld frmOwner, 9999, Name & ":", Value
'        ElseIf dataType = ptInt Then
'            Value = HpsToInt_(astr)
'            If Value <> 0 Then eJournalWriteFld frmOwner, 9999, Name & ":", CStr(Value)
'        ElseIf dataType = ptTime Then
'            Value = HpsToInt_(astr)
'            If Value <> 0 Then eJournalWriteFld frmOwner, 9999, Name & ":", CStr(Value)
'        ElseIf dataType = ptSmall Then
'            Value = HpsToSmall_(astr)
'            If Value <> 0 Then eJournalWriteFld frmOwner, 9999, Name & ":", CStr(Value)
'        ElseIf dataType = ptDecimal Then
'            Value = HPSToDecimal_(astr)
'            If Value <> 0 Then eJournalWriteFld frmOwner, 9999, Name & ":", CStr(Value)
'        ElseIf dataType = ptDate Then
'            Value = DateAdd("y", IIf(HpsToInt_(astr) >= 36526, HpsToInt_(astr) - 36526, 0), #1/1/100#)
'            eJournalWriteFld frmOwner, 9999, Name & ":", CStr(Value)
'        ElseIf dataType = ptStruct Then
'            If astr <> SubStruct.ClearData_ Then
'                eJournalWriteFld frmOwner, 9999, "---SubStructure: " & SubStruct.Name & ":", CStr(i)
'                For k = 1 To SubStruct.PartNum
'                    Set apart = SubStruct.ByIndex(k)
'                    Set copypart = New BufferPart
'                    With copypart
'                        Set .owner = owner
'                        Set .ownerBuffer = ownerBuffer
'                        Set .SubStruct = apart.SubStruct
'                        .Name = apart.Name
'                        .dataType = apart.dataType
'                        .Times = apart.Times
'                        .nLength_ = apart.nLength_
'                        .nStartPos_ = nStartPos_ + (i - 1) * aSlice + apart.nStartPos_ - 1
'                        .Report frmOwner
'                    End With
'                    Set copypart = Nothing
'                Next k
'            End If
'        End If
'
'
'    Next i
        
End Sub

Public Property Get Struct(Index) As Buffer
    If datatype = ptStruct Then Set Struct = SubStruct(Index) Else Set Struct = Nothing
End Property

Public Sub Clear()
    Set SubStruct = Nothing
    Set localnode = Nothing
    Set ownerBuffer = Nothing
    Set owner = Nothing
End Sub
