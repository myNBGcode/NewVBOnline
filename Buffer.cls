VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Definition As String, BuffType As String, PartNum As Integer
Private fName As String, fDescription As String, PartList() As BufferPart
Public xmlDocV2 As MSXML2.DOMDocument30, xmlStrV2 As String, localnode As MSXML2.IXMLDOMElement
Public StructID As String

Public data_ As String, ClearData_ As String, length As Long, LastLevel As Boolean

Public Property Let name(value As String)
    fName = value
    'If Not (xmlDocV2.documentElement Is Nothing) Then xmlDocV2.removeChild xmlDocV2.documentElement
    'Dim VName As String
    'If IsNumeric(Left(Value, 1)) Then VName = "_" & UCase(VName) Else VName = UCase(Value)
    'VName = Replace(VName, ":", "_"): VName = Replace(VName, " ", "_")
    'Set xmlDocV2.documentElement = xmlDocV2.createElement("IRISDOCUMENT")
End Property

Public Property Get name() As String
    name = fName
End Property

Public Sub Clear()
    Dim i As Integer
    For i = 1 To PartNum
        If PartList(i).datatype <> ptStruct Then Set PartList(i) = Nothing Else PartList(i).Clear
    Next i
    ReDim PartList(0): PartNum = 0
    Set xmlDocV2 = Nothing
    Set localnode = Nothing
End Sub

Public Sub ClearData()
    data_ = ClearData_
End Sub

Private Function NewPartXMLV2(inPart As BufferPart)
    Dim nodename As String, newElm As MSXML2.IXMLDOMElement, aelm As MSXML2.IXMLDOMElement, k As Long
    Dim newAttr As MSXML2.IXMLDOMAttribute, STAttr As MSXML2.IXMLDOMAttribute, LAttr As MSXML2.IXMLDOMAttribute
    Dim aDataType As String, alength As Long, astart As Long, xmlPart As String, xmlPart2 As String
    
    If IsNumeric(Left(inPart.name, 1)) Then nodename = "_" & UCase(inPart.name) Else nodename = UCase(inPart.name)

'-------------------------
    If Not (inPart.SubStruct Is Nothing) Then
       alength = inPart.SubStruct.length
    Else
       alength = inPart.nLength_
    End If
    astart = inPart.nStartPos_
        
    If inPart.datatype = ptChar Then
        aDataType = "char"
    ElseIf inPart.datatype = ptPicture Then
        aDataType = "picture"
    ElseIf inPart.datatype = ptVarchar Then
        aDataType = "varchar"
    ElseIf inPart.datatype = ptInt Then
        aDataType = "int"
    ElseIf inPart.datatype = ptTime Then
        aDataType = "time"
    ElseIf inPart.datatype = ptSmall Then
        aDataType = "small"
    ElseIf inPart.datatype = ptDecimal Then
        aDataType = "decimal"
    ElseIf inPart.datatype = ptTimeStamp Then
        aDataType = "timestamp"
    ElseIf inPart.datatype = ptDate Then
        aDataType = "date"
    ElseIf inPart.datatype = ptStruct Then
        aDataType = "struct"
    End If

    If Not (inPart.SubStruct Is Nothing) Then
       xmlPart = "<" & nodename & " DATATYPE=""" & aDataType & """ OCCURS=""" & inPart.times & """ START=""%ST%"" LENGTH=""" & alength & """"
       alength = inPart.SubStruct.length
       xmlPart = xmlPart & ">" & inPart.SubStruct.xmlStrV2 & "</" & nodename & ">"
    Else
       xmlPart = "<" & nodename & " DATATYPE=""" & aDataType & """ OCCURS=""" & inPart.times & """ START=""%ST%"" LENGTH=""" & alength & """"
       xmlPart = xmlPart & "/>"
    End If
    xmlPart2 = ""
    For k = 1 To inPart.times
        xmlPart2 = xmlPart2 & Replace(xmlPart, "%ST%", CStr(astart - 1))
        astart = astart + alength
    Next k
    
    xmlStrV2 = xmlStrV2 & xmlPart2
End Function

Public Function NewPart(inBuffers As Buffers, inName As String, inPartType As Integer, inTimes As Long, inLength As Long, inDataPart As String, _
    aStructName As String, astructtype As String, inSubStruct As Buffer) As BufferPart
    Dim k As Long
    PartNum = PartNum + 1
    ReDim Preserve PartList(1 To PartNum)
    
    Set PartList(PartNum) = New BufferPart
    With PartList(PartNum)
        Set .owner = inBuffers
        Set .ownerBuffer = Me
        .name = inName
        .datatype = inPartType
        .times = inTimes
        Set .SubStruct = inSubStruct
        'If inPartType = ptStruct Then .Setstructtype inBuffers, aStructName, aStructType
        .nStartPos_ = length + 1
        .nLength_ = inLength * inTimes
        For k = 1 To inTimes
            data_ = data_ & inDataPart
        Next k
        length = length + .nLength_
    End With
    Set NewPart = PartList(PartNum)
'    NewPartXML NewPart
    NewPartXMLV2 NewPart
    'Set NewPart.PartElm = newElm
End Function

Public Property Get ByIndex(Index, Optional idx) As BufferPart
    If IsMissing(idx) Then idx = 1
    Set ByIndex = PartList(Index)
End Property

Public Property Get xmlNode(aNodeName, Optional idx)
    Dim aNodeList As MSXML2.IXMLDOMNodeList
    Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
    Dim aPos As Long, alength As Long
    If IsMissing(idx) Then idx = 1
    If idx = 1 Then
        Set anode = xmlDocV2.selectSingleNode(".//" & UCase(aNodeName))
    Else
        Set aNodeList = xmlDocV2.SelectNodes(".//" & UCase(aNodeName))
        If aNodeList.length >= idx Then Set anode = aNodeList.item(idx - 1) Else Set anode = Nothing
    End If
    Set xmlNode = anode
End Property

Public Property Get v2Copy(aNodeName, Optional idx) As Buffer
Dim aNodeList As MSXML2.IXMLDOMNodeList
Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim aPos As Long, alength As Long
If IsMissing(idx) Then idx = 1
    Set anode = xmlNode(aNodeName, idx)
    
    If Not (anode Is Nothing) Then
        Dim rTitle As String
        rTitle = localnode.nodename
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        
        Set bNode = anode.parentNode
        While bNode.nodename <> rTitle
            aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
            Set bNode = bNode.parentNode
        Wend
        Set v2Copy = New Buffer
        With v2Copy
            .data_ = Mid(data_, aPos + 1, alength)
            .ClearData_ = Mid(ClearData_, aPos + 1, alength)
            .length = alength
            .LastLevel = True
            Set .xmlDocV2 = New MSXML2.DOMDocument30
            .xmlDocV2.LoadXML anode.XML
            .xmlStrV2 = anode.XML
            .BuffType = aNodeName
            .name = aNodeName
        End With
    Else
        v2Copy = Nothing
    End If
End Property


Public Property Get v2Data(aNodeName, Optional idx)
Dim aNodeList As MSXML2.IXMLDOMNodeList
Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim aPos As Long, alength As Long
If IsMissing(idx) Then idx = 1
    Set anode = xmlNode(aNodeName, idx)
    If Not (anode Is Nothing) Then
        Dim rTitle As String
        rTitle = xmlDocV2.documentElement.nodename
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        
        Set bNode = anode.parentNode
        While bNode.nodename <> rTitle
            aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
            Set bNode = bNode.parentNode
        Wend
        
        v2Data = Mid(data_, aPos + 1, alength)
    Else
        v2Data = ""
    End If
    
End Property

Public Property Let v2Data(aNodeName, Optional idx, aValue)
Dim aNodeList As MSXML2.IXMLDOMNodeList
Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim aPos As Long, alength As Long
If IsMissing(idx) Then idx = 1
    Set anode = xmlNode(aNodeName, idx)
    If Not (anode Is Nothing) Then
        Dim rTitle As String
        rTitle = xmlDocV2.documentElement.nodename
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        
        Set bNode = anode.parentNode
        While bNode.nodename <> rTitle
            aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
            Set bNode = bNode.parentNode
        Wend
        If aPos > 0 Then
            aValue = Right(Mid(data_, aPos, alength) & aValue, alength)
        Else
            aValue = Right(Left(data_, alength) & aValue, alength)
        End If
        
        Dim aLeftPart As String, aRightPart As String
        If aPos > 0 Then aLeftPart = Left(data_, aPos) Else aLeftPart = ""
        If aPos + alength < length Then aRightPart = Right(data_, length - (aPos + alength)) Else aRightPart = ""
        data_ = aLeftPart & aValue & aRightPart
        
    Else
    
    End If
    
End Property

Public Property Get v2Value(aNodeName, Optional idx)
Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim aPos As Long, alength As Long, aDataType As String, aDataPart As String
If IsMissing(idx) Then idx = 1
    Set anode = xmlNode(aNodeName, idx)
    If Not (anode Is Nothing) Then
        Dim rTitle As String
        rTitle = xmlDocV2.documentElement.nodename
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        aDataType = anode.Attributes.getNamedItem("DATATYPE").nodeValue
        Set bNode = anode.parentNode
        While bNode.nodename <> rTitle
            aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
            Set bNode = bNode.parentNode
        Wend
        aDataPart = Mid(data_, aPos + 1, alength)
        
        If aDataType = "char" Then
            v2Value = EbcdicToAscii_(aDataPart)
        ElseIf aDataType = "picture" Then
            v2Value = EbcdicToAscii_(aDataPart)
        ElseIf aDataType = "varchar" Then
            v2Value = EbcdicToAscii_(Right(aDataPart, alength - 2))
        ElseIf aDataType = "int" Then
            v2Value = HpsToInt_(aDataPart)
        ElseIf aDataType = "time" Then
            If OpenCobolServer Then
                Dim ahh As Long
                Dim amm As Long
                Dim ass As Long
                Dim amss As Long
                
                aDataPart = EbcdicToAscii_(aDataPart)
                ahh = Left(aDataPart, 2)
                amm = Mid(aDataPart, 4, 2)
                ass = Mid(aDataPart, 7, 2)
                amss = Right(aDataPart, 2)
                v2Value = ((ahh * 60 + amm) * 60 + ass) * 1000 + amss
            Else
                v2Value = HpsToInt_(aDataPart)
            End If
        
            'v2Value = HpsToInt_(aDataPart)
        ElseIf aDataType = "small" Then
            v2Value = HpsToSmall_(aDataPart)
        ElseIf aDataType = "decimal" Then
            v2Value = HPSToDecimal_(aDataPart)
        ElseIf aDataType = "timestamp" Then
            v2Value = HpsToInt_(Left(aDataPart, 4)) & "," & HpsToInt_(Mid(aDataPart, 5, 4)) & "," & HpsToInt_(Right(aDataPart, 4))
        ElseIf aDataType = "date" Then
            If OpenCobolServer Then
                v2Value = EbcdicToAscii_(aDataPart)
                v2Value = DateSerial(Left(v2Value, 4), Mid(v2Value, 6, 2), Right(v2Value, 2))
            Else
                v2Value = DateAdd("y", IIf(HpsToInt_(aDataPart) >= 36526, HpsToInt_(aDataPart) - 36526, 0), #1/1/100#)
            End If
            
            'v2Value = DateAdd("y", IIf(HpsToInt_(aDataPart) >= 36526, HpsToInt_(aDataPart) - 36526, 0), #1/1/100#)
        End If
    Else
        v2Value = ""
    End If
End Property

Public Property Get AsString(aNodeName, Optional idx)
    AsString = CStr(v2Value(aNodeName, idx))
End Property

Public Property Let v2Value(aNodeName, Optional idx, aValue)
Dim anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim aPos As Long, alength As Long, aDataType As String, aDataPart As String
If IsMissing(idx) Then idx = 1
    Set anode = xmlNode(aNodeName, idx)
    If Not (anode Is Nothing) Then
        Dim rTitle As String
        rTitle = xmlDocV2.documentElement.nodename
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        aDataType = anode.Attributes.getNamedItem("DATATYPE").nodeValue
        Set bNode = anode.parentNode
        While bNode.nodename <> rTitle
            aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
            Set bNode = bNode.parentNode
        Wend
        
        If aDataType = "char" Then
            aDataPart = AsciiToEbcdic_(Left(CStr(aValue) & String(alength, " "), alength))
        ElseIf aDataType = "picture" Then
            aDataPart = AsciiToEbcdic_(Left(CStr(aValue) & String(alength, " "), alength))
        ElseIf aDataType = "varchar" Then
            aDataPart = VarCharToHps_(CStr(aValue), alength - 2)
        ElseIf aDataType = "int" Then
            aDataPart = IntToHps_(CLng(aValue))
        ElseIf aDataType = "time" Then
            If OpenCobolServer Then
                Dim ahh As Long
                Dim amm As Long
                Dim ass As Long
                Dim amss As Long
                
                amss = aValue Mod 1000
                aValue = aValue \ 1000
                ass = aValue Mod 60
                aValue = aValue \ 60
                amm = aValue Mod 60
                aValue = aValue \ 60
                ahh = aValue Mod 60
                aDataPart = Right("00" & ahh, 2) & ":" & Right("00" & amm, 2) & ":" & Right("00" & ass, 2) & ":" & Right("000" & amss, 3)
                aDataPart = AsciiToEbcdic_(aDataPart)
            Else
                aDataPart = IntToHps_(CLng(aValue))
            End If
            'aDataPart = IntToHps_(CLng(aValue))
        ElseIf aDataType = "small" Then
            aDataPart = SmallToHps_(CLng(aValue))
        ElseIf aDataType = "decimal" Then
            aDataPart = DecimalToHPS_(CDbl(IIf(Trim(aValue) <> "", aValue, 0)), alength * 2 - 1)
        ElseIf aDataType = "timestamp" Then
            aDataPart = String(12, Chr(0))
        ElseIf aDataType = "date" Then
            If OpenCobolServer Then
                If Not IsDate(aValue) Then
                    aDataPart = Right(aValue, 4) + "-" + Mid(aValue, 3, 2) + "-" + Mid(aValue, 1, 2)
                Else
                    aDataPart = Right("0000" & Year(aValue), 4) & "-" & Right("00" & Month(aValue), 2) & "-" & Right("00" & Day(aValue), 2)
                End If
                aDataPart = AsciiToEbcdic_(aDataPart)
            Else
'                If aValue = DateSerial(1900, 1, 1) Then
'                    aDataPart = IntToHps_(0)
'                Else
'                    aDataPart = IntToHps_(36526 + DateDiff("y", #1/1/100#, CDate(aValue)))
'                End If
                
                If Not IsEmpty(aValue) And Not IsDate(aValue) Then
                    aValue = Trim(aValue)
                    aValue = Right(aValue, 4) + "-" + Mid(aValue, 3, 2) + "-" + Mid(aValue, 1, 2)
                End If
                If CDate(aValue) = DateSerial(1900, 1, 1) Then
                    aDataPart = IntToHps_(0)
                Else
                    aDataPart = IntToHps_(36526 + DateDiff("y", #1/1/100#, CDate(aValue)))
                End If
            End If
        End If
        
        Dim aLeftPart As String, aRightPart As String
        If aPos > 0 Then aLeftPart = Left(data_, aPos) Else aLeftPart = ""
        If aPos + alength < length Then aRightPart = Right(data_, length - (aPos + alength)) Else aRightPart = ""
        data_ = aLeftPart & aDataPart & aRightPart
        
        
'        If Length = apos + alength Then
'            Data_ = Left(Data_, Length - alength) & aDataPart
'        Else
'            Data_ = IIf(apos > 1, Left(Data_, apos - 1), "") & aDataPart & _
'                IIf(Length - apos - alength > 0, Right(Data_, Length - apos + 1 - alength), "")
'        End If
    End If
    
End Property

Public Function GetXMLView()
Dim i As Long, anode As MSXML2.IXMLDOMElement, bNode As MSXML2.IXMLDOMElement
Dim rTitle As String, aPos As Long, alength As Long, aDataType As String, aDataPart As String

For Each anode In xmlDocV2.documentElement.SelectNodes(".//*")
    If Not (anode Is Nothing) Then
        aPos = anode.Attributes.getNamedItem("START").nodeValue
        alength = anode.Attributes.getNamedItem("LENGTH").nodeValue
        aDataType = anode.Attributes.getNamedItem("DATATYPE").nodeValue
        If aDataType <> "struct" Then
            Set bNode = anode.parentNode
            rTitle = xmlDocV2.documentElement.nodename
            While bNode.nodename <> rTitle
                aPos = aPos + bNode.Attributes.getNamedItem("START").nodeValue
                Set bNode = bNode.parentNode
            Wend
            aDataPart = Mid(data_, aPos + 1, alength)
            If aDataType = "char" Then
                anode.Text = EbcdicToAscii_(aDataPart)
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "picture" Then
                anode.Text = EbcdicToAscii_(aDataPart)
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "varchar" Then
                anode.Text = EbcdicToAscii_(Right(aDataPart, alength - 2))
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "int" Then
                anode.Text = HpsToInt_(aDataPart)
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "time" Then
                If OpenCobolServer Then
                    anode.Text = EbcdicToAscii_(aDataPart)
                Else
                    anode.Text = HpsToInt_(aDataPart)
                End If
                anode.setAttribute "FORMATTEDTEXT", CalculateIRISTime(anode.Text)
            ElseIf aDataType = "small" Then
                anode.Text = HpsToSmall_(aDataPart)
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "decimal" Then
                anode.Text = HPSToDecimal_(aDataPart)
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "timestamp" Then
                anode.Text = HpsToInt_(Left(aDataPart, 4)) & "," & HpsToInt_(Mid(aDataPart, 5, 4)) & "," & HpsToInt_(Right(aDataPart, 4))
                anode.setAttribute "FORMATTEDTEXT", anode.Text
            ElseIf aDataType = "date" Then
                If OpenCobolServer Then
                    anode.Text = EbcdicToAscii_(aDataPart)
                Else
                    If HpsToInt_(aDataPart) = 367 Then
                        anode.Text = "01010001"
                        anode.setAttribute "FORMATTEDTEXT", "01/01/0001"
                    Else
                        Dim adate As Date
                        adate = DateAdd("y", IIf(HpsToInt_(aDataPart) >= 36526, HpsToInt_(aDataPart) - 36526, 0), #1/1/100#)
                        anode.Text = Right("00" & Day(adate), 2) & Right("00" & Month(adate), 2) & Right("0000" & Year(adate), 4)
                        anode.setAttribute "FORMATTEDTEXT", Right("00" & Day(adate), 2) & "/" & Right("00" & Month(adate), 2) & "/" & Right("0000" & Year(adate), 4)
                    End If
                End If
            End If
        End If
    End If
Next anode
        
Set GetXMLView = xmlDocV2

End Function

Public Property Get ByName(aPartName, Optional idx) As BufferPart
Dim foundflag As Boolean, i As Integer, PartPos As Integer
If IsMissing(idx) Then idx = 1
    Set ByName = Nothing
    foundflag = False
    If PartNum > 0 Then
        For i = 1 To PartNum
            If UCase(PartList(i).name) = UCase(aPartName) Then
                PartPos = i: foundflag = True: Exit For
            End If
        Next i
    End If
    If Not foundflag Then
        LogMsgbox "ΛΑΘΟΣ το τμήμα: " & aPartName & " δεν υπάρχει", vbCritical, "ΛΑΘΟΣ"
        Exit Property
    Else
        Set ByName = PartList(PartPos)
    End If
End Property

Public Sub Report(frmOwner)
Dim i As Long
    If data_ = ClearData_ Then Exit Sub
End Sub


Public Property Get Data(Optional idx) As String
    Data = data_
End Property

Public Property Let Data(Optional idx, value As String)
    data_ = value
End Property

Private Sub Class_Initialize()
    'Set xmlDoc = New MSXML2.DOMDocument
    Set xmlDocV2 = New MSXML2.DOMDocument
End Sub

Private Sub Class_Terminate()
    'Set xmlDocV2 = Nothing
    Clear
End Sub
