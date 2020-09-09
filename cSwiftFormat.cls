VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSwiftFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private InputValue As String
Private InputFormat As String
Public Property Let invalue(ByVal avalue As String)
      InputValue = avalue
End Property
Public Property Let InFormat(ByVal avalue As String)
      InputFormat = avalue
End Property
Public Property Get FormatedValue()
    If Trim(InputFormat) = "" Then
       FormatedValue = InputValue
       Exit Property
    End If
    If Trim(InputValue) = "" Then
       FormatedValue = InputValue
       Exit Property
    End If
    Dim apos1, apos2, LEvalue, TIvalue, STvalue  As Integer
    Dim vstr, fstr, fpart, rstr, fvalue As String
    Dim DEvalue, RPvalue, LPvalue, DFvalue As String
    Dim Totallen As Integer
    Dim pstr As String
    Dim returnstr As String
    
    returnstr = ""
    vstr = InputValue: fstr = InputFormat
    Do While (True)
        apos1 = InStr(fstr, "{")
        If apos1 > 0 Then
            apos2 = InStr(fstr, "}")
            fpart = Mid(fstr, apos1 + 1, apos2 - apos1 - 1)
            LEvalue = ExtractTokenValue("LE:", fpart)
            If LEvalue = -1 Then LEvalue = 0
            TIvalue = ExtractTokenValue("TI:", fpart)
            If TIvalue = -1 Then TIvalue = 0
            STvalue = ExtractTokenValue("ST:", fpart)
            If STvalue = -1 Then STvalue = 1
            DEvalue = ExtractTokenValue("DE:", fpart)
            If DEvalue = "-1" Then DEvalue = vbCrLf
            RPvalue = ExtractTokenValue("RP:", fpart)
            If RPvalue = "-1" Then RPvalue = ""
            LPvalue = ExtractTokenValue("LP:", fpart)
            If LPvalue = "-1" Then LPvalue = ""
            DFvalue = ExtractTokenValue("DF:", fpart)
            If DFvalue = "-1" Then DFvalue = ""
            Totallen = LEvalue * TIvalue
            If Totallen = 0 Then Totallen = Len(vstr)
            rstr = Mid(vstr, STvalue, Totallen)
            If Trim(rstr) <> "" Then
               pstr = FormatValue(rstr, LEvalue, TIvalue, DEvalue)
               If LPvalue <> "" Then pstr = Right(String(Totallen, LPvalue) & pstr, Totallen)
               If RPvalue <> "" Then pstr = Left(pstr & String(Totallen, RPvalue), Totallen)
               If DFvalue = "" Then
                  returnstr = returnstr & pstr
               Else
                  returnstr = returnstr & FormatDelimiterValue(pstr, DFvalue)
               End If
            Else
               
            End If
            fstr = Mid(fstr, apos2 + 1)
        Else
            Exit Do
        End If
    Loop
    If Right(returnstr, Len(DEvalue)) = DEvalue And Len(returnstr) > Len(DEvalue) Then
       returnstr = Left(returnstr, Len(returnstr) - Len(DEvalue))
    End If
    FormatedValue = returnstr
    'FormatedValue = "/1" & vbCrLf & "ADSFSDFSDFSDFSDFSDFSDFS"
End Property

Private Function ExtractTokenValue(ByVal atoken As String, ByVal apart As String) As String
    Dim aposstart, aposend As Integer
    Dim returnstr As String
    aposstart = InStr(apart, atoken)
    If aposstart > 0 Then
       aposend = InStr(Mid(apart, aposstart), "$")
       If aposend > 0 Then
          returnstr = Mid(apart, aposstart + 3, aposend - 4)
          If returnstr = "" And atoken <> "DE:" Then returnstr = "-1"
       Else
          returnstr = "-1"
       End If
    Else
       returnstr = "-1"
    End If
    ExtractTokenValue = returnstr
End Function

Public Function FormatValue(ByVal avalue As String, ByVal alinelength As Integer, Optional ByVal alinenum As Integer, Optional alineDEL, Optional ByVal alineindex As Integer) As String
    If IsMissing(alineindex) Or alineindex = 0 Then alineindex = 1
    If alineindex <= Int((Len(avalue) / alinelength)) Then
       FormatValue = Mid(avalue, (alineindex - 1) * alinelength + 1, alinelength) & alineDEL & _
                          FormatValue(avalue, alinelength, alinenum, alineDEL, alineindex + 1)
    Else
       FormatValue = Mid(avalue, (alineindex - 1) * alinelength + 1, alinelength)
    End If
End Function

Private Function FormatDelimiterValue(ByVal avalue As String, ByVal adel As String)
   If avalue = "" Or adel = "" Then
      FormatDelimiterValue = avalue
      Exit Function
   End If
   Dim returnstr, LineStr, curstr As String
   Dim apos, i, alinelength  As Integer
   returnstr = "": LineStr = ""
   apos = InStr(avalue, vbCrLf)
   If apos > 0 Then
      alinelength = apos - 1
   Else
      alinelength = Len(avalue)
   End If
   For i = 1 To Len(avalue)
      curstr = Mid(avalue, i, 1)
      If curstr <> vbCr And curstr <> vbLf And curstr <> adel Then
         returnstr = returnstr & curstr
         LineStr = LineStr & curstr
      End If
      If curstr = adel Then
         If LineStr <> "" Then
           returnstr = returnstr & vbCrLf & curstr
         Else
           returnstr = returnstr & curstr
         End If
         LineStr = curstr
      End If
      If Len(LineStr) = alinelength And curstr <> vbCr And curstr <> vbLf And curstr <> adel Then
         If LineStr <> adel Then
            returnstr = returnstr & vbCrLf
            LineStr = ""
         End If
      End If
   Next i
   FormatDelimiterValue = returnstr
End Function


