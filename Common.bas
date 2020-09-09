Attribute VB_Name = "Common"
Option Explicit

'Public Function Update_Jrn_After_Ok(pData, pIndex As Integer, pBoolTime As Boolean)
'    ' ενημερωνει σε ομαλή περίπτωση 1 ή 2 γραμμές τον πίνακα που θα εκτυπωθεί
'    Dim MStrMyString(2) As String
'    Dim minti As Integer
'    Dim MBolResult As Boolean
'    Dim mi As Integer
'
'    minti = 1
'    If pBoolTime Then
'        MStrMyString(minti) = Space(10) + "ΩΡΑ ΣΥΝ/ΓΗΣ ..: " & _
'                            Mid(cb.received_data, 14, 2) & ":" & _
'                            Mid(cb.received_data, 16, 2) & ":" & _
'                            Mid(cb.received_data, 18, 2)
'        minti = minti + 1
'    End If
'    MStrMyString(minti) = "TELLER Name"
'
'    For mi = 1 To minti
'        pData(pIndex + mi) = MStrMyString(mi)
'    Next
'    Update_Jrn_After_Ok = pIndex + minti
'End Function

'Public Function MaskPoso(pintPoso As Currency, pstrEidos As String) As String
'    'pstrEidos παιρνει τιμες :      "+" = Ανάληψη
'    '                               "-" = Κατάθεση
'    '                               " " = Υπόλοιπο
'
'    Dim strMask As String
'    Dim lonPoso As Currency
'    Dim intLength As Integer
'    Dim strposo As String
'    Dim strDekadika As String
'
'    If pintPoso = 0 Then
'        lonPoso = 0
'    Else
'        lonPoso = Int(pintPoso / 100)
'    End If
'    strposo = Trim(Str(lonPoso))
'    strDekadika = StrPad_(Trim(Right$(Str(pintPoso), 2)), 2, "0", "L")
'    intLength = Len(strposo)
'
'    If pstrEidos = "+" Or pstrEidos = "-" Then
'        Select Case intLength
'            Case 1 To 8:    strMask = StrPad_(StrPad_(strposo, 8, "*", "L"), 10, " ", "L")
'            Case 9, 10:     strMask = StrPad_(strposo, 10, " ", "L")
'        End Select
'        strMask = strMask & "," & strDekadika
'    Else
'        Select Case intLength
'            Case 1 To 9:    strMask = StrPad_(strposo, 9, "*", "L")
'            Case Else:      strMask = strposo
'        End Select
'        strMask = strMask & "," & strDekadika
'    End If
'
'    Select Case pstrEidos
'        Case "+":   strMask = strMask & Space(5)
'        Case "-":   strMask = Space(5) & strMask
'        Case " ":   strMask = strMask & Chr$(64)
'    End Select
'
'    MaskPoso = strMask
'End Function



