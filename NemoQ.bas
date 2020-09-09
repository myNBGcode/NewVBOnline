Attribute VB_Name = "NemoQ"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SysAllocString Lib "oleaut32" (ByVal olestr As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

Private Declare Function StartNQCashierConnection Lib "NQCashierConnection.dll" (ByVal hWnd As Long) As Long
Private Declare Sub StopNQCashierConnection Lib "NQCashierConnection.dll" ()
Private Declare Function LoginCharIntegerUserName Lib "NQCashierConnection.dll" (ByVal letter As Byte, ByVal UserCodeNumber As Long, ByVal name1 As Long, ByVal name2 As Long, ByVal name3 As Long) As Long
Private Declare Function LogoutUser Lib "NQCashierConnection.dll" () As Long
Private Declare Function CallNext_WideStringPointer_Result Lib "NQCashierConnection.dll" () As Long

Function GetString(lPtrToPtrToWidestring As Long) As String
    Dim ptr2 As Long
    ptr2 = Peek(lPtrToPtrToWidestring, vbLong)
           
    GetString = sUniPtrZToVBString(ptr2)
 End Function

Public Function sUniPtrZToVBString(lStrptr As Long) As String
' Convert 'pointer to (wide-)null-terminated Unicode string' to a 'VB string Value'
    sUniPtrZToVBString = vbNullString
    lStrptr = SysAllocString(lStrptr)
    CopyMemory ByVal VarPtr(sUniPtrZToVBString), ByVal VarPtr(lStrptr), 4
End Function


' read a value of any type from memory
Function Peek(ByVal address As Long, ByVal ValueType As VbVarType) As Variant
    Select Case ValueType
        Case vbByte
            Dim valueB As Byte
            CopyMemory valueB, ByVal address, 1
            Peek = valueB
        Case vbInteger
            Dim valueI As Integer
            CopyMemory valueI, ByVal address, 2
            Peek = valueI
        Case vbBoolean
            Dim valueBool As Boolean
            CopyMemory valueBool, ByVal address, 2
            Peek = valueBool
        Case vbLong
            Dim valueL As Long
            CopyMemory valueL, ByVal address, 4
            Peek = valueL
        Case vbSingle
            Dim valueS As Single
            CopyMemory valueS, ByVal address, 4
            Peek = valueS
        Case vbDouble
            Dim valueD As Double
            CopyMemory valueD, ByVal address, 8
            Peek = valueD
        Case vbCurrency
            Dim valueC As Currency
            CopyMemory valueC, ByVal address, 8
            Peek = valueC
        Case vbDate
            Dim valueDate As Date
            CopyMemory valueDate, ByVal address, 8
            Peek = valueDate
        Case vbVariant
            ' in this case we don't need an intermediate variable
            CopyMemory Peek, ByVal address, 16
        Case Else
            Err.Raise 1001, , "Unsupported data type"
    End Select

End Function

Public Sub StartNQCashierAndLogin(ByVal hWnd As Long, UserName As String, FullName As String)
    UserName = UCase(Trim(UserName))
    Dim res1 As Long
    res1 = StartNQCashierConnection(hWnd)
    If res1 = 0 Then
        Sleep 5000
        Dim res As Long
        Dim i As Integer
        For i = 1 To 50
            res = LoginCharIntegerUserName(Asc(Left(UserName, 1)), CLng(Mid(UserName, 2)), StrPtr(FullName), StrPtr(""), StrPtr(""))
            If res >= 0 Then
                HasWinPanelConnection = True
                Exit For
            End If
        Next i
        If res < 0 Then
            LogMsgbox "Αναμένεται σύνδεση με WinPanel...", vbOKOnly, "On Line Εφαρμογή"
            Sleep 30000
            res = LoginCharIntegerUserName(Asc(Left(UserName, 1)), CLng(Mid(UserName, 2)), StrPtr(FullName), StrPtr(""), StrPtr(""))
            If res >= 0 Then
                HasWinPanelConnection = True
            Else
                HasWinPanelConnection = False
                StopNQCashierConnection
                LogMsgbox "Απέτυχε η σύνδεση στο WinPanel." & vbCrLf & "Για να ξαναδοκιμάστετε κλείστε και ανοίξτε το On Line.", vbOKOnly, "On Line Εφαρμογή"
            End If
        End If
    Else
        LogMsgbox "Απέτυχε η έναρξη του WinPanel", vbOKOnly, "On Line Εφαρμογή"
    End If
End Sub

Public Function StopNQCashierAndLogout()
    LogoutUser
    StopNQCashierConnection
End Function

Public Function NQCashierCallNextCustomer() As String
On Error GoTo err_handler
    
    Dim res As String
    Dim ptrResult As Long
    ptrResult = CallNext_WideStringPointer_Result()
    res = GetString(ptrResult)
    
    NQCashierCallNextCustomer = res
    Exit Function
err_handler:
    LogMsgbox "Απέτυχε η ανάκτηση Αρ.Εισιτηρίου από το WinPanel", vbOKOnly, "On Line Εφαρμογή", Err
    Err.Clear
End Function

