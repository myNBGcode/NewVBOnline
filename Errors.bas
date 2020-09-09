Attribute VB_Name = "Errors"
Option Explicit
Public Sub NBG_error(routine_name As String, error As Integer)

    Call EventLog(1, "NBG Error: " & error & " in " & routine_name & "()")
    Call NBG_MsgBox("Runtime Error: " & error & " in " & routine_name & "()", True, " ")

End Sub
Public Sub Runtime_error(routine_name As String, error As Integer, error_msg As String)

    Call EventLog(1, "Runtime Error: " & error & " in " & routine_name & "() " & error_msg)
    Call NBG_MsgBox("Runtime Error: " & error & " in " & routine_name & "() " & error_msg, True, " ")
End Sub

Public Sub NBG_LOG_MsgBox(PStrMessage As String, _
                      Optional PBolBeep As Variant, Optional pstrTitle As Variant)
                      
' Η Procedure NBG_MsgBox πρέπει να χρησιμοποιείται
' για την εμφάνιση διαφόρων μηνυμάτων στην οθόνη
'
' Παράμετροι :
' PStrMessage το μήνυμα που θέλουμε να εμφανισθεί
' PBolBeep    προεραιτικά flag (True, False) αν
'             θέλουμε να κάνει Beep default True
' PstrTitle    ο τίτλος στο παράθυρο
' π.χ.
'
' Call NBG_MsgBox("Λανθασμένος Κωδικός !!", True,"Μήνυμα Λάθους")

  
  If IsMissing(PBolBeep) Then PBolBeep = True
  If PBolBeep Then Beep
  If IsMissing(pstrTitle) Then pstrTitle = "On Line Εφαρμογή"
  'MsgBox PStrMessage, , pstrTitle
  LogMsgbox PStrMessage, , CStr(pstrTitle)
  
  DoEvents
End Sub

Public Sub NBG_MsgBox(PStrMessage As String, _
                      Optional PBolBeep As Variant, Optional pstrTitle As Variant)
                      
' Η Procedure NBG_MsgBox πρέπει να χρησιμοποιείται
' για την εμφάνιση διαφόρων μηνυμάτων στην οθόνη
'
' Παράμετροι :
' PStrMessage το μήνυμα που θέλουμε να εμφανισθεί
' PBolBeep    προεραιτικά flag (True, False) αν
'             θέλουμε να κάνει Beep default True
' PstrTitle    ο τίτλος στο παράθυρο
' π.χ.
'
' Call NBG_MsgBox("Λανθασμένος Κωδικός !!", True,"Μήνυμα Λάθους")

  
  If IsMissing(PBolBeep) Then PBolBeep = True
  If PBolBeep Then Beep
  If IsMissing(pstrTitle) Then pstrTitle = "On Line Εφαρμογή"
  MsgBox PStrMessage, , pstrTitle
  
  DoEvents
End Sub

Public Sub Xml_ParseError(docerror As IXMLDOMParseError)
    LogMsgbox "Λάθος στην επεξεργασία εγγράφου: " & docerror.errorCode & " γραμμή " & docerror.Line & " Θέση " & docerror.linepos & " Θέση αρχείου " & docerror.filepos & _
        vbCrLf & " αιτιολογία " & docerror.reason & " κείμενο " & docerror.srcText, vbCritical, "Λάθος", Err
    'MsgBox "Λάθος στην επεξεργασία εγγράφου: " & docerror.errorCode & " γραμμή " & docerror.Line & " Θέση " & docerror.linepos & " Θέση αρχείου " & docerror.filepos & _
    '   vbCrLf & " αιτιολογία " & docerror.reason & " κείμενο " & docerror.srcText, vbCritical, "Λάθος"
End Sub

Public Sub LogMsgbox(message As String, Optional style As VbMsgBoxStyle, Optional Title As String, Optional error As Variant)
    MsgBox message, style, Title
    'eJournalWriteAll
    If IsMissing(error) Then
        eJournalWrite Title & ":" & message
    ElseIf error Is Nothing Then
        eJournalWrite Title & ":" & message
    Else
        Dim errornumber As Integer
        errornumber = error.Number
        eJournalWrite Title & ":" & message & " " & " Λάθος:" & error.Number & " " & error.description
        If (errornumber = 999) Then
            eJournalWrite "ΠΡΟΣΟΧΗ! ΑΠΕΤΥΧΕ Η ΛΗΨΗ ΑΠΑΝΤΗΣΗΣ.ΚΑΝΤΕ ΕΛΕΓΧΟ ΓΙΑ ΤΗΝ ΤΥΧΗ ΤΗΣ ΣΥΝΑΛΛΑΓΗΣ."
            MsgBox "ΠΡΟΣΟΧΗ! ΑΠΕΤΥΧΕ Η ΛΗΨΗ ΑΠΑΝΤΗΣΗΣ.ΚΑΝΤΕ ΕΛΕΓΧΟ ΓΙΑ ΤΗΝ ΤΥΧΗ ΤΗΣ ΣΥΝΑΛΛΑΓΗΣ."
        End If
    End If
    
    
End Sub

'Public Sub LogMsgbox(error, message As String, Optional style As VbMsgBoxStyle, Optional Title As String)
'    MsgBox message, style, Title
'    'eJournalWriteAll
'    If error Is Nothing Then
'        eJournalWrite Title & ":" & message
'    Else
'        eJournalWrite Title & ":" & message & " " & " Λάθος:" & error.Number & " " & error.description
'    End If
'
'
'End Sub

