Attribute VB_Name = "Files"
Option Explicit
' Common Object Variables for file handling --------------------------

'Public rsLogCounters As Recordset
'Public rsBatch As Recordset
'Public rsTransData As Recordset
'Public rsFile16 As Recordset

'Public Const dbFile = "c:\my_data\telLER.MDB"

'Public Const ERR_NOCURRENTREC = 3021
'
'Updates one Constant value given the column name and the new value
'
' Read from LogCounter
'Public Function fnReadLogCounters(strCounter As String) As Currency
'    Dim sCriterion As String
'    rsLogCounters.MoveFirst
'    sCriterion = "CounterName = '" & strCounter & "'"
'    rsLogCounters.FindFirst sCriterion
'    fnReadLogCounters = rsLogCounters("Amount")
'End Function
'Updates Counter value given the column name and the new value and the name of counter
'
'Public Sub sbUpdateLogCounters(ByVal vValue As Currency, _
'                                ByVal strCounter As String)
'    Dim bTrans As Boolean
'    Dim sCriterion As String
'
'    On Error GoTo sbUpdCon_Err
'
'    bTrans = False
'    wks.BeginTrans
'    bTrans = True
'
'        rsLogCounters.MoveFirst
'        sCriterion = "CounterName = '" & strCounter & "'"
'        rsLogCounters.FindFirst sCriterion
'        rsLogCounters.Edit
'        rsLogCounters("Amount") = vValue
'        rsLogCounters.Update
'
'    wks.CommitTrans
'    bTrans = False
'
'sbUpdCon_Exit:
'    Exit Sub
'
'sbUpdCon_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbUpdCon_Exit
'
'End Sub
'μηδενισμός αθροιστων
'
'Public Sub sbZeroLogCounters()
'    Dim bTrans As Boolean
'
'    On Error GoTo sbUpdCon_Err
'
'    bTrans = False
'    wks.BeginTrans
'    bTrans = True
'
'    rsLogCounters.MoveFirst
'    Do Until rsLogCounters.EOF
'        rsLogCounters.Edit
'        rsLogCounters("Amount") = 0
'        rsLogCounters.Update
'        rsLogCounters.MoveNext
'    Loop
'        rsBatch.MoveFirst
'        rsBatch.Edit
'        rsBatch("Amount") = 1
'        rsBatch.Update
'    wks.CommitTrans
'    bTrans = False
'
'sbUpdCon_Exit:
'    Exit Sub
'
'sbUpdCon_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbUpdCon_Exit
'
'End Sub


' Call this proc to close all open recordsets when quiting the application
'
'Sub sbCloseRecordSets()
'    ado_Constants.Close
'    rsBatch.Close
'    rsLogCounters.Close
'    rsTransData.Close
'End Sub

'Public Sub sbUpdateBatch(ByVal vAmount As Currency, _
'                        ByVal vCounterAmount As String, _
'                        blnStopOn20 As Boolean)
'
'    Dim bTrans As Boolean
'    Dim dblOldBatch, dblOldAa As Double
'    Dim sCriterion As String
'    Dim mcurAmount As Currency
'
'    On Error GoTo sbUpdCon_Err
'
'    bTrans = False
'    wks.BeginTrans
'    bTrans = True
'
'        rsBatch.MoveFirst
'        dblOldBatch = Int(rsBatch("Amount"))
'
'        rsBatch.MoveNext
'        dblOldAa = Int(rsBatch("Amount"))
'
'        sCriterion = "Recno = " & Str(2 + dblOldAa)
'        rsBatch.FindFirst sCriterion
'        If rsBatch.NoMatch Then GoTo sbUpdCon_Err
'        rsBatch.Edit
'        rsBatch("Amount") = vAmount
'        rsBatch("CounterAmount") = vCounterAmount
'        rsBatch.Update
'
'        sCriterion = "Recno = 2"
'        rsBatch.FindFirst sCriterion
'        rsBatch.Edit
'        rsBatch("Amount") = Int(rsBatch("Amount")) + 1
'        rsBatch.Update
'
'        'ενημερωση του αθροιστή στο αρχείο αθροιστών με το ποσό
'        ' για την συναλλαγή 2010 που θέλει 4Ζ1,4Ζ2,4Ζ3,4Ζ4,4Ζ5 στις αναλυτικές
'        ' ενώ στο αρχείο πηγαίνει στον αθροιστή 4Ζ δίνουμε στην function τα 2 πρώτα
'        mcurAmount = fnReadLogCounters(Mid(vCounterAmount, 1, 2)) + vAmount
'        Call sbUpdateLogCounters(mcurAmount, Mid(vCounterAmount, 1, 2))
'
'
'        If blnStopOn20 And (Int(rsBatch("Amount")) > 20) Then
'
'           Call Print_Batch
'
'           sCriterion = "Recno = 1"
'           rsBatch.FindFirst sCriterion
'           rsBatch.Edit
'           rsBatch("Amount") = Int(rsBatch("Amount")) + 1
'           rsBatch.Update
'
'           sCriterion = "Recno = 2"
'           rsBatch.FindFirst sCriterion
'           rsBatch.Edit
'           rsBatch("Amount") = 1
'           rsBatch.Update
'        End If
'
'    wks.CommitTrans
'    bTrans = False
'
'sbUpdCon_Exit:
'    Exit Sub
'
'sbUpdCon_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbUpdCon_Exit
'
'
'End Sub

'Public Sub sbInsertFile16()
'
'    Dim bTrans As Boolean
'    Dim minti As Integer
'
'    On Error GoTo sbInsFile16_Err
'
'    Screen.ActiveForm.stbStatus(0).Panels(1).Text = " ΠΑΡΑΚΑΛΩ ΠΕΡΙΜΕΝΕΤΕ ...... "
'
'    bTrans = False
'    minti = 1
'    wks.BeginTrans
'        bTrans = True
'        If Not rsFile16.BOF Then
'           rsFile16.MoveLast
'           minti = rsFile16!RecNo + 1
'        End If
'        rsFile16.AddNew
'        If Mid(cb.received_data, 1, 1) = 0 Then
'            rsFile16!NextPage = "1"
'        Else
'            rsFile16!NextPage = "0"
'        End If
'        rsFile16!NextLine = Mid(cb.received_data, 1, 1)
'        rsFile16!TextLine = Mid(cb.received_data, 2, 80)
'        rsFile16!RecNo = minti
'        rsFile16.Update
'    wks.CommitTrans
'    bTrans = False
'
'sbInsFile16_Exit:
'    Exit Sub
'
'sbInsFile16_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbInsFile16_Exit
'
'
'End Sub

'Public Sub sbDeleteFile16()
'
'    Dim bTrans As Boolean
'
'    On Error GoTo sbDelFile16_Err
'
'    bTrans = False
'    wks.BeginTrans
'        bTrans = True
'        If Not rsFile16.BOF Then
'            rsFile16.MoveFirst
'            Do While Not rsFile16.EOF
'                rsFile16.Delete
'                rsFile16.MoveNext
'            Loop
'            rsFile16.MoveFirst
'        End If
'    wks.CommitTrans
'    bTrans = False
'
'sbDelFile16_Exit:
'    Exit Sub
'
'sbDelFile16_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbDelFile16_Exit
'
'
'End Sub

'Public Sub fnReadFile16(strfield, pmIntPage As Integer)
'
'    Dim bTrans As Boolean
'    Dim inti As Integer
'    Dim sCriterion As String
'    Static mRecordPointer As Integer
'
'    On Error GoTo sbRdFile16_Err
'
'    If pmIntPage = 1 Then
'        mRecordPointer = 1
'    End If
'
'    bTrans = False
'    inti = 1
'    wks.BeginTrans
'        bTrans = True
'        sCriterion = "RecNo = " & Str(mRecordPointer)
'        rsFile16.FindFirst sCriterion
'        If rsFile16.NoMatch Then GoTo sbRdFile16_Err
'        Do While Not (rsFile16.EOF)
'            strfield(inti) = rsFile16!TextLine
''            If rsTrans2310!NextPage = 1 Then
''               Exit Do
''            End If
'            inti = inti + rsFile16!NextLine
'            mRecordPointer = mRecordPointer + 1
'            rsFile16.Move 1
'        Loop
'    wks.CommitTrans
'    bTrans = False
'    If rsFile16.EOF Then
'        pmIntPage = 0
'    Else
'        pmIntPage = pmIntPage + 1
'    End If
'sbRdFile16_Exit:
'    Exit Sub
'
'sbRdFile16_Err:
'
'    If bTrans = True Then
'        Call NBG_MsgBox("Κρίσιμο Λάθος - Rollback Updates !!!", True)
'        wks.Rollback    ' Page is now Released and changes are rolled back in case of failure
'        bTrans = False
'    Else
'        Call NBG_MsgBox("Error :" & error$, True)
'    End If
'
'    Resume sbRdFile16_Exit
'
'
'End Sub

'Public Sub initialize_files()
'
''    Set wks = DBEngine.Workspaces(0)
''    Set db = wks.OpenDatabase(dbFile)
'
''    Set rsConstants = db.OpenRecordset("tbl_Constants", dbOpenDynaset)
'    Set rsLogCounters = db.OpenRecordset("tbl_LogCounters", dbOpenDynaset)
'    Set rsBatch = db.OpenRecordset("tbl_Batch", dbOpenDynaset)
'    Set rsTransData = db.OpenRecordset("tbl_TransData", dbOpenDynaset)
'    Set rsFile16 = db.OpenRecordset("tbl_File16 ", dbOpenDynaset)
'
''    cBRANCH = fnReadConst("BranchId")
''    cBRANCH_NAME = fnReadConst("BranchName")
''    cTERMINAL = fnReadConst("TerminalID")
''    cHEAD = fnReadConst("Head")
''    cHMNIA = fnReadConst("DateCashier")
''    cPRINTERPOSITION = fnReadConst("PrinterPosition")
'    Call sbDeleteFile16
'
'End Sub

'Public Function ProperLogonStatus(ByVal sTransNum As String) As Boolean
'Dim sCriterion As String
'Dim NumKeysRequired As Integer
'Dim TellerRequired As Integer
'Dim ChiefTellerRequired As Integer
'Dim ManagerRequired As Integer
'Dim Message As String
'
'    ProperLogonStatus = False
'    rsTransData.MoveFirst
'
'    sCriterion = "Transaction = '" & sTransNum & "'"
'    rsTransData.FindFirst sCriterion
'
'    If rsTransData!Transaction <> sTransNum Then
'        Call NBG_MsgBox("Δεν υπάρχει η Συναλλαγή " & sTransNum, True)
'        Exit Function
'    End If
'
'    cb.Caption = rsTransData!TransactionName
'    cb.CodePage = rsTransData!UCS
'
'    NumKeysRequired = rsTransData!NumKeys
'    TellerRequired = rsTransData!Teller
'    ChiefTellerRequired = rsTransData!ChiefTeller
'    ManagerRequired = rsTransData!Manager
'
'    If NumKeysRequired = 0 Then
'        Call NBG_MsgBox("Δεν έχουν δοθεί δικαιοδοσίες για τη Συναλλαγή " & sTransNum, True)
'        Exit Function
'    End If
'
'    Select Case NumKeysRequired
'        Case 1
'            If (cb.TellerLogon + cb.ChiefTellerLogon + cb.ManagerLogon <> 1) Then
'                Message = "Πρέπει να ανοίξει 1 χρήστης: " & _
'                   IIf(TellerRequired = 1, "Teller ή", "") & _
'                   IIf(ChiefTellerRequired = 1, " Chief Teller ή", "") & _
'                   IIf(ManagerRequired = 1, " Manager ή", "")
'                Message = Mid(Message, 1, Len(Message) - 1)
'                ProperLogonStatus = False
'            Else
'                If (cb.TellerLogon = 1 And TellerRequired = 1) Or _
'                    (cb.ChiefTellerLogon = 1 And ChiefTellerRequired = 1) Or _
'                    (cb.ManagerLogon = 1 And ManagerRequired = 1) Then
'                    ProperLogonStatus = True
'                Else
'                    Message = "Πρέπει να ανοίξει 1 χρήστης: " & _
'                       IIf(TellerRequired = 1, "Teller ή", "") & _
'                       IIf(ChiefTellerRequired = 1, " Chief Teller ή", "") & _
'                       IIf(ManagerRequired = 1, " Manager ή", "")
'                    Message = Mid(Message, 1, Len(Message) - 1)
'                    ProperLogonStatus = False
'                End If
'            End If
'        Case 2
'            If (cb.TellerLogon + cb.ChiefTellerLogon + cb.ManagerLogon <> 2) Then
'                Message = "Πρέπει να ανοίξουν 2 χρήστες: " & _
'                    IIf(TellerRequired = 1, "  Teller και", "") & _
'                    IIf(ChiefTellerRequired = 1, " Chief Teller και", "") & _
'                    IIf(ManagerRequired = 1, " Manager και", "")
'                Message = Mid(Message, 1, Len(Message) - 3)
'                ProperLogonStatus = False
'            Else
'                If (cb.TellerLogon = 1 And TellerRequired <> 1) Or _
'                    (cb.ChiefTellerLogon = 1 And ChiefTellerRequired <> 1) Or _
'                    (cb.ManagerLogon = 1 And ManagerRequired <> 1) Then
'                    Message = "Πρέπει να ανοίξουν 2 χρήστες: " & _
'                        IIf(TellerRequired = 1, "  Teller και", "") & _
'                        IIf(ChiefTellerRequired = 1, "  Chief Teller και", "") & _
'                        IIf(ManagerRequired = 1, " Manager και", "")
'                    Message = Mid(Message, 1, Len(Message) - 1)
'                    ProperLogonStatus = False
'                Else
'                    ProperLogonStatus = True
'                End If
'            End If
'        Case 3
'            If (cb.TellerLogon + cb.ChiefTellerLogon + cb.ManagerLogon <> 3) Then
'                Message = "Πρέπει να ανοίξουν Teller και Chief Teller και Manager "
'                ProperLogonStatus = False
'            Else
'                ProperLogonStatus = True
'            End If
'    End Select
'
'    If ProperLogonStatus = False Then
'        Call NBG_MsgBox(Message, True)
'    End If
'
'End Function
