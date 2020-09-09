Attribute VB_Name = "Security"
Option Explicit
'Public Function require_account(account_type As String) As Boolean
'
'require_account = False
'If Screen.ActiveForm.Name = "frmMenu" Then
'    cb.curr_transaction = ""
'End If
'Select Case account_type
'    Case "TELLER"
'        frmPassword.Caption = "Teller Logon"
'    Case "CHIEF_TELLER"
'        frmPassword.Caption = "Chief-Teller Logon"
'    Case "MANAGER"
'        frmPassword.Caption = "Manager Logon"
'End Select
'
'frmPassword.Show 1
'
'End Function

'Public Sub InitializeLogonStatus(frmCurrent As Form)
'    If cb.TellerLogon = 1 Then
'        frmCurrent.TellerLogon.ForeColor = &H8000&
'    Else
'        frmCurrent.TellerLogon.ForeColor = &HFF&
'    End If
'
'    If cb.ChiefTellerLogon = 1 Then
'        frmCurrent.ChiefTellerLogon.ForeColor = &H8000&
'    Else
'        frmCurrent.ChiefTellerLogon.ForeColor = &HFF&
'    End If
'
'    If cb.ManagerLogon = 1 Then
'        frmCurrent.ManagerLogon.ForeColor = &H8000&
'    Else
'        frmCurrent.ManagerLogon.ForeColor = &HFF&
'    End If
'
'End Sub

'Public Sub ChangeTellerState(frmCurrent As Form)
'Dim OK As Boolean
'
'OK = False
'If cb.TellerLogon = 1 Then
'    cb.TellerLogon = 0
'    frmCurrent.TellerLogon.ForeColor = &HFF&
'Else
'    OK = require_account("TELLER")
'    DoEvents
'    If cb.TellerLogon = 1 Then
'        frmCurrent.TellerLogon.ForeColor = &H8000&
'    End If
'End If
'
'End Sub

'Public Sub ChangeChiefTellerState(frmCurrent As Form)
'Dim OK As Boolean
'
'OK = False
'If cb.ChiefTellerLogon = 1 Then
'    cb.ChiefTellerLogon = 0
'    frmCurrent.ChiefTellerLogon.ForeColor = &HFF&
'Else
'    OK = require_account("CHIEF_TELLER")
'    DoEvents
'    If cb.ChiefTellerLogon = 1 Then
'        frmCurrent.ChiefTellerLogon.ForeColor = &H8000&
'    End If
'End If
'
'End Sub
'Public Sub ChangeManagerState(frmCurrent As Form)
'Dim OK As Boolean
'
'OK = False
'If cb.ManagerLogon = 1 Then
'    cb.ManagerLogon = 0
'    frmCurrent.ManagerLogon.ForeColor = &HFF&
'Else
'    OK = require_account("MANAGER")
'    DoEvents
'    If cb.ManagerLogon = 1 Then
'        frmCurrent.ManagerLogon.ForeColor = &H8000&
'    End If
'End If
'
'End Sub

'Public Sub ChangeLogonState(frmCurrent As Form, KeyCode As Integer)
'Select Case KeyCode
'    Case vbKeyF1:
'        Call ChangeTellerState(frmCurrent)
'    Case vbKeyF2:
'        Call ChangeChiefTellerState(frmCurrent)
'    Case vbKeyF3:
'        Call ChangeManagerState(frmCurrent)
'End Select
'End Sub
