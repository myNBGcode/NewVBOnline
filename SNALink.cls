VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SNALink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public RetStr As String

Public Sub SetApplID(aID As String)
    cb.ApplId = Left(aID & String(8, " "), 8) & Chr(0)
End Sub

Public Sub SetLUName(aName As String)
    cLUName = Left(aName & String(8, " "), 8) & Chr(0)
End Sub

Public Sub SetComputerName(aName As String)
    MachineName = Left(aName & String(9, " "), 9)
End Sub

Public Sub SetUserName(aName As String)
    cUserName = Left(aName & String(8, " "), 8)
End Sub

Public Function DoConnect() As Integer
    initialize_cb
    DoConnect = CONNECT
End Function

Public Function DoDisconnect() As Integer
    DoDisconnect = DISCONNECT_
End Function

Public Function ReceiveStr() As String
    ReceiveStr = cb.receive_str
End Function

Public Function HPSCommunicate2(Trn As String, Rule As String, InputName As String, OutputName As String, InputView As String, OutputViewLength As Long) As Integer
Dim astr As String, aSize As Long, res As Integer, aComputerName As String
    
    RetStr = ""
    
    
    aComputerName = MachineName
    astr = Left(aComputerName & "         ", 9) & Left(cUserName & "         ", 8) & "00000"


    Rule = Left(Rule & "        ", 8)

    InputName = Left(InputName & "        ", 8)

    OutputName = Left(OutputName & "        ", 8)


    astr = AsciiToEbcdic_(Trn & astr) & _
            AsciiToEbcdic_(Rule) & AsciiToEbcdic_(InputName) & AsciiToEbcdic_(OutputName) & _
            IntToHps_(Len(InputView)) & _
            IntToHps_(OutputViewLength)
    aSize = Len(InputView)
    
    If HPSSEND_(astr & InputView) = SEND_OK Then 'Me,
        If aSize < 4096 Then res = HPSRECEIVE_() Else res = RECEIVE_OK 'Me
        If res = RECEIVE_OK Then
            If cb.receive_str = "" Then
                HPSCommunicate2 = 999

                Exit Function
            Else
                RetStr = cb.receive_str
            End If
        Else
            HPSCommunicate2 = 999
        End If
    Else
        HPSCommunicate2 = 999
    End If
'мео FORMAT: TCFMO + NXXXXXXXX + EXXXXX__ + XXXXX (TRNNo) + .....
End Function





