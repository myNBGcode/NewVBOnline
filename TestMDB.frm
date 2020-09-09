VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   540
      TabIndex        =   0
      Top             =   570
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ado_db As New ADODB.Connection
    ado_db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=n:\test2000.mdb"
Dim cmdstr As String
Dim ARS As ADODB.Recordset
    cmdstr = "Insert into Journal(LUName, TerminalID, UName, ChiefUName, ManagerUName, " & _
        " PostDate, TRNCount, TrnCode, FldNo, FldTitle, DataLine) " & _
        " Values('W123', '012B', 'u34000', '____', '____', #" & Format(Date, "yyyy/mm/dd") & "#, 1200, '1200', 0, '', 'safd s kjhl h hkjh  hjhkl hjjhjhljhlkjhklj hj jhkljhljhlkhlk jh lhajfdsadsaj')"
Dim i As Integer, res As Integer
    For i = 1 To 1000
        ado_db.Execute cmdstr, res
        Set ARS = New ADODB.Recordset
        ARS.Open "SELECT * FROM Journal", ado_db, adOpenStatic, adLockReadOnly
        ARS.MoveLast
        ARS.Close
        Set ARS = Nothing
    Next i

End Sub

