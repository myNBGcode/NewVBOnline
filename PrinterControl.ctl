VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl PrinterControl 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   ScaleHeight     =   1515
   ScaleWidth      =   2145
   Begin MSCommLib.MSComm ComControl 
      Left            =   960
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "PrinterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Printer Type
Const ptNONE = 0
Const ptOLIVETTI = 3
Const ptSIEMENS = 4
Const Set928 = "ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙáâãäåæçèéêëìíîïðñóòôõö÷ø"
Const Set928Ext = "ùÜÝÞúßüýþ¢¸¹º¼¾¿"
Const Set437 = "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬­®¯"

Public CurrentPrinter As Integer
Private MsgCounter As Long

Public Sub Initialize()
    If CurrentPrinter = ptNONE Then Exit Sub
Dim abyte As Byte
    
    If CurrentPrinter = ptSIEMENS Then
        With ComControl
            .CommPort = 1
            .Handshaking = comXOnXoff
            .RThreshold = 1
            .RTSEnable = True
            .Settings = "9600,e,8,1"
            .SThreshold = 1
            .OutBufferSize = 1
            .PortOpen = True
        End With
'        ComControl.Output = Chr(27) & "6"
 '       ComControl.Output = Chr(27) & "t" & Chr(1)
        
  '      ComControl.Output = Chr(27) & "0"
    '    abyte = 115
   '     ComControl.Output = Chr(27) & "U" & Chr(abyte)
     '   ComControl.Output = Chr(27) & "."
'        ComControl.Output = Chr(27) & "J" & Chr(36)
    ElseIf CurrentPrinter = ptOLIVETTI Then
        With ComControl
            .CommPort = 1
            .Handshaking = comRTS
            .RThreshold = 1
            .RTSEnable = True
            .Settings = "9600,e,8,1"
            .SThreshold = 1
            .OutBufferSize = 1
            .PortOpen = True
        End With
        ComControl.Output = Chr(27) & Chr(51) & Chr(43)
    End If
    MsgCounter = 0
End Sub

Public Sub ClosePrinter(Cancel As Integer)
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.PortOpen = False
End Sub

Public Sub LF(aLFNo As Integer)
    If CurrentPrinter = ptNONE Then Exit Sub
   
    Dim i As Integer
    i = 1
    Do Until i > aLFNo
        ComControl.Output = Chr(10) & Chr(13)
        i = i + 1
    Loop
End Sub

Public Sub FF()
    If CurrentPrinter = ptNONE Then Exit Sub
   
    ComControl.Output = Chr(12)
End Sub

Public Sub PrintLine(PrnString As String)
    If CurrentPrinter = ptNONE Then
        Exit Sub
    ElseIf CurrentPrinter = ptOLIVETTI Then
        Dim astr As String, achar As String, i As Integer, apos As Integer
        astr = ""
        For i = 1 To Len(PrnString)
            achar = Mid(PrnString, i, 1)
            apos = InStr(1, Set928, achar, vbBinaryCompare)
            If apos > 0 Then
                astr = astr & Mid(Set437, apos, 1)
            Else
                apos = InStr(1, Set928Ext, achar, vbBinaryCompare)
                If apos > 0 Then
                    astr = astr + Chr(223 + apos)
                Else
                    astr = astr & achar
                End If
            End If
        Next i
        PrnString = astr
    End If
    ComControl.Output = PrnString
End Sub


Public Sub PrintLf(PrnString As String)
    If CurrentPrinter = ptNONE Then Exit Sub
    PrintLine PrnString
    LF (1)
End Sub

Public Sub BoldOn()
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "("
End Sub

Public Sub BoldOff()
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & ")"
End Sub

Public Sub SupSubOn(aFlag As String)
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "`" & aFlag
End Sub

Public Sub SupSubOff()
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "{"
End Sub

Public Sub UnderlineOn()
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "*" & "0"
End Sub

Public Sub UnderlineOff()
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "+"
End Sub

Public Sub LeftMargin(aLM As String)
    If CurrentPrinter = ptNONE Then Exit Sub
    ComControl.Output = Chr(27) & "J" & aLM
End Sub

Public Function OutBufferCount() As Integer
    OutBufferCount = ComControl.OutBufferCount
End Function

Public Function InBuffer() As String
    InBuffer = ComControl.Input
End Function

'Public Sub ComControl_OnComm()
'   Select Case ComControl.CommEvent
'   ' Handle each event or error by placing
'   ' code below each case statement
'
'   ' Errors
'      Case comEventBreak   ' A Break was received.
'        aMsg = MsgBox("error break")
'      Case comEventFrame   ' Framing Error
'        aMsg = MsgBox("error Frame")
'      Case comEventOverrun   ' Data Lost.
'        aMsg = MsgBox("error Overrun")
'      Case comEventRxOver   ' Receive buffer overflow.
'        aMsg = MsgBox("error RxOver")
'      Case comEventRxParity   ' Parity Error.
'        aMsg = MsgBox("error RxParity")
'      Case comEventTxFull   ' Transmit buffer full.
'        aMsg = MsgBox("error TxFull")
'      Case comEventDCB   ' Unexpected error retrieving DCB]
'        aMsg = MsgBox("error DCB")
'
'   ' Events
'      Case comEvCD   ' Change in the CD line.
'        aMsg = MsgBox("EvCD")
'      Case comEvCTS   ' Change in the CTS line.
'        aMsg = MsgBox("error EvCTS")
'      Case comEvDSR   ' Change in the DSR line.
'        aMsg = MsgBox("error EvDSR")
'      Case comEvRing   ' Change in the Ring Indicator.
'        aMsg = MsgBox("error EvRing")
'      Case comEvReceive   ' Received RThreshold # of
'                        ' chars.
'        astr = ComControl.Input
'        astr = CStr(MsgCounter) + " " + "Rcv1:" + astr
'        MsgCounter = MsgCounter + 1
'        MsgBox astr
'
''        TxtRcv.Text = MSComm1.Input
'      Case comEvSend   ' There are SThreshold number of
'                     ' characters in the transmit
'                     ' buffer.
''        amsg = MsgBox("Send " & TxtSend.Text)
''        Text1.Text = TxtSend.Text
''        Text3.Text = MSComm1.OutBufferCount
'
'      Case comEvEOF   ' An EOF charater was found in
'                     ' the input stream
'   End Select
'End Sub
