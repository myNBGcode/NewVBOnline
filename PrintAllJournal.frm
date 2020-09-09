VERSION 5.00
Begin VB.Form PrintAllJournal 
   Caption         =   "Εκτύπωση Ημερολογίου"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel_CMD 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton ok_Cmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox PrinterList 
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Επιλογή Εκτυπωτή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "PrintAllJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WorkDate As Date
Private JournalLength As Integer

Public Sub UpdateForm()

Dim x As Printer, i As Integer
i = 0
For Each x In Printers
   PrinterList.AddItem x.DeviceName
   If x.DeviceName = Printer.DeviceName Then PrinterList.Selected(PrinterList.ListCount - 1) = True
   i = i + 1
Next
If i > 0 Then PrinterList.ListIndex = 0
JournalLength = 0

End Sub

Private Sub Cancel_CMD_Click()
Dim i As Integer
    For i = PrinterNameList.count To 1 Step -1
        PrinterNameList.Remove i
    Next i
    Unload Me
End Sub

Private Sub Form_Load()
    CenterFormOnScreen Me
End Sub

Private Sub ok_Cmd_Click()
Dim aPrinterName As String, i As Integer
    For i = PrinterNameList.count To 1 Step -1
        PrinterNameList.Remove i
    Next i
    If PrinterList.SelCount > 0 Then
        For i = 0 To PrinterList.ListCount - 1
            If PrinterList.Selected(i) Then PrinterNameList.Add PrinterList.List(i)
        Next i
    End If
    Unload Me
End Sub
