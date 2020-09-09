VERSION 5.00
Begin VB.Form SelectPrinterFrm 
   Caption         =   "Επιλογή Εκτυπωτή"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ok_Cmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Cancel_CMD 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox PrinterList 
      Height          =   2985
      Left            =   120
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "SelectPrinterFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WorkDate As Date
Private JournalLength As Integer
Public SelectedPrinter As String

Public Sub UpdateForm()

Dim x As Printer, i As Integer
i = 0
PrinterList.AddItem "Passbook"
PrinterList.ListIndex = 0
For Each x In Printers
    PrinterList.AddItem x.DeviceName
    If UCase(x.DeviceName) = UCase(Printer.DeviceName) Then
        PrinterList.ListIndex = PrinterList.ListCount - 1
    End If
    i = i + 1
Next

End Sub

Private Sub Cancel_CMD_Click()
    SelectedPrinter = "": Hide
End Sub

Private Sub Form_Activate()
    CenterFormOnScreen Me
    'UpdateForm
End Sub

Private Sub Form_Load()
    UpdateForm
End Sub

Private Sub ok_Cmd_Click()
    SelectedPrinter = PrinterList.Text: Hide
End Sub
