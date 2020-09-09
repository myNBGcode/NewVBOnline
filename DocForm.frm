VERSION 5.00
Begin VB.Form DocForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Εισαγωγή Παραστατικού"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Enter για Συνέχεια"
      Default         =   -1  'True
      Height          =   405
      Left            =   3750
      TabIndex        =   1
      Top             =   1560
      Width           =   2805
   End
   Begin VB.ListBox G0ListBox 
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6495
   End
End
Attribute VB_Name = "DocForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    ShowScrollBar (200)
End Sub

Private Sub Form_Load()
Dim i As Integer, astr As String
    For i = 1 To G0Data.count
        G0ListBox.AddItem G0Data.Item(i)
    Next i
    Caption = PrintMsg
End Sub

Public Sub ShowScrollBar(Optional maxlinelength As Integer)
    If IsMissing(maxlinelength) Then maxlinelength = 200
    Dim ascalemode As Integer, aSize As Integer, astr As String
    ScaleMode = 3
    astr = String(maxlinelength, "W")
    aSize = TextWidth(astr)
    SendMessage G0ListBox.hWnd, LB_SETHORIZONTALEXTENT, aSize, 0
End Sub


