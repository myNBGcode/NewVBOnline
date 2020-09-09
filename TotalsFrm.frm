VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form TotalsFrm 
   Caption         =   "Σύνολα Αθροιστών"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4470
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   4140
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   582
      ButtonWidth     =   2143
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "CommandImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Εκτύπωση"
            Key             =   "PRINT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Επιστροφή"
            Key             =   "RETURN"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList CommandImages 
      Left            =   2730
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TotalsFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TotalsFrm.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TotalsFrm.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vTotals 
      Height          =   2145
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   100
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
End
Attribute VB_Name = "TotalsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim i As Integer
    i = KeyAscii
    If i = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Top = GenWorkForm.Top + 1000
    Left = GenWorkForm.Left
    width = GenWorkForm.width
    height = GenWorkForm.height - 1000
    
Dim res As Boolean
    res = fnDisplayTotals(vTotals)
    
End Sub

Private Sub Form_Resize()
    With vTotals
        .Left = 0
        .Top = 0
        .width = Me.ScaleWidth
        .height = Toolbar1.Top
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "PRINT" Then
ElseIf Button.Key = "RETURN" Then
    Unload Me
End If
End Sub
