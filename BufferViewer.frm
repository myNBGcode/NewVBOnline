VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BufferViewer 
   Caption         =   "Buffer Viewer"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1080
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   240
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BufferViewer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BufferViewer.frx":0542
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar FrmToolbar 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "CLEARVIEW"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IRISCLEAR"
                  Text            =   "Clear"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IRISCLEAREMPTYNODES"
                  Text            =   "Clear Empty Nodes"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IRISCLEAREMPTYSTRUCTS"
                  Text            =   "Clear Empty Structs"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "SAVEVIEW"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebViewer 
      CausesValidation=   0   'False
      Height          =   5415
      Left            =   4800
      TabIndex        =   1
      Top             =   720
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox BufferList 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Splitter 
      Height          =   1335
      Left            =   4440
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "BufferViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public owner, inBufferList As Buffers
Private MoveSplitter As Boolean
Private tmp2path As String

Private Sub BufferList_DblClick()
Dim astr As String, i As Long, res As Integer, xstr As String
Dim aXSLT As New MSXML2.DOMDocument30
Dim adoc As New MSXML2.DOMDocument30
Dim bdoc As New MSXML2.DOMDocument30
    BufferList.Enabled = False
    astr = BufferList.list(BufferList.ListIndex)
    If astr = "Nothing" Then Exit Sub
    i = InStr(astr, "(")
    If i > 0 Then astr = Trim(Left(astr, i - 1))
    With inBufferList.ByName(astr)
        .GetXMLView
        
        '.xmlDocV2.save App.path & "\tmp2.xml"
        'WebViewer.navigate App.path & "\tmp2.xml"
        SaveXmlFile "BufferViewer.xml", .xmlDocV2
        WebViewer.navigate tmp2path & "\BufferViewer.xml"
        
    End With
    BufferList.Enabled = True

End Sub

Private Sub Form_Activate()
Dim i As Long
    tmp2path = NetworkHomeDir
    For i = 1 To inBufferList.BufferNum
        If Not (inBufferList.ByIndex(i) Is Nothing) Then
            If inBufferList.ByIndex(i).LastLevel Then _
            BufferList.AddItem inBufferList.ByIndex(i).name & " (" & CStr(Len(inBufferList.ByIndex(i).Data)) & ")"
        Else
            BufferList.AddItem "Nothing"
        End If
    Next i
    MoveSplitter = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If MoveSplitter Then
        Splitter.Left = Splitter.Left + x
        Form_Resize
    End If
End Sub

Private Sub Form_Resize()
    BufferList.Move 0, FrmToolbar.height, Splitter.Left - 20, ScaleHeight - FrmToolbar.height
    WebViewer.Move Splitter.Left + 40, FrmToolbar.height, ScaleWidth - Splitter.Left - 60, ScaleHeight - FrmToolbar.height
    Splitter.Move Splitter.Left, FrmToolbar.height, 60, ScaleHeight - FrmToolbar.height
    
End Sub

Private Sub FrmToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim aXSLT As New MSXML2.DOMDocument30
Dim adoc As New MSXML2.DOMDocument30
Dim bdoc As New MSXML2.DOMDocument30
    If Button.Tag = "SAVEVIEW" Then
        CommonDialog.filename = ""
        CommonDialog.ShowSave
        If CommonDialog.filename <> "" Then
'            adoc.Load App.path & "\tmp2.xml"
'            adoc.save CommonDialog.filename
            adoc.Load tmp2path & "\BufferViewer.xml"
            adoc.save CommonDialog.filename
        End If
    End If
End Sub

Private Sub FrmToolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
Dim astr As String
    astr = ""
End Sub

Private Sub FrmToolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim aXSLT As New MSXML2.DOMDocument30
Dim adoc As New MSXML2.DOMDocument30
Dim bdoc As New MSXML2.DOMDocument30
    If ButtonMenu.key = "IRISCLEAR" Then
        aXSLT.LoadXML xmlXSLTPack.selectSingleNode("//IRISCLEAR/xsl:stylesheet").XML
        'adoc.Load App.path & "\tmp2.xml"
        adoc.Load tmp2path & "\BufferViewer.xml"
        If adoc.parseError.errorCode <> 0 Then
            'MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ App.Path & \tmp2.xml": Exit Sub
            MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ " & tmp2path & "\BufferViewer.xml": Exit Sub
        End If
        bdoc.LoadXML adoc.documentElement.transformNode(aXSLT.documentElement)
        'bdoc.save App.path & "\tmp2.xml"
        'WebViewer.navigate App.path & "\tmp2.xml"
        SaveXmlFile "BufferViewer.xml", bdoc
        WebViewer.navigate tmp2path & "\BufferViewer.xml"
        
    ElseIf ButtonMenu.key = "IRISCLEAREMPTYNODES" Then
        aXSLT.LoadXML xmlXSLTPack.selectSingleNode("//IRISCLEAREMPTYNODES/xsl:stylesheet").XML
        'adoc.Load App.path & "\tmp2.xml"
         adoc.Load tmp2path & "\BufferViewer.xml"
        If adoc.parseError.errorCode <> 0 Then
            'MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ App.Path & \tmp2.xml": Exit Sub
            MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ " & tmp2path & "\BufferViewer.xml": Exit Sub
        End If
        bdoc.LoadXML adoc.documentElement.transformNode(aXSLT.documentElement)
        'bdoc.save App.path & "\tmp2.xml"
        'WebViewer.navigate App.path & "\tmp2.xml"
        SaveXmlFile "BufferViewer.xml", bdoc
        WebViewer.navigate tmp2path & "\BufferViewer.xml"
        
    ElseIf ButtonMenu.key = "IRISCLEAREMPTYSTRUCTS" Then
        aXSLT.LoadXML xmlXSLTPack.selectSingleNode("//IRISCLEAREMPTYSTRUCTS/xsl:stylesheet").XML
        'adoc.Load App.path & "\tmp2.xml"
        adoc.Load tmp2path & "\BufferViewer.xml"
        If adoc.parseError.errorCode <> 0 Then
            'MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ App.Path & \tmp2.xml": Exit Sub
            MsgBox "–—œ¬À«Ã¡ ”‘«Õ ¡Õ¡ ‘«”« ‘œ’ " & tmp2path & "\BufferViewer.xml": Exit Sub
        End If
        bdoc.LoadXML adoc.documentElement.transformNode(aXSLT.documentElement)
        'bdoc.save App.path & "\tmp2.xml"
        'WebViewer.navigate App.path & "\tmp2.xml"
        SaveXmlFile "BufferViewer.xml", bdoc
        WebViewer.navigate tmp2path & "\BufferViewer.xml"
    End If
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MoveSplitter = (Button = vbLeftButton)
End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If MoveSplitter And x <> 0 Then
        Splitter.Left = Splitter.Left + x
        Form_Resize
    End If
End Sub

Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MoveSplitter = False
End Sub
