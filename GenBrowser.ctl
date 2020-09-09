VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl GenBrowser 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin SHDocVwCtl.WebBrowser vControl 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   5318
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "GenBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private owner As Form
Private DisplayFlag(10) As Boolean
Private BrowserName As String
Public ScrLeft As Long, ScrTop As Long, ScrWidth As Long, ScrHeight As Long
Private ValidationControl As ScriptControl
Public inprogress As Boolean
Public htmlReport As cHTMLReport

Public Property Get Control()
    Set Control = VControl
End Property

Private Sub UserControl_Resize()
    VControl.Left = 0: VControl.Top = 0: VControl.width = width: VControl.height = height
End Sub

Public Function IsVisible(inPhase) As Boolean
    IsVisible = DisplayFlag(CInt(inPhase))
End Function

Public Sub navigate(URL As String)
    inprogress = True
    VControl.navigate URL
End Sub


Public Sub Initialize(inOwner As Form, inProcessControl As ScriptControl, Name As String, _
    wLeft As Long, wTop As Long, wWidth As Long, wHeight As Long)
Dim i As Integer
    Set owner = inOwner
    Set ValidationControl = inProcessControl
    BrowserName = Name
    ValidationControl.AddObject Name, Me, True
    
    For i = 1 To 10
        DisplayFlag(i) = True
    Next i
 
    ScrLeft = wLeft
    ScrWidth = wWidth
    ScrTop = wTop * 290
    ScrHeight = wHeight * 285
        
End Sub

Private Sub vControl_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    inprogress = False
End Sub

Public Sub ProcessHTMLReport(htmlfilename As String, data)
    If htmlReport Is Nothing Then
        Set htmlReport = New cHTMLReport
        htmlReport.PrepareFromGenBrowser Me
    End If
    
    Dim anode As IXMLDOMNode
    Set anode = data
    
    htmlReport.ProcessReport htmlfilename, anode

End Sub

Public Sub PrintDocument()
    VControl.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT, Null, Null

End Sub

Public Sub PreviewDocument()
    VControl.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT, Null, Null

End Sub




