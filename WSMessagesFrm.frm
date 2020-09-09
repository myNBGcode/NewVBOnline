VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form WSMessagesFrm 
   Caption         =   "WSMessagesFrm"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog OpenFileDialog 
      Left            =   1920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "WSMessagesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoc, aNSManager

Public Sub BuildWSStruct(astructdoc, aname As String)
    Dim bdoc As MSXML2.DOMDocument30, aNSManager
    Dim Message As MSXML2.IXMLDOMNode
    Set aNSManager = CreateObject("Msxml2.MXNamespaceManager.4.0")
    Set Message = astructdoc.documentElement.selectSingleNode("//*[namespace-uri() = 'http://www.nbg.gr/vbonline/mdl' and local-name() = 'message' and @name = '" & aname & "' ]")
    If Message Is Nothing Then
    
    
    Else
        Dim Messageparts As MSXML2.IXMLDOMNodeList
        Set Messageparts = Message.SelectNodes("//*[namespace-uri() = 'http://www.nbg.gr/vbonline/mdl' and local-name()='part']")
        Dim Node As MSXML2.IXMLDOMNode
        For Each Node In Messageparts
            Dim aattr As IXMLDOMAttribute
            Dim partname As String, parttype As String, typenameprefix As String
            Set aattr = Node.Attributes.getNamedItem("name")
            If Not (aattr Is Nothing) Then partname = aattr.value
            
            Set aattr = Node.Attributes.getNamedItem("type")
            If Not (aattr Is Nothing) Then
                Dim aPos As Integer
                parttype = aattr.value:
                aPos = InStr(1, parttype, ":")
                If aPos > 1 Then
                    typenameprefix = Left(parttype, aPos - 1)
                    If aPos < Len(parttype) Then
                        parttype = Right(parttype, Len(parttype) - aPos)
                    Else
                        parttype = ""
                    End If
                    Set aattr = astructdoc.documentElement.Attributes.getNamedItem("xmlns:" & typenameprefix)
                    typenameprefix = aattr.value
                End If
            End If
            
            Dim MessagePart As MSXML2.IXMLDOMNode
            Set MessagePart = astructdoc.documentElement.selectSingleNode("//*[namespace-uri() = 'http://www.nbg.gr/vbonline/Type' and local-name() = 'types' and @tns='" & typenameprefix & "']/*[local-name() = 'type' and @name = '" & parttype & "' ]")
            If Not MessagePart Is Nothing Then
            
            
            End If
            
        Next Node
    
    End If
    
End Sub

Private Sub Command1_Click()
    OpenFileDialog.filename = ""
    OpenFileDialog.ShowOpen
    If OpenFileDialog.filename <> "" Then
        
        Set adoc = CreateObject("Msxml2.DOMDocument.6.0")
        Set aNSManager = CreateObject("Msxml2.MXNamespaceManager.4.0")
        If adoc.Load(OpenFileDialog.filename) Then
           Dim MessageList As MSXML2.IXMLDOMNodeList
           Dim aattr As IXMLDOMAttribute
           
           Set MessageList = adoc.documentElement.SelectNodes("//*[namespace-uri() = 'http://www.nbg.gr/vbonline/mdl' and local-name() = 'message']")
           
           If MessageList.length > 0 Then
                Dim Node As IXMLDOMNode
                For Each Node In MessageList
                    Set aattr = Node.Attributes.getNamedItem("name")
                    If aattr Is Nothing Then
                    
                    Else
                        List1.AddItem aattr.value
                    End If
                Next Node
           
           
           End If
        
        
        End If
    End If
    
End Sub

Private Sub List1_DblClick()
    Dim aname As String
    aname = List1.Text
    
    BuildWSStruct adoc, aname
End Sub
