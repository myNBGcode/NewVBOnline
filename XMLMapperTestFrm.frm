VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aXMLMapperobj As XMLMapperInProc.XMLMapperLinkIn

Private Sub Command1_Click()
Dim soapclient As soapclient
Set soapclient = New soapclient
   Dim ares, bres
'    soapclient.mssoapinit "http://n00350030/XMLMapperListen/XMLMapperDebug.wsdl", "", "", ""
    soapclient.mssoapinit "D:\RecordsetMapper\DelphiVersion4\XMLMapperDebug.wsdl", "", "", ""
    
    soapclient.GetFullTaggedRowsXML "File Name=c:\RecordsetMapper.udl", "select top 10 * from tbl_totals", ares
    
    
   Set aXMLMapperobj = New XMLMapperInProc.XMLMapperLinkIn
   Dim aRSView As New XMLRecordsetView
   aRSView.Prepare Nothing, aXMLMapperobj
  
   
   aRSView.AddNew
   aRSView.Fields("name").Value = "DTotal1"
   aRSView.Fields("TerminalID").Value = "TEST"
   aRSView.Fields("CD").Value = 1
   aRSView.Fields("Currency").Value = 0
   aRSView.DeleteRow
   
   aRSView.ExecInsertX "File Name=c:\RecordsetMapper.udl", "select top 1 * from tbl_totals where name ='xxxx'"
   
   aRSView.ReadXMLMapperX "File Name=c:\RecordsetMapper.udl", "select top 10 * from tbl_totals"
   If aRSView.RecordCount > 0 Then
        aRSView.MoveFirst
        Do
            ares = aRSView.Fields("name2").Value
'            ares = aRSView.Fields("name").Value
            aRSView.MoveNext
        Loop Until aRSView.Eof
   End If
   
   
   Exit Sub
   
   aXMLMapperobj.GetFullTaggedRowsXML "File Name=c:\RecordsetMapper.udl", "select top 10 * from tbl_totals", ares
   
   MsgBox ares
   Exit Sub
   
   'Dim aRSView As XMLRecordsetView
   Set aRSView = New XMLRecordsetView
   aRSView.Prepare Nothing, aXMLMapperobj
   
   
   
   aRSView.ReadXMLMapperX "File Name=c:\RecordsetMapper", "select top 10 * from Cs"

End Sub
