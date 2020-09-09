VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form IRISMsgFrm 
   Caption         =   "лгмулата ияис"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton okBtn 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox MsgFld 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"IRISErrFrm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "IRISMsgFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MsgView
Public MsgViewXML As IXMLDOMElement

Public Sub ShowMessages(inMsgView)
    Dim i As Integer, astr As String
    
    MsgFld.Text = ""
    With inMsgView
       If .v2Value("STD_DEC_3") > 0 Then
          MsgFld.Text = MsgFld.Text & "аяихлос...........................пеяицяажг.......................йыд. " & vbCrLf
          For i = 1 To .v2Value("STD_DEC_3")
             If i > 5 Then Exit For
             Select Case .v2Value("COD_ANTCN", i)
             Case 1:      astr = "емглеяысг"
             Case 2:      astr = "еидопоигсг"
             Case 3:      astr = "уповяеытийг емеяцеиа"
             Case 4:      astr = "пяоеидопоигсг"
             End Select

             astr = gFormat_("%10ST% %20ST%   %30ST%   %2ST% %11ST% %7ST% ", Array( _
                format("0000000000", .v2Value("NUMERO_ANTCN", i)), astr, .v2Value("STD_CHAR_30", i), _
                .v2Value("SUBCD_ANTCN", i), .v2Value("STD_CHAR_11", i), .v2Value("STD_CHAR_07", i)))
             MsgFld.Text = MsgFld.Text & astr & vbCrLf
          Next
       End If
    End With
End Sub

Public Sub ShowMessagesXML(inMsgView As IXMLDOMElement)
    Dim i As Integer, astr As String
    Dim MsgElement As IXMLDOMNode
    Dim MsgElementList As IXMLDOMNodeList
    Dim COD_ANTCN As IXMLDOMNode, a1str As String
    Dim STD_CHAR_11 As IXMLDOMNode, a2str As String
    Dim STD_CHAR_07 As IXMLDOMNode, a3str As String
    Dim SUBCD_ANTCN As IXMLDOMNode, a4str As String
    Dim STD_CHAR_30 As IXMLDOMNode, a5str As String
    Dim NUMERO_ANTCN As IXMLDOMNode, a6str As String
    
    MsgFld.Text = ""
    Set MsgElementList = inMsgView.SelectNodes("//STD_AN_AV_MSJ_LS[COD_ANTCN!='']")
    If Not (MsgElementList Is Nothing) Then
       If MsgElementList.length > 0 Then
        MsgFld.Text = MsgFld.Text & "аяихлос...........................пеяицяажг.......................йыд. " & vbCrLf
        For i = 0 To MsgElementList.length - 1
          If i > 5 Then Exit For
          Set MsgElement = MsgElementList(i)
          Set COD_ANTCN = MsgElement.selectSingleNode("COD_ANTCN")
          If Not (COD_ANTCN Is Nothing) Then
            Select Case COD_ANTCN.Text
                Case 1:      a1str = "емглеяысг"
                Case 2:      a1str = "еидопоигсг"
                Case 3:      a1str = "уповяеытийг емеяцеиа"
                Case 4:      a1str = "пяоеидопоигсг"
            End Select
          End If
          Set NUMERO_ANTCN = MsgElement.selectSingleNode("NUMERO_ANTCN")
          If Not (NUMERO_ANTCN Is Nothing) Then
            a6str = NUMERO_ANTCN.Text
          End If
          Set STD_CHAR_30 = MsgElement.selectSingleNode("STD_DESCR_C_ANTCN_V/STD_CHAR_30")
          If Not (STD_CHAR_30 Is Nothing) Then
            a5str = STD_CHAR_30.Text
          End If
          Set SUBCD_ANTCN = MsgElement.selectSingleNode("SUBCD_ANTCN")
          If Not (SUBCD_ANTCN Is Nothing) Then
            a4str = SUBCD_ANTCN.Text
          End If
          Set STD_CHAR_11 = MsgElement.selectSingleNode("DESCRIP_ANTCN_V/STD_CHAR_11")
          If Not (STD_CHAR_11 Is Nothing) Then
            a2str = STD_CHAR_11.Text
          End If
          Set STD_CHAR_07 = MsgElement.selectSingleNode("DESC_IND_PRDAD_V/STD_CHAR_07")
          If Not (STD_CHAR_07 Is Nothing) Then
            a3str = STD_CHAR_07.Text
          End If
          astr = gFormat_("%10ST% %20ST%   %30ST%   %2ST% %11ST% %7ST% ", Array( _
                format("0000000000", a6str), a1str, a5str, a4str, a2str, a3str))
                MsgFld.Text = MsgFld.Text & astr & vbCrLf
        Next i
       End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If Not (MsgView Is Nothing) Then
        ShowMessages MsgView
    ElseIf Not (MsgViewXML Is Nothing) Then
        ShowMessagesXML MsgViewXML
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MsgView = Nothing
    Set MsgViewXML = Nothing
End Sub

Private Sub okBtn_Click()
    Unload Me
End Sub
