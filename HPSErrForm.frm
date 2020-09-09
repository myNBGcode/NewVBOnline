VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form HPSErrForm 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox ErrFld 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"HPSErrForm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "HPSErrForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ErrBuffer, StructName As String

Private Sub DisplayError(inPart, idx As Long, Optional From)
Dim i As Integer, apos As Integer, astr As String, ErrID As String, ErrStr As String, anode
    If SkipCRAUse Then
        MsgBox "Δεν υποστηρίζεται η λειτουργία από το σύστημα. (DisplayError)", vbCritical: Exit Sub
    End If
    
    ErrID = Trim(inPart.ByName("X_SET_NM", idx).value) & Trim(CStr(inPart.ByName("N_ERR", idx).value))
    On Error Resume Next
    Set anode = xmlCRAErrors.documentElement.selectSingleNode(ErrID)
    If Not (anode Is Nothing) Then
        ErrStr = anode.selectSingleNode("D").Text
'        ErrList.AddItem aNode.childnodes.Item("B").Text
        Caption = anode.selectSingleNode("T").Text & " - " & anode.selectSingleNode("U").Text
    Else
        ErrFld.Text = ErrFld.Text & CStr(inPart.ByName("N_ERR", idx).value) & vbCrLf
        ErrFld.Text = ErrFld.Text & inPart.ByName("X_SET_NM", idx).value & vbCrLf
'        ErrList.AddItem CStr(inPart.ByName("N_ERR", idx).Value)
'        ErrList.AddItem inPart.ByName("X_SET_NM", idx).Value
    
    End If
    
'    ErrList.AddItem ErrStr
'    ErrList.AddItem "------------------"
    With inPart
        If IsMissing(From) Then From = "VCUER04"
        For i = 1 To inPart.ByName(From, idx).Times
            apos = InStr(ErrStr, "%" & CStr(i))
'            ErrList.AddItem inPart.ByName("VCUER04", idx).ByName("T_ERR_PARM_1", i).Value
            If apos > 0 Then _
                ErrStr = Left(ErrStr, apos - 1) & inPart.ByName(From, idx).ByName("T_ERR_PARM_1", i).value & Right(ErrStr, Len(ErrStr) - apos - 1)
            
        Next i
    End With
'    ErrList.AddItem "------------------"
    ErrFld.Text = ErrFld.Text & ErrStr & vbCrLf
    
'    ErrList.AddItem ErrStr
End Sub

Private Sub Form_Activate()
Dim i As Integer, apos As Integer, ars As New ADODB.Recordset, astr As String, ErrStr As String
    
If UCase(ErrBuffer.name) = "VCUER01" Then
    DisplayError ErrBuffer, 1
ElseIf UCase(ErrBuffer.name) = "CUF_ERR_MSG_D" Then
    'ErrBuffer.name = "VCUER01"
    DisplayError ErrBuffer, 1, "CUF_ERR_MSG_PARMS"
ElseIf UCase(ErrBuffer.name) = "VCUER05" Then
    For i = 1 To ErrBuffer.Times
        If ErrBuffer.ByName("VCUER01", i).ByName("N_ERR", 1).value > 0 Then DisplayError ErrBuffer.ByName("VCUER01", i), CLng(i)
    Next i
ElseIf UCase(ErrBuffer.name) = "CUF_ERR_OCC_D" Then
    ErrBuffer.name = "VCUER05"
    For i = 1 To ErrBuffer.Times
        If ErrBuffer.ByName("VCUER01", i).ByName("N_ERR", 1).value > 0 Then DisplayError ErrBuffer.ByName("VCUER01", i), CLng(i), "CUF_ERR_MSG_PARMS"
    Next i
Else
    For i = 1 To ErrBuffer.ByName(StructName).Times
        If ErrBuffer.ByName(StructName).ByName("N_ERR", i).value > 0 Then DisplayError ErrBuffer.ByName(StructName), CLng(i)
    Next i
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set ErrBuffer = Nothing
End Sub
