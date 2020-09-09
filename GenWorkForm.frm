VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form GenWorkForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Εφαρμογή OnLine Συναλλαγών"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GenWorkForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CommandFrame 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8745
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   1815
      Begin MSComctlLib.ImageList KeyImages 
         Left            =   900
         Top             =   6930
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":0CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":1138
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":158A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":19DC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar KeyBar 
         Height          =   810
         Left            =   0
         TabIndex        =   23
         Top             =   7740
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1429
         ButtonWidth     =   1032
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "KeyImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "T"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CT"
               ImageIndex      =   3
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "M"
               ImageIndex      =   3
               Style           =   1
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar CommandToolbar 
         Height          =   810
         Left            =   240
         TabIndex        =   21
         Top             =   2850
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1429
         ButtonWidth     =   2566
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "F7 Προηγούμενο"
               Key             =   "F7"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "F8 Επόμενο"
               Key             =   "F8"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "F10 Ημερολόγιο"
               Key             =   "F10"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "F11 Αθροιστές"
               Key             =   "F11"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Έναρξη / Λήξη"
               Key             =   "START"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   210
         Top             =   6960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":1E2E
               Key             =   "F7"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":2280
               Key             =   "F8"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":259A
               Key             =   "F10"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":29EC
               Key             =   "F11"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":2E3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":55F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenWorkForm.frx":5A42
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox shortkey 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   2250
         Width           =   1545
      End
      Begin VB.Image Image6 
         Height          =   585
         Left            =   120
         Picture         =   "GenWorkForm.frx":631C
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label ShortLabel 
         Caption         =   "Συναλλαγή:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "E.T.E."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image7 
         Height          =   1215
         Left            =   60
         Picture         =   "GenWorkForm.frx":92E6
         Stretch         =   -1  'True
         Top             =   750
         Width           =   1695
      End
      Begin VB.Image LogoImage 
         Height          =   915
         Left            =   60
         Picture         =   "GenWorkForm.frx":A98F
         Top             =   810
         Visible         =   0   'False
         Width           =   1590
      End
   End
   Begin MSComctlLib.StatusBar vStatus 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   9240
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19932
            MinWidth        =   19932
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "GenWorkForm.frx":B4BA
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "GenWorkForm.frx":B706
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame TitleFrame 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   660
         Left            =   90
         Picture         =   "GenWorkForm.frx":B95A
         Top             =   150
         Width           =   6735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "On Line Σύστημα Συναλλαγών"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   585
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   6495
      End
   End
   Begin TabDlg.SSTab SSTabControl 
      Height          =   8295
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Συναλλαγές"
      TabPicture(0)   =   "GenWorkForm.frx":1A14C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ZControl"
      Tab(0).Control(1)=   "ComTimer"
      Tab(0).Control(2)=   "MenuCommand(0)"
      Tab(0).Control(3)=   "MenuCommand(1)"
      Tab(0).Control(4)=   "MenuCommand(2)"
      Tab(0).Control(5)=   "MenuCommand(3)"
      Tab(0).Control(6)=   "MenuCommand(4)"
      Tab(0).Control(7)=   "MenuCommand(5)"
      Tab(0).Control(8)=   "MenuCommand(6)"
      Tab(0).Control(9)=   "MenuCommand(7)"
      Tab(0).Control(10)=   "MenuCommand(8)"
      Tab(0).Control(11)=   "MenuCommand(9)"
      Tab(0).Control(12)=   "TitleList"
      Tab(0).Control(13)=   "BackFrame"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Αθροιστές"
      TabPicture(1)   =   "GenWorkForm.frx":1A168
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TotalsGrid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ημερολόγιο"
      TabPicture(2)   =   "GenWorkForm.frx":1A184
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vJournal"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Παράμετροι"
      TabPicture(3)   =   "GenWorkForm.frx":1A1A0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GenBrowser1"
      Tab(3).Control(1)=   "Command3"
      Tab(3).Control(2)=   "RebuildComareaBtn"
      Tab(3).Control(3)=   "StationInfo"
      Tab(3).Control(4)=   "RebuildViewBtn"
      Tab(3).Control(5)=   "PrnTestCmd"
      Tab(3).Control(6)=   "SRJournalChk"
      Tab(3).Control(7)=   "ReceiveJournalWriteChk"
      Tab(3).Control(8)=   "SendJournalWriteChk"
      Tab(3).Control(9)=   "EventLogWriteChk"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "RECEIVE"
      TabPicture(4)   =   "GenWorkForm.frx":1A1BC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TrnInputBox"
      Tab(4).ControlCount=   1
      Begin shine.CompressZIt ZControl 
         Left            =   -72120
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin shine.GenBrowser GenBrowser1 
         Height          =   735
         Left            =   -69480
         TabIndex        =   34
         Top             =   6720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   -68040
         TabIndex        =   33
         Top             =   5640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton RebuildComareaBtn 
         Caption         =   "Rebuild ComArea"
         Height          =   495
         Left            =   -72240
         TabIndex        =   32
         Top             =   5520
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Timer ComTimer 
         Enabled         =   0   'False
         Left            =   -71040
         Top             =   1680
      End
      Begin VB.ListBox StationInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         ItemData        =   "GenWorkForm.frx":1A1D8
         Left            =   -74880
         List            =   "GenWorkForm.frx":1A1DA
         TabIndex        =   31
         Top             =   480
         Width           =   9495
      End
      Begin VB.CommandButton RebuildViewBtn 
         Caption         =   "Rebuild View"
         Height          =   495
         Left            =   -72240
         TabIndex        =   30
         Top             =   6120
         Width           =   2655
      End
      Begin VB.CommandButton PrnTestCmd 
         Appearance      =   0  'Flat
         Caption         =   "Test Εκτυπωτή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -68040
         TabIndex        =   28
         Top             =   6120
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.CheckBox SRJournalChk 
         Caption         =   "Εκτύπωση S/R στο ημερολόγιο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74760
         TabIndex        =   27
         Top             =   5760
         Width           =   2445
      End
      Begin VB.CheckBox ReceiveJournalWriteChk 
         Caption         =   "Καταγρφή Εισερχομένων στο Ημερολόγιο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74760
         TabIndex        =   26
         Top             =   6900
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox SendJournalWriteChk 
         Caption         =   "Καταγραφή Εξερχομένων στο Ημερολόγιο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -74760
         TabIndex        =   25
         Top             =   6360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox EventLogWriteChk 
         Caption         =   "Καταγραφή στο Event Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74760
         TabIndex        =   24
         Top             =   5370
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74730
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74730
         TabIndex        =   15
         Top             =   660
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74730
         TabIndex        =   14
         Top             =   990
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74730
         TabIndex        =   13
         Top             =   1350
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74730
         TabIndex        =   12
         Top             =   1710
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -74730
         TabIndex        =   11
         Top             =   2070
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -74730
         TabIndex        =   10
         Top             =   2430
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   -74730
         TabIndex        =   9
         Top             =   2790
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   -74730
         TabIndex        =   8
         Top             =   3150
         Width           =   1935
      End
      Begin VB.CommandButton MenuCommand 
         Appearance      =   0  'Flat
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   -74730
         TabIndex        =   7
         Top             =   3510
         Width           =   1935
      End
      Begin VB.ListBox TitleList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -74760
         TabIndex        =   6
         Top             =   3870
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame BackFrame 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   -72480
         TabIndex        =   20
         Top             =   750
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid TotalsGrid 
         Height          =   4035
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7117
         _Version        =   393216
         Rows            =   100
         Cols            =   4
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
         _Band(0).Cols   =   4
      End
      Begin RichTextLib.RichTextBox vJournal 
         Height          =   4125
         Left            =   -74910
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   7276
         _Version        =   393217
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         TextRTF         =   $"GenWorkForm.frx":1A1DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox TrnInputBox 
         Height          =   3645
         Left            =   -74820
         TabIndex        =   29
         Top             =   600
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   6429
         _Version        =   393217
         TextRTF         =   $"GenWorkForm.frx":1A27A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New Greek"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "GenWorkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectedMenu As Integer
Dim OldShortKey As String
Dim GetL As Variant

Public AppBuffers As New Buffers
Public AppRS As New Collection
Public AppSP As New Collection
Public AppVariables As New Collection
Public AppRS_S As New Collection
Public AppCRS As New Collection
Public L2AppDocs As New cCollection

Dim crypto As New VBPJCrypto

Dim aXMLMapperObj

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
    On Error GoTo errorpos1
    vStatus.Panels(1).Text = sMessage

    Exit Sub
errorpos1:
    Call Runtime_error("Write Status Message", Err.number, Err.description)
End Sub

Public Function fnReadStatusMessage() As String
    On Error GoTo errorpos1
    fnReadStatusMessage = vStatus.Panels(1).Text
    Exit Function
errorpos1:
    Call Runtime_error("Read Status Message", Err.number, Err.description)
End Function

Public Sub sbShowCommStatus(ByVal bActive As Boolean)
    vStatus.Panels(2).Visible = bActive
    vStatus.Panels(3).Visible = Not bActive
    If bActive Then Screen.MousePointer = vbDefault Else Screen.MousePointer = vbArrowHourglass
End Sub

Private Function ReplaceCommandFileVariables(inCmd As String) As String
Dim aPos As Integer, oldLine As String
    
    oldLine = ""
    While oldLine <> inCmd
    
        oldLine = inCmd
        aPos = InStr(inCmd, "%VBONLINESERVER")
        If aPos > 0 Then
            inCmd = Left(inCmd, aPos - 1) & Right(LogonServer, Len(LogonServer) - 2) & _
                    Right(inCmd, Len(inCmd) - aPos - 14)
        End If
        aPos = InStr(inCmd, "%COMPUTERNAME")
        If aPos > 0 Then
            inCmd = Left(inCmd, aPos - 1) & MachineName & _
                    Right(inCmd, Len(inCmd) - aPos - 12)
        End If
        aPos = InStr(inCmd, "%LOCALUSERNAME")
        If aPos > 0 Then
            inCmd = Left(inCmd, aPos - 1) & LOCALUSERNAME & _
                    Right(inCmd, Len(inCmd) - aPos - 13)
        End If
        aPos = InStr(inCmd, "%LOCALUSERPASSWORD")
        If aPos > 0 Then
            inCmd = Left(inCmd, aPos - 1) & LOCALUSERPASSWORD & _
                    Right(inCmd, Len(inCmd) - aPos - 17)
        End If
        
    Wend
    ReplaceCommandFileVariables = inCmd

End Function

Private Function ProcessPublicCommandFile() As Boolean
Dim s As String, sHead As String, sBody As String, sCMD As String, aPos As Integer
    On Error GoTo ErrorPos
    Open LogonDir & "PublicCMD.cfg" For Input As #1
    Do While Not Eof(1)
        Line Input #1, s
        
        aPos = InStr(1, s, "=")
        If aPos > 1 Then
            sHead = Trim(UCase(Left(s, aPos - 1)))
            sBody = Trim(Right(s, Len(s) - aPos))
            If sHead = "BATCHRUN" Then
                sCMD = ReplaceCommandFileVariables(sBody)
                ExecCmd "cmd /c" & sCMD
            ElseIf sHead = "EXERUN" Then
                sCMD = ReplaceCommandFileVariables(sBody)
                ExecCmd sCMD
            End If
        End If
   
    Loop
    Close #1
ErrorPos:

End Function

Private Function InitSpace() As Boolean

cKMODEFlag = False
cKMODEValue = ""
SessID = 1

cVersion = 20040513
cTotalsVersion = 2005
cBRANCHIndex = "0"
cDebug = 0
cPassbookPrinter = 0
cListToPassbook = 1

Dim i As Integer
Dim aMachineName As String * MAX_COMPUTERNAME_LENGTH
Dim res As Integer, astr As String
Dim s As String, aPos As Integer, sHead As String, sBody As String

s = Chr(7) & Chr(148) & Chr(40) & Chr(16) & Chr(84) & Chr(44)
Dim adec As Double
Dim SetManager As String
Dim SetChief As String
Dim SetTeller As String

    adec = HPSToDecimal_(s)
    
    LOCALUSERNAME = "AdminIC"
    LOCALUSERPASSWORD = "-aITIS-1"
    SQLSERVERUSERNAME = "sa"
    SQLSERVERUSERPASSWORD = "sp750"
    
afterWorkData:
    
    Dim commandstr As String
    commandstr = command()
    LocalFlag = False
    If commandstr <> "" And Len(commandstr) >= 5 Then
        If UCase(Left(commandstr, 5)) = "LOCAL" Then
            LocalFlag = True:
            If Len(commandstr) > 5 Then commandstr = Trim(Right(commandstr, Len(commandstr) - 5)) Else commandstr = ""
        End If
    End If
    
    cLogonServer = Trim(Environ("LOGONSERVER"))
    cPDC = GetPrimaryDCName("", "")
    If cLogonServer = "" Then cLogonServer = cPDC
        
    LogonServer = ClearFixedString(cPDC)
    LogonShare = "VBOnline"
    'If Environ("VBOnline_SERVER") <> "" Then LogonServer = Environ("VBOnline_SERVER")
    
    cClientIP = ""
    HasPad = ""
    
    
    Dim commandargs() As String
    commandargs = Split(commandstr, " ")
    If UBound(commandargs) >= 0 Then LogonServer = commandargs(0)
    If UBound(commandargs) >= 1 Then LogonShare = commandargs(1)
    If UBound(commandargs) >= 2 Then cClientName = commandargs(2)
    If UBound(commandargs) >= 3 Then cClientIP = commandargs(3)
    
    If cClientName = "" Then
        res = GetComputerName(aMachineName, MAX_COMPUTERNAME_LENGTH)
        MachineName = ClearFixedString(aMachineName)
    Else
        MachineName = cClientName
    End If
    
    LogonDir = LogonServer & "\" & LogonShare & "\VBLogon\"
    ReadDir = LogonServer & "\" & LogonShare & "\VBRead\"
    WorkDir = LogonServer & "\" & LogonShare & "\NETWORK\"
    If UBound(commandargs) >= 4 Then
        AuthDir = commandargs(4)
    Else
        AuthDir = WorkDir
    End If

    DoEvents
    
    cPRINTERSERVER = MachineName
    cOCRREADERSERVER = ""
    cPrinterPort = 999
    cOCRPort = 999
    
    cUserName = GetUserName
    cUserName = ClearFixedString(cUserName)
    cIRISUserName = cUserName
    
    cSecretToken = ""
    cWebDavPath = "\\v000010354\WebDAV"
    cLocalEncryptedPath = "C:\temp\abcd"

'    cSecretToken = ReadSecretToken(cWebDavPath, cUserName)
'    If cSecretToken = "" Then
'        Dim tokenfilename As String
'        tokenfilename = cLocalEncryptedPath & "\" & cUserName
'        cSecretToken = CreateTokenFile(tokenfilename)
'        If cSecretToken <> "" Then
'            Dim uploadres As Boolean
'            uploadres = WebDavUpload(cWebDavPath, tokenfilename)
'            If Not uploadres Then
'                cSecretToken = ""
'            End If
'        End If
'    End If
    
    cIRISComputerName = MachineName
    cBranchProfileName = "BRANCH"
    cDefUserProfileName = "TELLER"
    cUserProfileName = "TELLER"
    
    astr = CurDir$
    i = AddFontResource(astr & "\" & "bbsecr1.ttf")
    
    xmlEnvironment.appendChild xmlEnvironment.createElement("ROOT")
    UpdatexmlEnvironment "ComputerName", MachineName
    UpdatexmlEnvironment "READDIR", ReadDir
    UpdatexmlEnvironment UCase("IRISUserName"), cIRISUserName
    UpdatexmlEnvironment UCase("IRISComputerName"), cIRISComputerName
    UpdatexmlEnvironment UCase("CHeckTellerBranch"), CStr(False)
    UpdatexmlEnvironment "USERNAME", UCase(cIRISUserName)
    UpdatexmlEnvironment "DISABLESQLSERVER", "1"
    UpdatexmlEnvironment UCase("TOTALSVERSION"), CStr(cTotalsVersion)
    
    On Error GoTo WebLinksError
    prepareWebLinks
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
        If MsgBox("Ζητήσατε να ξεκινήσει εκπαιδευτικό περιβάλλον. Θέλετε να συνεχίσετε;", vbYesNo, "Εφαρμογή Online") = vbNo Then
            InitSpace = False
            Unload Me
            Exit Function
        End If
    End If
    
    On Error GoTo NoCfgError
    Dim station As cWorkstationConfigurationMessage
    Set station = New cWorkstationConfigurationMessage
    Set station = station.Initialize(ReadDir + "\XmlBlocks.xml")
    station.ComputerName = UCase(MachineName)
    station.UserName = UCase(cUserName)

    Dim method As cXMLWebMethod
    Set method = New cXMLWebMethod
    Dim aWeblink As cXMLWebLink
    Set aWeblink = New cXMLWebLink
    aWeblink.VirtualDirectory = TRNFrm.WebLink("OBJECTDISPATCHER_WEBLINK")
    Set method = aWeblink.DefineDocumentMethod("DispatchObject", "http://www.nbg.gr/online/obj")
          
    Dim ares As String
    'ares = Method.LoadXml(station.Message)
    ares = method.LoadXmlNoTrnUpdate(station.Message)
    Dim tempdoc
    Set tempdoc = CreateObject("Msxml2.DOMDocument.6.0")
    tempdoc.LoadXML ares
    Dim returnNode As IXMLDOMElement
    Set returnNode = GetXmlNode(tempdoc.documentElement, "//RESULT/RETURNCODE")
    If returnNode.Text = "0" Then InitSpace = False: Exit Function
    
    Set xmlstation = New cXmlWorkstation
    Set xmlstation = xmlstation.Initialize(tempdoc.documentElement.XML)
    
    Set tempdoc = Nothing
    Set aWeblink = Nothing
    Set method = Nothing
    Set station = Nothing
    
    If Not xmlstation.branch Is Nothing Then
        cBRANCH = xmlstation.branch.Text: UpdatexmlEnvironment "BRANCH", cBRANCH
        UpdatexmlEnvironment "CRABRANCH", cBRANCH
    End If
    If Not xmlstation.BranchName Is Nothing Then cBRANCHName = xmlstation.BranchName.Text: UpdatexmlEnvironment "BRANCHNAME", cBRANCHName
    If Not xmlstation.NewTerminalId Is Nothing Then cTERMINALID = xmlstation.NewTerminalId.Text: UpdatexmlEnvironment "NEWTERMINALID", cTERMINALID
    If Not xmlstation.PassbookPrinter Is Nothing Then cPassbookPrinter = xmlstation.PassbookPrinter.Text: UpdatexmlEnvironment "PASSBOOKPRINTER", CStr(cPassbookPrinter)
    If Not xmlstation.ListToPassbook Is Nothing Then cListToPassbook = xmlstation.ListToPassbook.Text: UpdatexmlEnvironment "LISTTOPASSBOOK", CStr(cListToPassbook)
    If Not xmlstation.PDC Is Nothing Then cPDC = xmlstation.PDC.Text: UpdatexmlEnvironment "PDC", cPDC
    If Not xmlstation.Debugg Is Nothing Then
        cDebug = CInt(Trim(xmlstation.Debugg.Text)): Flag620 = True: Flag630 = True
        UpdatexmlEnvironment "DEBUG", Trim(xmlstation.Debugg.Text)
    End If
    'If Not xmlstation.SmtpServer Is Nothing Then cSMTPServer = xmlstation.SmtpServer.Text: UpdatexmlEnvironment "SMTPSERVER", cSMTPServer
    'If Not xmlstation.SmtpPort Is Nothing Then cSMTPPort = xmlstation.SmtpPort.Text: UpdatexmlEnvironment "SMTPPORT", CStr(cSMTPPort)
    If Not xmlstation.PrinterServer Is Nothing Then cPRINTERSERVER = xmlstation.PrinterServer.Text: UpdatexmlEnvironment "PRINTERSERVER", cPRINTERSERVER
    If Not xmlstation.OcrReaderServer Is Nothing Then cOCRREADERSERVER = xmlstation.OcrReaderServer.Text: UpdatexmlEnvironment "OCRREADERSERVER", cOCRREADERSERVER
    If Not xmlstation.BranchProfileName Is Nothing Then
        If Trim(xmlstation.BranchProfileName.Text) <> "" Then
            cBranchProfileName = xmlstation.BranchProfileName.Text: UpdatexmlEnvironment "BRANCHPROFILENAME", cBranchProfileName
        End If
    End If
    If Not xmlstation.IRISComputerName Is Nothing Then
        If Trim(xmlstation.IRISComputerName.Text) <> "" Then
            cIRISComputerName = xmlstation.IRISComputerName.Text: UpdatexmlEnvironment "IRISCOMPUTERNAME", cIRISComputerName
        End If
    End If
    If Not xmlstation.IRISUSERName Is Nothing Then
        If Trim(xmlstation.IRISUSERName.Text) <> "" Then
            cIRISUserName = xmlstation.IRISUSERName.Text: UpdatexmlEnvironment "IRISUSERNAME", cIRISUserName
            'UpdatexmlEnvironment "FORCEDCRAUSER", cIRISUserName
        End If
    End If
    If Not xmlstation.DebugSNAPoolLink Is Nothing Then
        If xmlstation.DebugSNAPoolLink.Text = "1" Then DebugSNAPoolLink = True: UpdatexmlEnvironment "DEBUGSNAPOOLLINK", "1"
    End If
    If Not xmlstation.ExecRuleB64ReceiveFile Is Nothing Then ExecRuleB64ReceiveFile = xmlstation.ExecRuleB64ReceiveFile.Text: UpdatexmlEnvironment "EXECRULEB64RECEIVEFILE", ExecRuleB64ReceiveFile
    If Not xmlstation.PrinterPort Is Nothing Then cPrinterPort = xmlstation.PrinterPort.Text: UpdatexmlEnvironment "PRINTERPORT", CStr(cPrinterPort)
    If Not xmlstation.OcrPort Is Nothing Then cOCRPort = xmlstation.OcrPort.Text: UpdatexmlEnvironment "OCRPORT", CStr(cOCRPort)
        If Not xmlstation.OpenCobol Is Nothing Then
            If xmlstation.OpenCobol.Text = "1" Then OpenCobolServer = True: UpdatexmlEnvironment "OPENCOBOL", "1"
            If xmlstation.OpenCobol.Text = "0" Then OpenCobolServer = False: UpdatexmlEnvironment "OPENCOBOL", "0"
        End If
    If Not xmlstation.UseActiveDirectory Is Nothing Then
        If xmlstation.UseActiveDirectory.Text = "1" Then UseActiveDirectory = True: UpdatexmlEnvironment "USEACTIVEDIRECTORY", "1"
        If xmlstation.UseActiveDirectory.Text = "0" Then UseActiveDirectory = False: UpdatexmlEnvironment "USEACTIVEDIRECTORY", "0"
    End If
    If Not xmlstation.SetManager Is Nothing Then SetManager = xmlstation.SetManager.Text: UpdatexmlEnvironment "SETMANAGER", SetManager
    If Not xmlstation.SetChief Is Nothing Then SetChief = xmlstation.SetChief.Text: UpdatexmlEnvironment "SETCHIEF", SetChief
    If Not xmlstation.SetTeller Is Nothing Then SetTeller = xmlstation.SetTeller.Text: UpdatexmlEnvironment "SETTELLER", SetTeller
    If Not xmlstation.JournalType Is Nothing Then
        If xmlstation.JournalType.Text = "1" Then
            cNewJournalType = True
            UpdatexmlEnvironment "JOURNALTYPE", "1"
        Else
            cNewJournalType = False
            UpdatexmlEnvironment "JOURNALTYPE", "0"
        End If
    End If
    If Not xmlstation.LastTrnNumber Is Nothing Then
        cTRNNum = CInt(xmlstation.LastTrnNumber.Text)
        UpdatexmlEnvironment "TRNNUM", CStr(cTRNNum)
    End If
    If Not xmlstation.CicsUserInfo Is Nothing Then
        If xmlstation.CicsUserInfo.Text = "1" Then
            cUseCicsUserInfo = True
        Else
            cUseCicsUserInfo = False
        End If
    Else
        cUseCicsUserInfo = False
    End If
    If Not xmlstation.HasWinPanel Is Nothing Then
        If xmlstation.HasWinPanel.Text = "1" Then
            cHasWinPanel = True
        Else
            cHasWinPanel = False
        End If
    Else
        cHasWinPanel = False
    End If
    If Not xmlstation.TellerTrn Is Nothing Then
        If xmlstation.TellerTrn.Text = "1" Then
            cHasTellerTrnGroup = True
        Else
            cHasTellerTrnGroup = False
        End If
    Else
        cHasTellerTrnGroup = False
    End If
    If Not xmlstation.HasPad Is Nothing Then
        If xmlstation.HasPad.Text = "1" Then
            HasPad = "1"
            UpdatexmlEnvironment "HASPAD", HasPad
        End If
    End If
    
    UpdatexmlEnvironment UCase("IRISWORKSTATIONNAME"), Right(String(8, " ") & cIRISComputerName, 8)
    
    If cDebug = 1 Then RebuildViewBtn.Visible = True Else RebuildViewBtn.Visible = False
    If cDebug = 1 Then RebuildComareaBtn.Visible = True Else RebuildComareaBtn.Visible = False
    
    ProcessPublicCommandFile
   
    On Error GoTo InvalidChkUser:
    
    isTeller = False
    isChiefTeller = False
    isManager = False
    ChkUser
    If UseActiveDirectory Then
        Dim adTool As New cADTool
        adTool.Initialize
        UserGroups.add "test"
        While UserGroups.count > 0
            UserGroups.Remove UserGroups.count
        Wend
        Dim GroupName
        For Each GroupName In adTool.UserGroups
            If UCase(GroupName) = "TELLER" Then isTeller = True
            If UCase(GroupName) = "CHIEF TELLER" Then isChiefTeller = True
            If UCase(GroupName) = "MANAGER" Then isManager = True
            If UCase(GroupName) = "IMPORT USERS" Then isImportUser = True
            UserGroups.add UCase(GroupName)
        Next GroupName
        
        UpdatexmlEnvironment "TELLER", CStr(isTeller)
        UpdatexmlEnvironment "CHIEFTELLER", CStr(isChiefTeller)
        UpdatexmlEnvironment "MANAGER", CStr(isManager)
        UpdatexmlEnvironment "IMPORTUSER", CStr(isImportUser)
    End If
    
    If cUseCicsUserInfo Then
        isTeller = False
        isChiefTeller = False
        isManager = False
        If Not xmlstation.IsHostTeller Is Nothing Then
            If xmlstation.IsHostTeller.Text = "1" Then
                isTeller = True
            End If
        End If
        If Not xmlstation.IsHostChief Is Nothing Then
            If xmlstation.IsHostChief.Text = "1" Then
                isChiefTeller = True
            End If
        End If
        If Not xmlstation.IsHostManager Is Nothing Then
            If xmlstation.IsHostManager.Text = "1" Then
                isManager = True
            End If
        End If
    End If
    
    If SetManager <> "" Then
        If SetManager = "1" Then
            isManager = True
        ElseIf SetManager = "0" Then
            isManager = False
        End If
        UpdatexmlEnvironment "MANAGER", CStr(isManager)
    End If
    If SetChief <> "" Then
        If SetChief = "1" Then
            isChiefTeller = True
        ElseIf SetChief = "0" Then
            isChiefTeller = False
        End If
        UpdatexmlEnvironment "CHIEFTELLER", CStr(isChiefTeller)
    End If
    If SetTeller <> "" Then
        If SetTeller = "1" Then
            isTeller = True
        ElseIf SetTeller = "0" Then
            isTeller = False
        End If
        UpdatexmlEnvironment "TELLER", CStr(isTeller)
    End If
    
    On Error GoTo InvalidL2TrnListFile:
    Set L2TrnListFile = New MSXML2.DOMDocument30
    L2TrnListFile.Load ReadDir & "\" & "L2TrnList.xml"
    On Error Resume Next
    Set L2AddInFile = New MSXML2.DOMDocument30
    L2AddInFile.Load ReadDir & "\" & "AddIns.xml"
    
    On Error GoTo InvalidDBInit:
    
    Call PrepareEnv
    
    On Error GoTo InvalidXMLInit:
    
    Call PrepareXML
    On Error GoTo InvalidProfile:
    
    If Not PrepareNewProfiles Then
        InitSpace = False: Exit Function
    End If
    
    If cHasWinPanel Then
        Call StartNQCashierAndLogin(Me.hwnd, cUserName, cFullUserName)
    End If
    
'    If cHasWinPanel Then
'        On Error Resume Next
'        Dim winpanelinifilename As String
'        winpanelinifilename = "c:\nq2000\nqwinpanel\Cashier2000.ini"
'        If fnChkFileExistAbs(winpanelinifilename) Then
'            Dim winpanelserver As String
'            Dim iStr As String
'            iStr = String(255, Chr(0))
'            winpanelserver = Left(iStr, GetPrivateProfileString("SYSTEM", ByVal "ServerAddress", "", iStr, Len(iStr), winpanelinifilename))
'            winpanelserver = Trim(winpanelserver)
'            If winpanelserver = "" Or winpanelserver = "?" Then
'                cHasWinPanel = False
'            Else
'                Call StartNQCashierAndLogin(Me.hwnd, cUserName, cFullUserName)
'            End If
'        Else
'            cHasWinPanel = False
'        End If
'    End If
    
    LastTRNCode = cTRNCode
    LastTRNNum = cTRNNum
    EventLogWrite = False
    SendJournalWrite = False
    ReceiveJournalWrite = False
    SRJournal = False
    ASCII_CP_STRING = ""
    ' ο χαρακτηρας είναι 255 έχει γραφεί με το alt και 255 από το αριθμητικό πληκτρολογιο
    For i = 1 To 31
        ASCII_CP_STRING = ASCII_CP_STRING & " "
    Next
    ASCII_CP_STRING = ASCII_CP_STRING & " !" & Chr$(34) & "#$%&'()*+,-./0123456789:;<=>?" & _
    "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €" & _
    "‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©ª«¬­®―°±²³΄µ¶·ΈΉΊ»Ό½ΎΏΐΑ" & _
    "ΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡÒΣΤΥΦΧΨΩΪΫάέήίΰαβγδεζηθικλμνξοπρςστυφχψωϊϋόύώÿ" & Chr$(0)
    EBCDIC_CP_STRING = ""
    ' ο χαρακτηρας είναι 255 έχει γραφεί με το alt και 255 από το αριθμητικό πληκτρολογιο
    For i = 1 To 63
        EBCDIC_CP_STRING = EBCDIC_CP_STRING & " "
    Next
    EBCDIC_CP_STRING = EBCDIC_CP_STRING & " ΑΒΓΔΕΖΗΘΙ[.<(+!&ΚΛΜΝΞΟΠΡΣ]$*);^-/ΤΥΦΧΨΩΪΫ|,%_>?" & _
    "¨ΆΈΉ ΊΌΎΏ`:#@'=" & Chr$(34) & "΅abcdefghiαβγδεζ°jklmnop" & _
    "qrηθικλμ`~stuvwxyzνξοπρσ£άέήϊίόύϋώςτυφχψ{ABCDEFGHI­ωΐΰ‘-}JKLMNOPQ" & _
    "R±½ ·’|\ STUVWXYZ²§  «¬0123456789³©  » " & Chr$(0)
    
    PASSBOOK_CLEAR_STRING = " !" & Chr$(34) & "#$%&'()*+,-./0123456789:;<=>?" & _
    "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_{|}~‚ΈΉΊΌΎΏΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡÒΣΤΥΦΧΨΩΪΫ"
    
    
    
    cb.CodePage = UCS_OLD
    cb.com_debug = 1  '0 : no debug
                      '1 : write to Event Log
                      '2 : write to Event Log AND show to Screen
    cb.app_debug = 1  '0 : no debug
                      '1 : write to Event Log
                      '2 : write to Event Log AND show to Screen
    
    cb.com_debug = 0
    
    On Error GoTo InvalidTotalInit:
    initialize_cb
    
    GetL = GetKeyboardLayout(0)
    If Right(CStr(Hex(GetL)), 2) <> "08" Then ActivateKeyboardLayout 0, 0
    
    LoadJournal
    
step3:

'Flag610 = True:
'Flag620 = True:
'Flag630 = True
On Error Resume Next
Call SetOnLineLocaleInfo
InitSpace = True

Exit Function

InvalidLogonPath:
    NBG_LOG_MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α1) " & error() & " " & LogonDir & "server.cfg", True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidLogonPath2:
    NBG_LOG_MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α1B) " & error() & " " & LogonDir & MachineName & ".cfg", True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidL2TrnListFile:
    NBG_LOG_MsgBox "Λάθος στο άνοιγμα Αρχείου Συναλλαγών L2... (Α1C) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidDBInit:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Βάσης Δεδομένων... (Α2) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidXMLInit:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Παραμέτρων Συναλλαγών... (Α3) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidProfile:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Παραμέτρων Χρήστη... (Α3Β) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidJournalInit:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Ημερολογίου... (Α4) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidTotalInit:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Αθροιστών... (Α5) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidChkUser:
    NBG_LOG_MsgBox "Λάθος στην Ταυτοποίηση Χρήστη... (Α6) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
InvalidPrinterInit:
    NBG_LOG_MsgBox "Λάθος στην Έναρξη Λειτουργίας Εκτυπωτή... (Α7) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
WebLinksError:
    NBG_LOG_MsgBox "Λάθος κατά το διάβασμα του αρχείου WebLinks " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function
NoCfgError:
    NBG_LOG_MsgBox "Λάθος κατά την Εναρξη Λειτουργίας Τερματικού (Α8) " & error(), True, "ΛΑΘΟΣ"
    InitSpace = False
    Exit Function

End Function

Private Sub SetKeys()
    If isTeller Then
        KeyBar.Buttons(1).Image = 2
        KeyBar.Buttons(1).value = tbrPressed
    Else
        KeyBar.Buttons(1).Image = 1
        KeyBar.Buttons(1).value = tbrUnpressed
    End If
    If isChiefTeller Then
        KeyBar.Buttons(2).Image = 2
        KeyBar.Buttons(2).value = tbrPressed
    Else
        KeyBar.Buttons(2).Image = 1
        KeyBar.Buttons(2).value = tbrUnpressed
    End If
    
    If isManager Then
        KeyBar.Buttons(3).Image = 2
        KeyBar.Buttons(3).value = tbrPressed
    Else
        KeyBar.Buttons(3).Image = 1
        KeyBar.Buttons(3).value = tbrUnpressed
    End If
End Sub

Private Sub SetSelectedMenu(Index As Integer)
Dim StartTop As Integer, i As Integer, k As Integer
Dim oldSelectedMenu As Integer
    
    SSTabControl.Tab = 0
    oldSelectedMenu = SelectedMenu
    SelectedMenu = Index
    BackFrame.Top = 330: BackFrame.Left = 30
    BackFrame.width = SSTabControl.width - 60
    BackFrame.height = SSTabControl.height - 360
    TitleList.Left = 50
    TitleList.width = SSTabControl.width - 100
    StartTop = 360
    
    For i = 0 To Index
        MenuCommand(i).Left = 50
        MenuCommand(i).Top = StartTop
        MenuCommand(i).width = SSTabControl.width - 100
        StartTop = StartTop + 375
    Next i
    TitleList.Top = StartTop + 30
    StartTop = StartTop + TitleList.height
    StartTop = SSTabControl.height
    For i = 9 To Index + 1 Step -1
        StartTop = StartTop - 375
        MenuCommand(i).Left = 50
        MenuCommand(i).Top = StartTop
        MenuCommand(i).width = SSTabControl.width - 100
    Next i
    If StartTop - TitleList.Top > 0 Then
        TitleList.height = StartTop - TitleList.Top - 60
    End If
    TitleList.Visible = True
    If oldSelectedMenu <> SelectedMenu Then
        TitleList.Clear

Dim anode As Variant, bNode As Variant, astr As String
Dim HiddenFlag As Boolean
        Set anode = xmlNewMenu.documentElement.selectSingleNode("MenuItem[@CD='" & Trim(CStr(Index)) & "']")
        If Not (anode Is Nothing) Then
            For Each bNode In anode.childNodes
                HiddenFlag = False
                If Not (bNode.getAttributeNode("hidden") Is Nothing) Then
                    If Trim(bNode.getAttributeNode("hidden").nodeValue) = "1" Then
                        HiddenFlag = True: Exit For
                    End If
                End If
                If Not HiddenFlag Then
                    astr = bNode.getAttributeNode("id").nodeValue
                    astr = StrPad_(astr, 4, "0", "L")
                    TitleList.AddItem (astr & " - " & bNode.getAttributeNode("name").nodeValue)
                    TitleList.ItemData(TitleList.NewIndex) = CInt(astr)
                End If
            Next
        End If

'        Set anode = xmlMenu.documentElement.selectSingleNode("M" & Trim(CStr(Index + 1)))
'        If Not (anode Is Nothing) Then
'            For i = 1 To anode.childNodes.length - 1
'                Set bNode = anode.childNodes.item(i)
'                HiddenFlag = False
'                If Not (bNode Is Nothing) Then
'                    If bNode.childNodes.length > 0 Then
'                        For k = 1 To bNode.childNodes.length - 1
'                            If UCase(bNode.childNodes.item(k).tagName) = "HIDDEN" Then
'                                HiddenFlag = True: Exit For
'                            End If
'                        Next k
'                    End If
'                End If
'                If Not HiddenFlag Then
'                    astr = anode.childNodes.item(i).tagName
'                    astr = StrPad_(Right(astr, Len(astr) - 1), 4, "0", "L")
'                    TitleList.AddItem (astr & " - " & anode.childNodes.item(i).Text)
'                    TitleList.ItemData(TitleList.NewIndex) = CInt(astr)
'                End If
'            Next i
'        End If
        
        If SelectedMenu < 9 Then
            For i = 9 To SelectedMenu + 1 Step -1
                MenuCommand(i).TabIndex = i + 3
            Next i
        End If
        TitleList.TabIndex = SelectedMenu + 3
        For i = SelectedMenu To 0 Step -1
            MenuCommand(i).TabIndex = i + 2
        Next i
    End If
    
End Sub


Private Sub Command3_Click()
'    Dim aapp As cApplication
'
'    Set aapp = New cApplication
'    Dim ComArea
'  Set ComArea = aapp.DeclareComArea("myS4A00", "S4A00", "S4A00", "4A00", "@S4A00", "IDATA", "")
'  With ComArea.BufferByName("S4A00")
'        .ClearData
'        .v2Value("PROD", 1) = "51"
'        .v2Value("APP_NO") = 1234570
'        .v2Value("AMOUNT") = 900000
'        .v2Value("I_IP") = "6456766958"
'        .v2Value("C_BRANCH") = "076"
'        .v2Value("C_ACCOUNT") = "1001000"
'        .v2Value("P_BRANCH") = "053"
'        .v2Value("P_ACCOUNT") = "8036861"
'        .v2Value("I_REV") = ""
'   End With
'   Dim xmltostring As String
'   Dim res
'      xmltostring = aapp.ParseComArea(ComArea)
'  Set res = ComArea.BufferByName("S4A00").GetXMLView

    
End Sub

Private Sub ComTimer_Timer()
Dim comstatus As Integer
    ComTimerCounter = ComTimerCounter + 1
    Screen.MousePointer = vbArrowHourglass
    If (ComTimerCycle >= 5 And ComTimerCounter = 50) _
    Or (ComTimerCycle < 5 And ComTimerCounter = 10) Then
        ComTimerCycle = ComTimerCycle + 1
        ComTimerCounter = 0
        If Restore_Connection = 1 Then
            ComTimer.Enabled = False
            sbShowCommStatus (True)
            MsgBox "Η Επικοινωνία αποκαταστάθηκε.", vbOKOnly, "On Line Εφαρμογή"
            Screen.MousePointer = vbDefault
        End If
            
    End If
End Sub

Private Sub EventLogWriteChk_Click()
    EventLogWrite = EventLogWriteChk.value
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    shortkey.SetFocus
End Sub

Private Sub Form_Initialize()

    shortkey.width = SSTabControl.width
    SetSelectedMenu (SelectedMenu)
    
'    vStatus.Panels(1).Width = Round(vStatus.Width * 0.9)
    
    CommandFrame.Left = 0: CommandFrame.Top = -60
    CommandFrame.height = vStatus.Top - CommandFrame.Top - 30
    
    ShortLabel.Left = 30
    shortkey.Top = ShortLabel.Top + ShortLabel.height + 30
    shortkey.Left = 30
    shortkey.width = CommandFrame.width - 120
    
    TitleFrame.Top = -60
    TitleFrame.Left = CommandFrame.width + 30
    TitleFrame.width = width - TitleFrame.Left - 100
    
    KeyBar.Left = 30: KeyBar.Top = CommandFrame.height - KeyBar.height - 60
    KeyBar.width = CommandFrame.width - 60
    
    SSTabControl.Left = TitleFrame.Left
    SSTabControl.width = TitleFrame.width
    SSTabControl.Top = TitleFrame.height + TitleFrame.Top + 10
    SSTabControl.height = vStatus.Top - SSTabControl.Top - 30
    
    If SSTabControl.TabVisible(4) Then
        SSTabControl.Tab = 4
        TrnInputBox.Left = 30: TrnInputBox.Top = 330
        TrnInputBox.width = SSTabControl.width - 60
        TrnInputBox.height = SSTabControl.height - 360
        
        SSTabControl.TabVisible(4) = (cDebug > 0)
    End If
    
    SSTabControl.Tab = 3
    StationInfo.width = SSTabControl.width - StationInfo.Left - 80
    
    SSTabControl.Tab = 2
    vJournal.Left = 30: vJournal.Top = 330
    vJournal.width = SSTabControl.width - 60
    vJournal.height = SSTabControl.height - 360
    'If cNewJournalType = True Then SSTabControl.TabVisible(2) = False
    
    SSTabControl.Tab = 1
    TotalsGrid.Left = 30: TotalsGrid.Top = 330
    TotalsGrid.width = SSTabControl.width - 60
    TotalsGrid.height = SSTabControl.height - 360
    SSTabControl.TabVisible(1) = False
    
    SSTabControl.Tab = 0
    If SelectedMenu <= 0 Then SetSelectedMenu (0)
    
    If Left(Right(WorkEnvironment_, 8), 4) = "EDUC" Then
        Caption = Caption & " (ΕΚΠΑΙΔΕΥΤΙΚΟ ΠΕΡΙΒΑΛΛΟΝ)"
    Else
        Caption = Caption & " (ΠΕΡΙΒΑΛΛΟΝ ΠΑΡΑΓΩΓΗΣ)"
    End If
    
    Dim aDesc As String
    If Not SkipCRAUse Then
        On Error GoTo CRAStructures_Error
        aDesc = xmlCRAStructures.documentElement.selectSingleNode("VCUUP01").Text

    
'        AppBuffers.DefineBuffer "VCUUP01", aDesc, , True:
        AppBuffers.name = "AppBuffers"
        
        BuildCRAAppStruct "VCUUP01", "VCUUP01", True
        With AppBuffers.ByName("VCUUP01")
            .ByName("I_ENTP").value = 2 '1 σωστο αλλά 2 για να χτυπησει Αλλαγή 13-10-2005
            .ByName("I_USR_FI").value = 2
            .ByName("C_ACOD_FI").value = "001"
            .ByName("I_USR_OU").value = 741
            .ByName("C_ACOD_OU").value = Right("000" & cBRANCH, 3)
            .ByName("D_PROC").value = Date
            If cIRISUserName <> "" And cDebug = 1 Then
                .ByName("C_USR_ID").value = cIRISUserName
            Else
                .ByName("C_USR_ID").value = UCase(cUserName)
            End If
            
            .ByName("C_WKST_ID").value = "307"
            .ByName("I_LOC_DFLT_PRFL").value = 1
            .ByName("C_GEO_PRFL").value = "GR"
            .ByName("C_PREF_LANG_TP_PRFL").value = "GRK"
            .ByName("I_CLSF_EXCH_MEDM_K_PRFL").value = 1100001
            .ByName("C_CLSF_EXCH_MEDM_K_PRFL").value = "GRD"
            .ByName("I_CLSF_PRTFL_K_PRFL").value = 500001
            .ByName("C_CLSF_PRTFL_K_PRFL").value = "ΟΛΟΙ"
            .ByName("C_SRCH_PRFL").value = "EN"
        End With
        
        AppBuffers.DefineBuffer "NAMEFORMATLINE", "NAMEFORMATLINE", "Lbl1 char 40, Lbl2 char 1, Flag small", "NAMEFORMATLINE", False
        AppBuffers.DefineBuffer "NAMEFORMAT", "NAMEFORMAT", "NAMEFORMATLINE struct NAMEFORMATLINE 12", "NAMEFORMAT", True
        BuildCRAAppStruct "ZAFNDLE", "ZAFNDLE", True
        BuildCRAAppStruct "ZAFNELE", "ZAFNELE", True
        
        
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("VCUER04").Text
'        AppBuffers.DefineBuffer "VCUER04", aDesc
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("VCUER01").Text
'        AppBuffers.DefineBuffer "VCUER01", aDesc
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("ZAFNDLE").Text
'        AppBuffers.DefineBuffer "ZAFNDLE", aDesc
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("ZAFNFLE").Text
'        AppBuffers.DefineBuffer "ZAFNFLE", aDesc
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("ZAFNELE").Text
'        AppBuffers.DefineBuffer "ZAFNELE", aDesc
        
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("ZAAC9WJ").Text
'        AppBuffers.DefineBuffer "ZAAC9WJ", aDesc, "ZAAC9WJ", True
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("RCUUT60A").Text
'        AppBuffers.DefineBuffer "RCUUT60A", aDesc, "RCUUT60A", True
'        aDesc = xmlCRAStructures.documentElement.selectSingleNode("ZAADAWJ").Text
'        AppBuffers.DefineBuffer "ZAADAWJ_C", aDesc, "ZAADAWJ", True
        
        BuildCRAAppStruct "ZAAC9WJ", "ZAAC9WJ", True
        BuildCRAAppStruct "ZAADAWJ", "ZAADAWJ_C", True
        
    End If
        On Error GoTo IRISStructures_Error
        BuildIRISAppStruct "TR_APERTURA_PUESTO_TRN_I", "TR_APERTURA_PUESTO_TRN_I", True
        BuildIRISAppStruct "TR_APERTURA_PUESTO_TRN_O", "TR_APERTURA_PUESTO_TRN_O", True
        BuildIRISAppStruct "TR_CONS_CENTRO_TRN_I", "TR_CONS_CENTRO_TRN_I", True
        BuildIRISAppStruct "TR_CONS_CENTRO_TRN_O", "TR_CONS_CENTRO_TRN_O", True
        
        Dim anode As MSXML2.IXMLDOMElement
        If Not (xmlIRISRules.documentElement Is Nothing) Then
            Set anode = xmlIRISRules.documentElement.selectSingleNode("A54MDV")
            If Not (anode Is Nothing) Then
                BuildIRISAppStruct "TR_CONNECT_IRIS_ICL_TRN_I", "TR_CONNECT_IRIS_ICL_TRN_I", True
                BuildIRISAppStruct "TR_CONNECT_IRIS_ICL_TRN_O", "TR_CONNECT_IRIS_ICL_TRN_O", True
            End If
        End If
        
        
        
'        BuildIRISAppStruct "ZAJ8DM3", "ZAJ8DM3", True
        'BuildIRISAppStruct "ZAJ8EM3", "ZAJ8EM3", True
        'BuildIRISAppStruct "ZANTZOY", "ZANTZOY", True
        'BuildIRISAppStruct "ZANTYOY", "ZANTYOY", True
        
        
'ZAB74PN struct ZAB74PN 1, ZAJ8JM3 struct ZAJ8JM3 1
    'End If
    Exit Sub
CRAStructures_Error:
    NBG_LOG_MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α10 CRA) " & error() & " " & LogonDir & "server.cfg", True, "ΛΑΘΟΣ"
IRISStructures_Error:
    NBG_LOG_MsgBox "Λάθος Παράμετροι Έναρξης της Εφαρμογής... (Α10 IRIS) " & error() & " " & LogonDir & "server.cfg", True, "ΛΑΘΟΣ"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        KeyCode = 0
        On Error GoTo 0
        If cHasWinPanel And HasWinPanelConnection Then
            NQCashierTicketID = NQCashierCallNextCustomer()
        End If
    ElseIf KeyCode = vbKeyF4 And ((Shift And vbAltMask) = 0) Then
        KeyCode = 0
        On Error GoTo 0
    ElseIf KeyCode = vbKeyF5 Then
        KeyCode = 0
        cTRNCode = 2100: OldShortKey = "": shortkey.Text = ""
        On Error GoTo 0: OpenTrnFrm
    ElseIf KeyCode = vbKeyF6 Then
        KeyCode = 0
        cTRNCode = 2000: OldShortKey = "": shortkey.Text = ""
        On Error GoTo 0: OpenTrnFrm
    ElseIf KeyCode = vbKeyF7 Then
       SSTabControl.Tab = IIf(SSTabControl.Tab = 0, SSTabControl.Tabs - 2, SSTabControl.Tab - 1)
    ElseIf KeyCode = vbKeyF8 Then
       SSTabControl.Tab = IIf(SSTabControl.Tab = SSTabControl.Tabs - 2, 0, SSTabControl.Tab + 1)
    ElseIf KeyCode = vbKeyF10 Then
        KeyCode = 0
'        If cNewJournalType = False Then
'            eJournalFrm.Show vbModal, Me
'        Else
            Dim aTRNHandler As New L2TrnHandler
            aTRNHandler.ExecuteForm "9989"
            aTRNHandler.CleanUp
            Set aTRNHandler = Nothing
'        End If
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        Dim bTRNHandler As New L2TrnHandler
        bTRNHandler.ExecuteForm "9747"
        bTRNHandler.CleanUp
        Set bTRNHandler = Nothing
    ElseIf KeyCode = 83 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-s
        MailSlotFrm.Show vbModal, Me
    ElseIf KeyCode = 65 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-a
        KeyCode = 0
        Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = AppBuffers
        BufferViewer.Show vbModal, Me
        Unload BufferViewer
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then shortkey.SetFocus
End Sub
Public Sub UpdateStationInfo()
    StationInfo.Clear
    StationInfo.AddItem "Version            : " & cVersion
    StationInfo.AddItem "Σταθμος            : " & MachineName
    StationInfo.AddItem "Τερματικό          : " & cTERMINALID
    StationInfo.AddItem "Κατάστημα          : " & cBRANCH & " " & cBRANCHName
    StationInfo.AddItem "Ημερομηνία         : " & cPOSTDATE
    StationInfo.AddItem "Χρήστης            : " & cUserName & " " & cFullUserName
    StationInfo.AddItem "Server             : " & LogonServer & " " & LogonDir
    StationInfo.AddItem "Συναλλαγές         : " & ReadDir
    StationInfo.AddItem "Εγκρίσεις          : " & AuthDir
    'StationInfo.AddItem "Printer            : " & AuthDir
    StationInfo.AddItem "Περιβάλλον         : " & WorkEnvironment_
    StationInfo.AddItem "Χρήση Κεντρικού SNA: " & "Ναι"
    StationInfo.AddItem "Disable SQL Server : " & "Ναι"
    StationInfo.AddItem "Active Directory   : " & IIf(UseActiveDirectory, "Ναι", "Οχι")
    
    
End Sub

Private Sub Form_Load()
Dim res As Boolean
    If App.PrevInstance <> 0 Then GoTo ApplicationRunning

    If Not InitSpace Then GoTo FatalError
    
    EventLogWrite = False: SendJournalWrite = True: ReceiveJournalWrite = True
    
    If EventLogWrite Then EventLogWriteChk.value = vbChecked Else EventLogWriteChk.value = vbUnchecked
    If SendJournalWrite Then SendJournalWriteChk.value = vbChecked Else SendJournalWriteChk.value = vbUnchecked
    If ReceiveJournalWrite Then ReceiveJournalWriteChk.value = vbChecked Else ReceiveJournalWriteChk.value = vbUnchecked
    If SRJournal Then SRJournalChk.value = vbChecked Else SRJournalChk.value = vbUnchecked
    
'    vStatus.Panels(1).Width = Round(vStatus.Width * 0.9)
    Dim i As Integer
    i = 0
    Dim anode As IXMLDOMNode
    For Each anode In xmlNewMenu.documentElement.childNodes
        MenuCommand(i).Caption = anode.Attributes.getNamedItem("name").Text
        'MenuCommand(i).Caption = anode.Attributes.getNamedItem("CD").Text & " - " & _
        '    anode.Attributes.getNamedItem("name").Text
        i = i + 1
    Next
    
    If Not fnDisplayTotals(TotalsGrid) Then GoTo FatalError
    
    If format(cPOSTDATE, "dd/mm/yyyy") <> format(Date, "dd/mm/yyyy") Then cPOSTDATE = Date: cTRNNum = 0: UpdateParams
    SelectedMenu = -1
    Left = 0: width = 12000: Top = 0: height = 8650
    
    SetKeys
    UpdateStationInfo

Exit Sub
ApplicationRunning:
    NBG_MsgBox "Η εφαρμογή βρίσκεται ήδη σε λειτουργία....", True, "ΛΑΘΟΣ"
FatalError:
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorQuit
    
    If cHasWinPanel And HasWinPanelConnection Then
        StopNQCashierAndLogout
    End If
    
    Dim pID As Long
    'KillProcess GetCurrentProcess(), 0
    KillProcesses ("shine.exe")
    Exit Sub
ErrorQuit: End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'    Close_Response_Pipe
Dim Status As Long, aFlag As Boolean
    Unload TRNFrm
    SaveJournal
    ActivateKeyboardLayout GetL, 1
End Sub

Private Sub MenuCommand_Click(Index As Integer)
    shortkey.Text = ""
    SetSelectedMenu (Index)
End Sub

Private Sub MenuCommand_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then _
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then shortkey.SetFocus
End Sub

Private Sub PrnTestCmd_Click()
    Dim astr As String, i As Integer, alldata As String, aFlag As Boolean
    astr = "123456789_123456789_123456789_123456789_123456789_123456789_123456789_"
    
    Dim DocLines(54) As String
    For i = 0 To 54
        DocLines(i) = astr & CStr(i + 1)
    Next
    If cPassbookPrinter = 9 Then
    ElseIf cPassbookPrinter <> 0 Then
        aFlag = (cPassbookPrinter = 3 Or cPassbookPrinter = 4)
        If aFlag Then
        Else
            alldata = ""
            For i = 0 To 54
                alldata = alldata & Trim("Line" & StrPad_(CStr(i + 1), 3, "0", "L") & "=" & DocLines(i)) & "|"
            Next i
            alldata = alldata & Chr(0)
            
        End If
    End If
    
    
End Sub

Private Sub RebuildComareaBtn_Click()
    Dim aviewname As String
    Dim aviewfile As String
    aviewname = InputBox("Ονομα ComArea:", "Ορισμός View")
    aviewfile = InputBox("Ονομα Αρχείου:", "Αρχείο Δεδομένων")
    
    Dim area As cXmlComArea
    Set area = DeclareComArea_("S" & aviewname, "S" & aviewname, "S" & aviewname, aviewname, "@S" & aviewname, "IDATA", "ODATA")
    If aviewfile <> "" Then
        Dim aStruct As Buffer
        Dim datastr As String, DataPart
        Set aStruct = AppBuffers.ByName("S" & aviewname)
        Open aviewfile For Binary As #3
        Do While (Loc(3) < LOF(3))
            DataPart = input(1, #3)
            datastr = datastr & DataPart
        Loop
        Close #3
        'datastr = Mid(datastr, IRIS_OFFSET + 1, Len(datastr) - IRIS_OFFSET)
        aStruct.Data = datastr
    End If
End Sub

Private Sub RebuildViewBtn_Click()
'    Dim aBuffer As Buffers
'    Dim i As Integer
'    For i = 1 To 20
'        Set aBuffer = New Buffers
'        BuildIRISStruct aBuffer, "CRA_CUST_SRCH_BRO_TRN_I", "CRA_CUST_SRCH_BRO_TRN_I"
'        BuildIRISStruct aBuffer, "CRA_CUST_SRCH_BRO_TRN_O", "CRA_CUST_SRCH_BRO_TRN_O"
'        BuildIRISStruct aBuffer, "TR_PRESENTACION_AC_TRN_I", "TR_PRESENTACION_AC_TRN_I"
'        BuildIRISStruct aBuffer, "TR_PRESENTACION_AC_TRN_O", "TR_PRESENTACION_AC_TRN_O"
'        BuildIRISStruct aBuffer, "TR_CONS_GLOBAL_H_KP_TRM_TRN_I", "TR_CONS_GLOBAL_H_KP_TRM_TRN_I"
'        BuildIRISStruct aBuffer, "TR_CONS_GLOBAL_H_KP_TRM_TRN_O", "TR_CONS_GLOBAL_H_KP_TRM_TRN_O"
'        'abuffer.DefineComArea structurecode, StructureName, False
'
'        BuildComArea aBuffer, "S1000", "@S1000"
'        aBuffer.ClearAll
'        Set aBuffer = Nothing
'    Next i
'    Exit Sub
    
    Dim aviewname As String
    Dim aviewfile As String
    aviewname = InputBox("Ονομα View:", "Ορισμός View")
    aviewfile = InputBox("Ονομα Αρχείου:", "Αρχείο Δεδομένων")
    
    If Not AppBuffers.Exists(aviewname) Then
        BuildIRISAppStruct aviewname, aviewname, True
    End If
    If aviewfile <> "" Then
    Dim aStruct As Buffer
    Dim datastr As String, DataPart
    Set aStruct = AppBuffers.ByName(aviewname)
    Open aviewfile For Binary As #3
    Do While (Loc(3) < LOF(3))
        DataPart = input(1, #3)
        datastr = datastr & DataPart
    Loop
    Close #3
    
    datastr = Mid(datastr, IRIS_OFFSET + 1, Len(datastr) - IRIS_OFFSET)
    aStruct.Data = datastr
    End If
End Sub

Private Sub ReceiveJournalWriteChk_Click()
    ReceiveJournalWrite = ReceiveJournalWriteChk.value
End Sub

Private Sub SendJournalWriteChk_Click()
    SendJournalWrite = SendJournalWriteChk.value
End Sub

Private Sub shortkey_Change()
Dim avar As String
Dim astr As String
    
    astr = shortkey.Text
    If Len(Trim(astr)) > 0 Then
'        On Error Resume Next
        On Error GoTo clearShortkey
        
        avar = Str(CInt(astr))
'        On Error GoTo clearShortkey
        
        If Len(shortkey.Text) = 1 Then
            'SetSelectedMenu (Round(Val(astr)))
        ElseIf Len(astr) = 4 And avar <> "" Then
            If avar = 610 Or avar = 611 Then
                DoEvents: cTRNCode = CInt(astr): OldShortKey = "": shortkey.Text = ""
                On Error GoTo 0
                T0611New.Show vbModal, Me
            Else
                cTRNCode = CInt(astr)
                OldShortKey = ""
                shortkey.Text = ""
                cEnableHiddenTransactions = False
                On Error GoTo 0
                OpenTrnFrm
            End If
        End If
        GoTo endShortKeyChk
clearShortkey:
        shortkey.Text = OldShortKey
        shortkey.SelStart = Len(OldShortKey)
        shortkey.SelLength = 0
endShortKeyChk:
        OldShortKey = shortkey.Text
    End If
End Sub

Private Sub shortkey_Validate(Cancel As Boolean)
Dim aval As Double
On Error GoTo cancelupdate
    aval = val(shortkey.Text)
    Cancel = False
    GoTo bye
cancelupdate:
    Cancel = True
bye:
    
End Sub

Private Sub SRJournalChk_Click()
    SRJournal = SRJournalChk.value
End Sub

Private Sub TitleList_DblClick()
Dim avar As Long
    avar = TitleList.ItemData(TitleList.ListIndex)
    shortkey.Text = StrPad_(Trim(Str(avar)), 4, "0", "L")
End Sub

Private Sub TitleList_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then
    If KeyCode = vbKeyUp Then
        shortkey.SetFocus
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
        shortkey.SetFocus
        KeyCode = 0
    End If
ElseIf KeyCode = vbKeyReturn Then
Dim avar As Long
    avar = TitleList.ItemData(TitleList.ListIndex)
    shortkey.Text = StrPad_(Trim(Str(avar)), 4, "0", "L")
End If
End Sub

Private Sub CommandToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.key = "F7" Then
       SSTabControl.Tab = IIf(SSTabControl.Tab = 0, SSTabControl.Tabs - 2, SSTabControl.Tab - 1)
    ElseIf Button.key = "F8" Then
       SSTabControl.Tab = IIf(SSTabControl.Tab = SSTabControl.Tabs - 2, 0, SSTabControl.Tab + 1)
    ElseIf Button.key = "F10" Then
'        If cNewJournalType = False Then
'            eJournalFrm.Show vbModal, Me
'        Else
            Dim aTRNHandler As New L2TrnHandler
            aTRNHandler.ExecuteForm "9989"
            aTRNHandler.CleanUp
            Set aTRNHandler = Nothing
'        End If
    ElseIf Button.key = "F11" Then
        Dim bTRNHandler As New L2TrnHandler
        bTRNHandler.ExecuteForm "9747"
        bTRNHandler.CleanUp
        Set bTRNHandler = Nothing
    ElseIf Button.key = "START" Then
        T0611New.Show vbModal, Me
    ElseIf Button.key = "EXIT" Then
        Unload Me
    End If
End Sub

Private Sub OpenTrnFrm()
Dim aStatus As Boolean
StartPos:
Dim astr As String
    On Error GoTo 0
    Dim atrnnode As IXMLDOMElement
    
    Set atrnnode = TrnNodeFromTrnCode(Right("0000" & cTRNCode, 4))
    If Not (atrnnode Is Nothing) Then
        Dim HiddenFlag As Boolean
        HiddenFlag = HiddenFlagFromTrnNode(atrnnode)
        If HiddenFlag Then
            GoTo ExitError
        End If
        Dim aTRNHandler As New L2TrnHandler
        aTRNHandler.ExecuteForm Right("0000" & cTRNCode, 4)
        aTRNHandler.CleanUp
        Set aTRNHandler = Nothing
        Exit Sub
    Else
        On Error Resume Next
        Close #1
        On Error GoTo ExitError
        astr = ReadDir & CStr(cTRNCode) & ".xml"
        Open astr For Input As #1
        Close #1
        
        Dim aTRnFrm As TRNFrm
        Set aTRnFrm = New TRNFrm
        On Error Resume Next:
        Load aTRnFrm:
        On Error GoTo ExitError
        If aTRnFrm.CloseTransactionFlag Then
            Unload aTRnFrm: Exit Sub
        Else
            aTRnFrm.Show vbModal, Me
            Unload aTRnFrm:
            Set aTRnFrm = Nothing
        End If
        
        If TRNQueue.count > 0 Then
            cEnableHiddenTransactions = True
            DoEvents
            cTRNCode = TRNQueue(1)
            GoTo StartPos
        End If
        Exit Sub
    End If
ExitError:
    LogMsgbox "Λάθος Κωδικός Συναλλαγής", vbCritical, "Εφαρμογή OnLine", Err
End Sub

Public Function XML() As String
    Dim anode As IXMLDOMElement
    Set anode = xmlEnvironment.selectSingleNode("//FORMATEDDATE")
    If anode Is Nothing Then
        Set anode = xmlEnvironment.createElement("FORMATEDDATE")
        xmlEnvironment.documentElement.appendChild anode
    End If
    anode.Text = format(Date, "dd/mm/yyyy")
    Set anode = xmlEnvironment.selectSingleNode("//FORMATEDTIME")
    If anode Is Nothing Then
        Set anode = xmlEnvironment.createElement("FORMATEDTIME")
        xmlEnvironment.documentElement.appendChild anode
    End If
    anode.Text = format(Time, "hh:mm:ss")
    
    XML = xmlEnvironment.XML
End Function

