VERSION 5.00
Begin VB.Form T0611New 
   Caption         =   "Έναρξη εφαρμογής"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox InfoList 
      Height          =   1620
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4335
   End
   Begin VB.ComboBox ConnectionType 
      Height          =   315
      ItemData        =   "T0611New.frx":0000
      Left            =   0
      List            =   "T0611New.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton DisconnectBtn 
      Caption         =   "Αποσύνδεση"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton ConnectBtn 
      Caption         =   "Σύνδεση"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Τύπος Σύνδεσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "T0611New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim widthmargin As Long, heightmargin As Long
Dim DocumentManager As New cXMLDocumentManager
Dim handler As New L2TrnHandler
Dim comareas As New Collection
Dim ChiefUserName As String
Private com_status As Integer

Private Function TranslateOData(datanode As IXMLDOMElement)
    InfoList.Clear
    
    Dim user_  As IXMLDOMElement
    Dim username_  As IXMLDOMElement
    Dim userprofile_  As IXMLDOMElement
    Dim profiletype_  As IXMLDOMElement
    Dim irisbranch_  As IXMLDOMElement
    Dim branchname_  As IXMLDOMElement
    Dim codmdldadcent_  As IXMLDOMElement
    Dim atermid_  As IXMLDOMElement
    Dim department_  As IXMLDOMElement
    Dim transcd_  As IXMLDOMElement
    Dim currency_  As IXMLDOMElement
    Dim todaysvaleur_  As IXMLDOMElement
    Dim accountingdate_  As IXMLDOMElement
    Dim accountingdaten_  As IXMLDOMElement
    Dim accountingdatep_  As IXMLDOMElement
    Dim valeur0_  As IXMLDOMElement
    Dim valeur1_  As IXMLDOMElement
    Dim valeur2_  As IXMLDOMElement
    Dim valeur3_  As IXMLDOMElement
    Dim valeur4_  As IXMLDOMElement
    Dim valeur7_  As IXMLDOMElement
   
    

    Set user_ = GetXmlNode(datanode, "./USER_ID", "USER", "ODATA")
    Set username_ = GetXmlNode(datanode, "./USER_NAME", "USER_NAME", "ODATA")
    Set userprofile_ = GetXmlNode(datanode, "./USER_PROFILE", "USER_PROFILE", "ODATA")
    Set profiletype_ = GetXmlNode(datanode, "./PROFILE_TYPE", "PROFILE_TYPE", "ODATA")
    Set irisbranch_ = GetXmlNode(datanode, "./IRIS_BRANCH", "IRIS_BRANCH", "ODATA")
    Set branchname_ = GetXmlNode(datanode, "./BRANCH_NAME", "BRANCH_NAME", "ODATA")
    Set codmdldadcent_ = GetXmlNode(datanode, "./COD_MDLDAD_CENT", "COD_MDLDAD_CENT", "ODATA")
    Set atermid_ = GetXmlNode(datanode, "./ATERM_ID", "ATERM_ID", "ODATA")
    Set department_ = GetXmlNode(datanode, "./DEPARTMENT", "DEPARTMENT", "ODATA")
    Set transcd_ = GetXmlNode(datanode, "./TRAN_SCD", "TRAN_SCD", "ODATA")
    Set currency_ = GetXmlNode(datanode, "./CURRENCY", "CURRENCY", "ODATA")
    Set todaysvaleur_ = GetXmlNode(datanode, "./TODAYS_VALEUR", "TODAYS_VALEUR", "ODATA")
    Set accountingdate_ = GetXmlNode(datanode, "./ACCOUNTING_DATE", "ACCOUNTING_DATE", "ODATA")
    Set accountingdaten_ = GetXmlNode(datanode, "./ACCOUNTING_DATE_N", "ACCOUNTING_DATE_N", "ODATA")
    Set accountingdatep_ = GetXmlNode(datanode, "./ACCOUNTING_DATE_P", "ACCOUNTING_DATE_P", "ODATA")
    Set valeur1_ = GetXmlNode(datanode, "./VALEUR_01", "VALEUR_01", "ODATA")
    Set valeur2_ = GetXmlNode(datanode, "./VALEUR_02", "VALEUR_02", "ODATA")
    Set valeur3_ = GetXmlNode(datanode, "./VALEUR_03", "VALEUR_03", "ODATA")
    Set valeur4_ = GetXmlNode(datanode, "./VALEUR_04", "VALEUR_04", "ODATA")
    Set valeur7_ = GetXmlNode(datanode, "./VALEUR_07", "VALEUR_07", "ODATA")

    
    InfoList.AddItem "Τύπος Σύνδεσης: " & ConnectionType.Text
    InfoList.AddItem "Χρήστης: " & user_.Text
    InfoList.AddItem "Ονομα: " & username_.Text
    
    InfoList.AddItem "Κατάστημα: " & irisbranch_.Text
    InfoList.AddItem "Επωνυμία: " & branchname_.Text
    InfoList.AddItem "Τερματικό: " & Decode_Greek_(atermid_.Text)
    InfoList.AddItem "Λογιστική Ημερομηνία: " & accountingdate_.Text
    InfoList.AddItem "Valeur 0: " & todaysvaleur_.Text
    InfoList.AddItem "Valeur 1: " & valeur1_.Text
    InfoList.AddItem "Valeur 2: " & valeur2_.Text
    InfoList.AddItem "Valeur 3: " & valeur3_.Text
    InfoList.AddItem "Valeur 4: " & valeur4_.Text
    InfoList.AddItem "Valeur 7: " & valeur7_.Text
    
    GenWorkForm.AppBuffers.ByName("ZAFNDLE").ByName("C_ACOD_OU").Value = cBRANCH
    GenWorkForm.AppBuffers.ByName("ZAFNDLE").ByName("C_WKST_ID").Value = Right(MachineName, 4)
    GenWorkForm.AppBuffers.ByName("ZAFNDLE").ByName("C_USR_ID").Value = user_.Text 'UCase(cUserName)

    Dim atermid As String
    atermid = atermid_.Text
    If Len(atermid) > 4 Then atermid = Right(atermid, 4)
    
    With GenWorkForm.AppBuffers.ByName("VCUUP01")
        .ByName("I_ENTP").Value = 1
        .ByName("I_USR_FI").Value = "7942810542"
        .ByName("C_ACOD_FI").Value = "001"
        .ByName("I_USR_OU").Value = 0 '????
        .ByName("C_ACOD_OU").Value = cBRANCH
        .ByName("C_USR_ID").Value = user_.Text
        .ByName("C_WKST_ID").Value = atermid 'Right(MachineName, 4) μηπως????
        .ByName("D_PROC").Value = AsDate(todaysvaleur_.Text)
        .ByName("I_LOC_DFLT_PRFL").Value = profiletype_.Text
        .ByName("C_GEO_PRFL").Value = "GR"
        .ByName("C_PREF_LANG_TP_PRFL").Value = "GRK"
        .ByName("I_CLSF_EXCH_MEDM_K_PRFL").Value = "1100001"
        .ByName("C_CLSF_EXCH_MEDM_K_PRFL").Value = "GRD"
        .ByName("I_CLSF_PRTFL_K_PRFL").Value = "500001"
        .ByName("C_CLSF_PRTFL_K_PRFL").Value = "ΟΛΟΙ"
        .ByName("C_SRCH_PRFL").Value = "EN"
    End With
    
    With GenWorkForm.AppBuffers.ByName("ZAFNELE").ByName("ZAFNFLE")
        .ByName("I_ENTP", 1).Value = 1
        .ByName("I_USR_FI", 1).Value = "7942810542"
        .ByName("C_ACOD_FI", 1).Value = "001"
        .ByName("I_USR_OU", 1).Value = 0 '????
        .ByName("C_ACOD_OU", 1).Value = cBRANCH
        .ByName("C_USR_ID", 1).Value = user_.Text
        .ByName("C_WKST_ID", 1).Value = atermid 'Right(MachineName, 4) μηπως????
        .ByName("D_PROC", 1).Value = AsDate(todaysvaleur_.Text)
        .ByName("I_LOC_DFLT_PRFL", 1).Value = profiletype_.Text
        .ByName("C_GEO_PRFL", 1).Value = "GR"
        .ByName("C_PREF_LANG_TP_PRFL", 1).Value = "GRK"
        .ByName("I_CLSF_EXCH_MEDM_K_PRFL", 1).Value = "1100001"
        .ByName("I_CLSF_PRTFL_K_PRFL", 1).Value = "500001"
        .ByName("C_CLSF_PRTFL_K_PRFL", 1).Value = "ΟΛΟΙ"
    End With
    
    Dim tempComputerName
    GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_O").v2Value("RTRN_CD") = 1
    With GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_I")
        Dim cTime
        cTime = Time
        .v2Value("ID_INTERNO_TERM_TN") = Right(String(8, " ") & cIRISComputerName, 8)
        If Len(cIRISComputerName) > 9 Then
            tempComputerName = Right(String(9, " ") & cIRISComputerName, 9)
            .v2Value("ID_INTERNO_TERM_TN") = Mid(tempComputerName, 1, 5) & Mid(tempComputerName, 7, 3)
        End If
        .v2Value("COD_TX") = "VPU20MOU"
        .v2Value("COD_NRBE_EN") = "0011"
        .v2Value("ID_INTERNO_TERM_TN", 2) = Right(String(8, " ") & cIRISComputerName, 8)
        If Len(cIRISComputerName) > 9 Then
            tempComputerName = Right(String(9, " ") & cIRISComputerName, 9)
            .v2Value("ID_INTERNO_TERM_TN", 2) = Mid(tempComputerName, 1, 5) & Mid(tempComputerName, 7, 3)
        End If
        .v2Value("ID_INTERNO_EMPL_EP") = UCase(cIRISUserName)
        .v2Value("COD_INTERNO_UO") = Right("0000" & Trim(cBRANCH), 4)
        .v2Value("HORA_PC") = 1# * ((hour(cTime) * 60 + Minute(cTime)) * 60 + Second(cTime)) * 1000
        .v2Value("FECHA_PC") = Date
        GenWorkForm.AppBuffers.ByName("STD_TRN_I_PARM_V").Data = .v2Data("STD_TRN_I_PARM_V")
    End With
    
    With GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_O").ByName("TR_CONNECT_IRIS_ICL_EVT_Z", 1).ByName("EP_DATA_V", 1)
        .ByName("ID_INTERNO_EMPL_EP", 1).Value = UCase(cIRISUserName)
        .ByName("NOMB_50", 1).Value = username_.Text
        .ByName("NOM_PERFIL_EN", 1).Value = userprofile_.Text
        '<ID_INTERNO_EMPL_EP>E34000</ID_INTERNO_EMPL_EP>
        '<ID_INTERNO_PE>120896769</ID_INTERNO_PE>
        '<NOM_PERFIL_EN>TODOS</NOM_PERFIL_EN>
        '<FECHA_FIN_PERFIL>31129999</FECHA_FIN_PERFIL>
        '<COD_PERFIL>1</COD_PERFIL>
        '<NOMB_50>ΜΑΡΙΝΟΣ ΓΕΩΡΓΙΟΣ ΔΙΟΝΥ</NOMB_50>
    End With
    
    With GenWorkForm.AppBuffers.ByName("TR_APERTURA_PUESTO_TRN_I")
        cTime = Time
        .v2Value("ID_INTERNO_TERM_TN") = Right(String(8, " ") & cIRISComputerName, 8)
        If Len(cIRISComputerName) > 9 Then
            tempComputerName = Right(String(9, " ") & cIRISComputerName, 9)
            .v2Value("ID_INTERNO_TERM_TN") = Mid(tempComputerName, 1, 5) & Mid(tempComputerName, 7, 3)
        End If
        .v2Value("NUM_SEC") = 0
        .v2Value("COD_TX") = "VPU20MOU"
        .v2Value("FECHA_PC") = Date
        .v2Value("HORA_PC") = 1# * ((hour(cTime) * 60 + Minute(cTime)) * 60 + Second(cTime)) * 1000
        .v2Value("ID_INTERNO_TERM_TN", 2) = UCase(Right(String(8, " ") & cIRISComputerName, 8))
        If Len(cIRISComputerName) > 9 Then
            tempComputerName = Right(String(9, " ") & cIRISComputerName, 9)
            .v2Value("ID_INTERNO_TERM_TN", 2) = Mid(tempComputerName, 1, 5) & Mid(tempComputerName, 7, 3)
        End If
        .v2Value("ID_INTERNO_EMPL_EP") = UCase(cIRISUserName)
        .v2Value("COD_NRBE_EN_FSC") = "0011"
        .v2Value("COD_NRBE_EN") = "0011"
        .v2Value("COD_INTERNO_UO") = Right("0000" & Trim(cBRANCH), 4)
    End With
    
    'With GenWorkForm.AppBuffers.ByName("TR_CONNECT_IRIS_ICL_TRN_O")
    '    GenWorkForm.AppBuffers.ByName("STD_AN_AL_MSJ_V").data = .v2Data("STD_AN_AL_MSJ_V")
    'End With
    'ShowIRISMessages_ GenWorkForm.AppBuffers.ByName("STD_AN_AL_MSJ_V")
                    
    With GenWorkForm.AppBuffers
        .ByName("TR_CONNECT_IRIS_ICL_TRN_O").v2Value("COD_NRBE_EN") = "0011"
        .ByName("TR_CONNECT_IRIS_ICL_TRN_O").v2Value("FECHA_CTBLE") = AsDate(accountingdate_.Text)
        With .ByName("TR_CONNECT_IRIS_ICL_TRN_O").ByName("TR_CONNECT_IRIS_ICL_EVT_Z").ByName("BRANCH_DATA_V", 1)
            .ByName("COD_INTERNO_UO", 1).Value = irisbranch_.Text
            .ByName("COD_CSB_OF", 1).Value = irisbranch_.Text
            .ByName("NOMB_CENT_UO", 1).Value = branchname_.Text
            .ByName("COD_MDLDAD_CENT", 1).Value = codmdldadcent_.Text
            '.v2Value("IND_CENT_FICTIC_UO") = "N" ???
            '.v2Value("IND_CENT_CTRL_UO") = "S" ???
        End With
        .ByName("BRANCH_DATA_V").Data = .ByName("TR_CONNECT_IRIS_ICL_TRN_O").v2Data("BRANCH_DATA_V")
        With .ByName("UO_CENTRO_E")
            .v2Value("COD_NRBE_EN") = "0011"
            .v2Value("COD_INTERNO_UO") = irisbranch_.Text
            .v2Value("NOMB_CENT_UO") = branchname_.Text
            '.v2Value("NUM_AR_GEO") = .ByName("BRANCH_DATA_V").v2Value("NUM_AR_GEO") ???
            .v2Value("COD_MDLDAD_CENT") = codmdldadcent_.Text
            .v2Value("COD_CSB_OF") = irisbranch_.Text
        End With
                    
        .ByName("TR_CONS_CENTRO_TRN_O").v2Data("UO_CENTRO_E") = .ByName("UO_CENTRO_E").Data
        .ByName("TR_CONS_CENTRO_TRN_O").v2Data("PY_PARAM_VVV_E") = .ByName("PY_PARAM_VVV_E").Data '???
                    
        .ByName("TR_APERTURA_PUESTO_TRN_O").ByName("TR_APERTURA_PUESTO_EVT_Z").ByName("FECHA_CTBLE", 1).Value = _
            .ByName("TR_CONNECT_IRIS_ICL_TRN_O").v2Value("FECHA_CTBLE")
                    
        UpdatexmlEnvironment UCase("IRISPostDate"), Replace(.ByName("TR_APERTURA_PUESTO_TRN_O").ByName("TR_APERTURA_PUESTO_EVT_Z").ByName("FECHA_CTBLE", 1).FormatedDate8, "/", "")
        UpdatexmlEnvironment UCase("IRISBranchProfile"), codmdldadcent_.Text
      
    End With
    
    Dim I As Integer
    For I = 0 To InfoList.ListCount - 1
        eJournalWrite InfoList.list(I)
    Next I
    
End Function

Private Function GetChiefKey() As Boolean
'αίτηση για κλειδι Chief Teller
    On Error GoTo ErrorPos
    GetChiefKey = False
    ManagerRequest = False
    If isChiefTeller Then
        KeyAccepted = False: ChiefRequest = True: Load KeyWarning: Set KeyWarning.owner = Screen.activeform
        KeyWarning.Show vbModal, Screen.activeform
    Else
        KeyAccepted = False: Set SelKeyFrm.owner = Screen.activeform: ChiefRequest = True
        SelKeyFrm.Show vbModal, Screen.activeform
    End If
    If Not KeyAccepted Then Exit Function
    ChiefUserName = cIRISUserName
    If Not (isChiefTeller) Then ChiefUserName = cCHIEFUserName Else cCHIEFUserName = cIRISUserName
    GetChiefKey = True
    Exit Function
ErrorPos:
    
End Function

Private Sub ConnectBtn_Click()
    
    Dim response
    Dim d As Integer
    
    If ConnectionType.ListIndex < 0 Then
        NBG_MsgBox "Επιλέξτε Τύπο Σύνδεσης", True, _
                "Πρόβλημα στη διαδικασία Σύνδεσης..."
        Exit Sub
    End If
    
    If ConnectionType.ListIndex = 1 Then
    
        response = MsgBox("Προσοχή! Σύνδεση με απογευματινή λειτουργία ταμείου....", vbOKCancel, "Προσοχή!!!")
        eJournalWrite "Προσοχή! Σύνδεση με απογευματινή λειτουργία ταμείου...."
        
        If response = vbOK Then
            eJournalWrite "Αποδοχή"
        ElseIf response = vbCancel Then
            eJournalWrite "’κυρο"
            Exit Sub
        End If
    End If
    
    cPOSTDATE = Date
        
    Dim adoc As MSXML2.DOMDocument30
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If adoc Is Nothing Then Exit Sub
    
    Dim Node As IXMLDOMElement
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='SSTRT']", "SSTRT", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If Node Is Nothing Then Exit Sub

    Dim methodnode As IXMLDOMElement
    Set methodnode = Node.selectSingleNode("./method")

    If ConnectionType.ListIndex < 3 Then
        If Not GetChiefKey Then
            Call NBG_LOG_MsgBox("ΔΕΝ ΔΟΘΗΚΕ ΕΓΚΡΙΣΗ ΠΡΟΙΣΤΑΜΕΝΟΥ")
            Exit Sub
        End If
    End If
        
    
    Dim ComArea As cXmlComArea
    Set ComArea = New cXmlComArea
    Set ComArea.content = Node
    
    'διορθωση cXmlComArea.Container
    'Set ComArea.owner = DocumentManager
    Set ComArea.Container = DocumentManager.TrnBuffers
    
    Dim datanode As IXMLDOMElement
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "SSTRT", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If datanode Is Nothing Then Exit Sub
    
    Dim datadoc As MSXML2.DOMDocument30
    Set datadoc = XmlLoadString(datanode.xml, "DataDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If datadoc Is Nothing Then Exit Sub
    
    Dim trankey As IXMLDOMElement
    Dim AuthUser As IXMLDOMElement
    Dim branch As IXMLDOMElement
    Dim termid As IXMLDOMElement
    Dim operation As IXMLDOMElement
    
    Set trankey = GetXmlNode(datadoc.documentElement, "//NT_HEADER//TRAN_KEY", "TRAN_KEY", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If trankey Is Nothing Then Exit Sub
    Set AuthUser = GetXmlNode(datadoc.documentElement, "//NT_HEADER/AUTHORISATION/AUTH_USER", "AUTH_USER", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If AuthUser Is Nothing Then Exit Sub
    Set branch = GetXmlNode(datadoc.documentElement, "//STDDATA/BRANCH", "BRANCH", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If branch Is Nothing Then Exit Sub
    Set termid = GetXmlNode(datadoc.documentElement, "//STDDATA/ATERM_ID", "STDDATA_ATERM_ID", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If termid Is Nothing Then Exit Sub
    Set operation = GetXmlNode(datadoc.documentElement, "//IDATA/OPERATION_MODE", "OPERATION_MODE", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If operation Is Nothing Then Exit Sub
    
    If ConnectionType.ListIndex < 3 Then
        trankey.Text = "C.T."
        AuthUser.Text = UCase(ChiefUserName)
    Else
        trankey.Text = "TELLER"
    End If
    branch.Text = cBRANCH

    
    operation.Text = ConnectionType.ItemData(ConnectionType.ListIndex)
        
    Dim Result As String
    Flag610 = True
    Result = ComArea.LoadXML(datadoc.xml)
    Flag610 = False
    Dim resultdoc As MSXML2.DOMDocument30
    Set resultdoc = XmlLoadString(Result, "ResultDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If resultdoc Is Nothing Then Me.Enabled = True: Exit Sub
    If Not ComArea.HandleResp(resultdoc) Then Me.Enabled = True: Exit Sub
    Dim odatanode As IXMLDOMElement
    Set odatanode = GetXmlNode(resultdoc.documentElement, "//ODATA", "ODATA", "ResultDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If odatanode Is Nothing Then Me.Enabled = True: Exit Sub
    TranslateOData odatanode
    
    Dim authuser_ As IXMLDOMElement
    Set authuser_ = GetXmlNode(resultdoc.documentElement, "//STDDATA/AUTH_USER", "AUTH_USER", "data/comarea")
    Dim key_ As IXMLDOMElement
    Set key_ = GetXmlNode(resultdoc.documentElement, "//STDDATA/USERKEY", "USERKEY", "data/comarea")
    
    InfoList.AddItem "Έγκριση:  " & " " & key_.Text & ": " & authuser_.Text
    eJournalWrite "Έγκριση:  " & " " & key_.Text & ": " & authuser_.Text
    
    Dim headnode As IXMLDOMElement
    Dim termidnode As IXMLDOMElement
    Dim departmentnode As IXMLDOMElement
    
    Set headnode = GetXmlNode(resultdoc.documentElement, "//ODATA/TRAN_SCD", "TRAN_SCD", "ODATA", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If headnode Is Nothing Then Me.Enabled = True: Exit Sub
    
    Set termidnode = GetXmlNode(resultdoc.documentElement, "//ODATA/ATERM_ID", "ATERM_ID", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If termidnode Is Nothing Then Me.Enabled = True: Exit Sub
    Set departmentnode = GetXmlNode(resultdoc.documentElement, "//ODATA/DEPARTMENT", "DEPARTMENT", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If departmentnode Is Nothing Then Me.Enabled = True: Exit Sub
    
    Dim kmodeNode As IXMLDOMElement
    Set kmodeNode = GetXmlNodeIfPresent(resultdoc.documentElement, "//IDATA/DEPOSITS_MODE")
    If Not kmodeNode Is Nothing Then
        cKMODEFlag = True
        cKMODEValue = Trim(kmodeNode.Text)
    End If
    
    cTERMINALID = Decode_Greek_(termidnode.Text)
    termid.Text = cTERMINALID
    
    cDepartment = departmentnode.Text
    UpdatexmlEnvironment "DEPARTMENT", cDepartment
    
    cHEAD = headnode.Text
    UpdatexmlEnvironment "TERMINALID", cTERMINALID
    UpdatexmlEnvironment "SESSIONCD", CStr(cHEAD)
    UpdatexmlEnvironment "OPERATIONMODE", operation.Text
    UpdatexmlEnvironmentNode GenWorkForm.AppBuffers.ByName("STD_TRN_I_PARM_V").GetXMLView.documentElement
    
    Dim xmldoc As New MSXML2.DOMDocument60
    Dim s As IXMLDOMElement
    Dim n As IXMLDOMNode
    Set s = xmldoc.createElement("CUF_USR_OL_PRFL_D")
    Set xmldoc.documentElement = s
    
    Set n = xmldoc.createNode(NODE_ELEMENT, "I_ENTP", "")
    n.Text = "1"
    s.appendChild n
    
    Set n = xmldoc.createNode(NODE_ELEMENT, "C_ACOD_FI", "")
    n.Text = "001"
    s.appendChild n
    
    Set n = xmldoc.createNode(NODE_ELEMENT, "C_ACOD_OU", "")
    n.Text = cBRANCH
    s.appendChild n
        
    Dim aUser As String
    aUser = UCase(cUserName)
    If cIRISUserName <> "" Then
        aUser = UCase(cIRISUserName)
    End If
    Set n = xmldoc.createNode(NODE_ELEMENT, "C_USR_ID", "")
    n.Text = aUser
    s.appendChild n
    
    Dim amachine As String
    amachine = Right(MachineName, 4)
    If cIRISComputerName <> "" Then
     amachine = UCase(Right(String(4, " ") & cIRISComputerName, 4))
    End If
    Set n = xmldoc.createNode(NODE_ELEMENT, "C_WKST_ID", "")
    n.Text = amachine
    s.appendChild n
    
    UpdatexmlEnvironmentNode xmldoc.documentElement
    
    GenWorkForm.UpdateStationInfo
        
    UpdateChiefKey ""
    UpdateManagerKey ""
    
    Flag610 = True
    Flag620 = True
    Flag630 = True
    
    GenWorkForm.vStatus.Panels(2).Visible = True
    GenWorkForm.vStatus.Panels(3).Visible = False

    SaveJournal
    
    Exit Sub
processingError:
    Dim aMsg As String
    aMsg = "Απέτυχε η εκτέλεση της διαδικασίας: " & "ConnectBtn_Click" & " " & Err.number & " " & Err.description & vbCrLf
    LogMsgbox aMsg, vbCritical, "Λάθος ..."
End Sub

Private Sub DisconnectBtn_Click()
    Dim adoc As MSXML2.DOMDocument30
'    Set adoc = XmlLoadFile(ReadDir & "\InitBlocks.xml", "InitBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    Set adoc = XmlLoadFile(ReadDir & "\XmlBlocks.xml", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If adoc Is Nothing Then Exit Sub
    
    Dim Node As IXMLDOMElement
'    Set Node = GetXmlNode(adoc.documentElement, "//comarea[@name='SFINI']", "SFINI", "InitBlock", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    Set Node = GetXmlNode(adoc.documentElement, "//comareas/comarea[@name='SFINI']", "SFINI", "XmlBlocks", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If Node Is Nothing Then Exit Sub
    
    Dim ComArea As New cXmlComArea
    Set ComArea.content = Node
    'διορθωση cXmlComArea.Container
    'Set ComArea.owner = DocumentManager
    Set ComArea.Container = DocumentManager.TrnBuffers
    
    Dim datanode As IXMLDOMElement
    Set datanode = GetXmlNode(Node, "./data/comarea", "data/comarea", "SFINI", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If datanode Is Nothing Then Exit Sub
    
    Dim datadoc As MSXML2.DOMDocument30
    Set datadoc = XmlLoadString(datanode.xml, "DataDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If datadoc Is Nothing Then Exit Sub
    
    'Dim operation As IXMLDOMElement
    'Set operation = GetXmlNode(datadoc.documentElement, "//IDATA/OPERATION_MODE", "OPERATION_MODE", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    'If operation Is Nothing Then Exit Sub
    Dim wsid As IXMLDOMElement
    Set wsid = GetXmlNode(datadoc.documentElement, "//IDATA/WS_ID", "WS_ID", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If wsid Is Nothing Then Exit Sub
    Dim userid As IXMLDOMElement
    Set userid = GetXmlNode(datadoc.documentElement, "//IDATA/USER_ID", "USER_ID", "data/comarea", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If userid Is Nothing Then Exit Sub
    
    wsid.Text = UCase(cIRISComputerName)
    userid.Text = UCase(cIRISUserName)
        
    Dim Result As String
    Dim aFlag As Boolean
    aFlag = Flag610
    Flag610 = True
    Result = ComArea.LoadXML(datadoc.xml)
    Flag610 = aFlag
    
    Dim resultdoc As MSXML2.DOMDocument30
    Set resultdoc = XmlLoadString(Result, "ResultDoc", "Πρόβλημα στη διαδικασία Σύνδεσης...")
    If resultdoc Is Nothing Then
        Exit Sub
    End If
    If Not ComArea.HandleResp(resultdoc) Then Exit Sub
    InfoList.Clear
    InfoList.AddItem "Η αποσύνδεση ολοκληρώθηκε"
    
    GenWorkForm.vStatus.Panels(2).Visible = False
    GenWorkForm.vStatus.Panels(3).Visible = True
    
    eJournalWrite "Η αποσύνδεση ολοκληρώθηκε"
    
    Flag610 = False
    Flag620 = False
    Flag630 = False
        
    SaveJournal
End Sub

Private Sub Form_Load()
    cTRNCode = "0610"
    
    Set handler.DocumentManager = DocumentManager
    Set handler.activeform = Me
    Set DocumentManager.owner = handler
    
    ConnectionType.AddItem "Πρωινή Λειτουργία Ταμείου"
    ConnectionType.ItemData(ConnectionType.NewIndex) = 1
    ConnectionType.AddItem "Απογευματινή Λειτουργία Ταμείου"
    ConnectionType.ItemData(ConnectionType.NewIndex) = 2
    ConnectionType.AddItem "Πληροφοριακές Συναλλαγές Online-IRIS"
    ConnectionType.ItemData(ConnectionType.NewIndex) = 3
    ConnectionType.AddItem "Πληροφοριακές Συναλλαγές IRIS"
    ConnectionType.ItemData(ConnectionType.NewIndex) = 4
    ConnectionType.ListIndex = -1
    
    widthmargin = width - InfoList.width
    heightmargin = height - InfoList.height
    
    Dim I As Integer
    For I = 1 To ConnectionInfo.Count
        InfoList.AddItem ConnectionInfo(I)
    Next I
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Control, TControl, foundflag As Boolean
    If Me.Enabled Then
        foundflag = False
        If KeyCode = 65 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-a
            KeyCode = 0
            Load BufferViewer: Set BufferViewer.owner = Me: Set BufferViewer.inBufferList = GenWorkForm.AppBuffers
            BufferViewer.Show vbModal, Me
            Unload BufferViewer
        ElseIf KeyCode = 66 And ((Shift And vbCtrlMask) > 0) Then 'ctrl-b
            KeyCode = 0
            Load BufferViewer: Set BufferViewer.owner = Me:
            Set BufferViewer.inBufferList = DocumentManager.TrnBuffers
            BufferViewer.Show vbModal, Me
            Unload BufferViewer
        ElseIf KeyCode = 27 Then
            KeyCode = 0
            Unload Me
        End If
    End If
End Sub

Public Sub sbWriteStatusMessage(ByVal sMessage As String)
   GenWorkForm.sbWriteStatusMessage sMessage
End Sub

Private Sub Form_Resize()
    InfoList.width = width - widthmargin
    ConnectionType.width = IIf(width - widthmargin > 0, width - widthmargin, 0)
    InfoList.height = IIf(height - heightmargin > 0, height - heightmargin, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ConnectionInfo = New Collection
    Dim Item As String
    Dim I As Integer
    For I = 0 To InfoList.ListCount - 1
        ConnectionInfo.add InfoList.list(I)
    Next I
End Sub
