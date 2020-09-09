VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCurrencyPassbookLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public EIDOS As String
Public TRN_DATE As String
Public KATTERM As String
Public TRN_NOM As String
Public TRN_POSO As String
Public COD_AIT As String
Public LEK_AIT As String
Public LOG_POSO As String
Public TRN_SIGN As String
Public YPOL_LOG As String
Public YPOL_SIGN As String
Public EPITOKIO As String
Public VALEUR_ANAN As String
Public LHXH As String


Public Sub PrepareEntry(LineEntry As IXMLDOMElement)
    If Not (LineEntry.selectSingleNode("./BIBL_EIDOS") Is Nothing) Then EIDOS = LineEntry.selectSingleNode("./BIBL_EIDOS").Text
    If Not (LineEntry.selectSingleNode("./BIBL_TRN_DATE") Is Nothing) Then TRN_DATE = LineEntry.selectSingleNode("./BIBL_TRN_DATE").Text
    If Not (LineEntry.selectSingleNode("./BIBL_KATTERM") Is Nothing) Then KATTERM = LineEntry.selectSingleNode("./BIBL_KATTERM").Text
    If Not (LineEntry.selectSingleNode("./BIBL_TRN_NOM") Is Nothing) Then TRN_NOM = LineEntry.selectSingleNode("./BIBL_TRN_NOM").Text
    If Not (LineEntry.selectSingleNode("./BIBL_TRN_POSO") Is Nothing) Then TRN_POSO = LineEntry.selectSingleNode("./BIBL_TRN_POSO").Text
    If Not (LineEntry.selectSingleNode("./BIBL_COD_AIT") Is Nothing) Then COD_AIT = LineEntry.selectSingleNode("./BIBL_COD_AIT").Text
    If Not (LineEntry.selectSingleNode("./BIBL_LEK_AIT") Is Nothing) Then LEK_AIT = LineEntry.selectSingleNode("./BIBL_LEK_AIT").Text
    If Not (LineEntry.selectSingleNode("./BIBL_LOG_POSO") Is Nothing) Then LOG_POSO = LineEntry.selectSingleNode("./BIBL_LOG_POSO").Text
    If Not (LineEntry.selectSingleNode("./BIBL_TRN_SIGN") Is Nothing) Then TRN_SIGN = LineEntry.selectSingleNode("./BIBL_TRN_SIGN").Text
    If Not (LineEntry.selectSingleNode("./BIBL_YPOL_LOG") Is Nothing) Then YPOL_LOG = LineEntry.selectSingleNode("./BIBL_YPOL_LOG").Text
    If Not (LineEntry.selectSingleNode("./BIBL_YPOL_SIGN") Is Nothing) Then YPOL_SIGN = LineEntry.selectSingleNode("./BIBL_YPOL_SIGN").Text
    If Not (LineEntry.selectSingleNode("./BIBL_EPITOKIO") Is Nothing) Then EPITOKIO = LineEntry.selectSingleNode("./BIBL_EPITOKIO").Text
    If Not (LineEntry.selectSingleNode("./BIBL_VALEUR_ANAN") Is Nothing) Then VALEUR_ANAN = LineEntry.selectSingleNode("./BIBL_VALEUR_ANAN").Text
    If Not (LineEntry.selectSingleNode("./BIBL_LHXH") Is Nothing) Then LHXH = LineEntry.selectSingleNode("./BIBL_LHXH").Text
End Sub


Public Function Lines() As Collection
    Set Lines = New Collection
    Lines.add EIDOS & " Line1: " & TRN_POSO
    Lines.add "Line2: " & LEK_AIT
End Function
