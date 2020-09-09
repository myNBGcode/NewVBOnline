VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXMLWebLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public content As IXMLDOMElement
Public VirtualDirectory As String
Public Methods As New Collection

Public Function Method(inMethodname As String) As cXMLWebMethod
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods.item(inMethodname)
continueproc:
    On Error GoTo 0
    Set Method = aMethod
End Function

Public Function AddParam(inMethodname As String, ParamName As String) As cXMLWebMethod
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.AddParam ParamName
    Set AddParam = aMethod
End Function

Public Function AddParamValue(inMethodname As String, ParamName As String, ParamValue) As cXMLWebMethod
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.AddParamValue ParamName, ParamValue
    Set AddParamValue = aMethod
End Function

Public Function DefineMethod(inMethodname As String, inNamespace As String, Optional PNameArray, Optional PValueArray) As cXMLWebMethod
    Set DefineMethod = Nothing
    Dim aMethod As New cXMLWebMethod
    Methods.add aMethod, inMethodname
    aMethod.methodname = inMethodname
    aMethod.namespace = inNamespace
    aMethod.isDocument = False
    If IsMissing(PNameArray) Then aMethod.PNameArray = Array() Else aMethod.PNameArray = PNameArray
    If IsMissing(PValueArray) Then aMethod.PValueArray = Array() Else aMethod.PValueArray = PValueArray
    
    Set aMethod.owner = Me
    Set DefineMethod = aMethod
End Function

Public Function DefineDocumentMethod(inMethodname As String, inNamespace As String) As cXMLWebMethod
    Set DefineDocumentMethod = Nothing
    Dim aMethod As New cXMLWebMethod
    Methods.add aMethod, inMethodname
    aMethod.methodname = inMethodname
    aMethod.namespace = inNamespace
    aMethod.isDocument = True
    aMethod.PNameArray = Array()
    aMethod.PValueArray = Array()
    
    Set aMethod.owner = Me
    Set DefineDocumentMethod = aMethod
    Set aMethod = Nothing
End Function

Public Function ExecMethod(inMethodname As String, Optional PValueArray) As cXMLWebMethod
    Set ExecMethod = Nothing
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    'UpdateTrnNum_
    aMethod.PrepareMethod PValueArray
    aMethod.ExecMethod
    Set ExecMethod = aMethod
    
End Function

Public Function ExecMethod_usingXml(inMethodname As String, Paramdoc As IXMLDOMElement) As cXMLWebMethod
    Set ExecMethod_usingXml = Nothing
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    'UpdateTrnNum_
    aMethod.PrepareMethod_usingXml Paramdoc
    aMethod.ExecMethod
    Set ExecMethod_usingXml = aMethod
End Function

Public Function ExecDocumentMethod(inMethodname As String, Optional inDocument) As cXMLWebMethod
    Set ExecDocumentMethod = Nothing
    Dim aMethod As cXMLWebMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    'UpdateTrnNum_
    eJournalWrite "диадийасиа: " & inMethodname
    aMethod.PrepareDocumentMethod inDocument
    aMethod.ExecMethod
    Set ExecDocumentMethod = aMethod
    Set aMethod = Nothing
    
End Function

'Public Function ExecDocumentMethod_Old(inMethodname As String, Optional inDocument) As cXMLWebMethod
'    Set ExecDocumentMethod_Old = Nothing
'    Dim aMethod As cXMLWebMethod
'    Set aMethod = Nothing
'    On Error GoTo continueproc
'    Set aMethod = Methods(inMethodname)
'continueproc:
'    On Error GoTo 0
'    If aMethod Is Nothing Then Exit Function
'    UpdateTrnNum_
'    aMethod.PrepareDocumentMethod_Old inDocument
'    aMethod.ExecMethod
'    Set ExecDocumentMethod_Old = aMethod
'End Function
'
 



