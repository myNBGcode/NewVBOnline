VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSoapClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public VirtualDirectory As String
Public Methods As New Collection

Public Function Method(inMethodname As String) As cSoapClientMethod
    Dim aMethod As cSoapClientMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods.item(inMethodname)
continueproc:
    On Error GoTo 0
    Set Method = aMethod
End Function

Public Function AddParam(inMethodname As String, ParamName As String) As cSoapClientMethod
    Dim aMethod As cSoapClientMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.AddParam ParamName
    Set AddParam = aMethod
End Function

Public Function AddParamValue(inMethodname As String, ParamName As String, ParamValue) As cSoapClientMethod
    Dim aMethod As cSoapClientMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.AddParamValue ParamName, ParamValue
    Set AddParamValue = aMethod
End Function

Public Function DefineMethod(inMethodname As String, Optional PNameArray, Optional PValueArray) As cSoapClientMethod
    Set DefineMethod = Nothing
    Dim aMethod As New cSoapClientMethod
    Methods.add aMethod, inMethodname
    aMethod.methodname = inMethodname
    If IsMissing(PNameArray) Then aMethod.PNameArray = Array() Else aMethod.PNameArray = PNameArray
    If IsMissing(PValueArray) Then aMethod.PValueArray = Array() Else aMethod.PValueArray = PValueArray
    
    Set aMethod.owner = Me
    Set DefineMethod = aMethod
End Function

Public Function ExecMethod(inMethodname As String, Optional PValueArray) As cSoapClientMethod

    Set ExecMethod = Nothing
    Dim aMethod As cSoapClientMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.PrepareMethod PValueArray
    aMethod.ExecMethod
    Set ExecMethod = aMethod
    
End Function

Public Function ExecSQLXmlMethod(inMethodname As String, Optional PValueArray) As cSoapClientMethod

    Set ExecSQLXmlMethod = Nothing
    Dim aMethod As cSoapClientMethod
    Set aMethod = Nothing
    On Error GoTo continueproc
    Set aMethod = Methods(inMethodname)
continueproc:
    On Error GoTo 0
    If aMethod Is Nothing Then Exit Function
    aMethod.PrepareSQLXmlMethod PValueArray
    aMethod.ExecSQLXmlMethod
    Set ExecSQLXmlMethod = aMethod
    
End Function



