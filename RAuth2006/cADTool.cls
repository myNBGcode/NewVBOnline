VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cADTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objRootDSE, objNetwork, UserName
Dim strUserDN, objGroupList, objUser, strDNSDomain
'Dim strComputer, objComputer, strComputerDN,
'Dim strHomeDrive, strHomeShare, objTrans, strNetBIOSDomain, strNTName , objApp, objCommand,
Dim objConnection, strbase, strattributes
Dim strDN, strGN, strsn, strSA, strUN

Public UserGroups As New Collection

' Constants for the NameTranslate object.
'Const ADS_NAME_INITTYPE_GC = 3
'Const ADS_NAME_TYPE_NT4 = 3
'Const ADS_NAME_TYPE_1779 = 1

Public Sub Initialize()
    Set objNetwork = CreateObject("Wscript.Network")
    UserName = ""
    UserName = objNetwork.UserName
    Set objRootDSE = GetObject("LDAP://RootDSE")
    strDNSDomain = objRootDSE.Get("defaultNamingContext")
    
    Dim objCommand, objRecordSet, strfilter, strattributes, strQuery, strbase
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection

    strbase = "<LDAP://" & strDNSDomain & ">"
    'Define the filter elements
    strfilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & UserName & "))"
    'List all attributes you will require
    strattributes = "distinguishedName,sAMAccountName,givenName,sn,userPrincipalName"
    'compose query
    strQuery = strbase & ";" & strfilter & ";" & strattributes & ";subtree"
    objCommand.CommandText = strQuery
    objCommand.Properties("Page Size") = 99999
    objCommand.Properties("Timeout") = 300
    objCommand.Properties("Cache Results") = False
    Set objRecordSet = objCommand.Execute
    objRecordSet.MoveFirst
    If Not objRecordSet.EOF Then
        strDN = objRecordSet.Fields("distinguishedName")
        strDN = Replace(strDN, "/", "\/")
        strGN = objRecordSet.Fields("givenName")
        strsn = objRecordSet.Fields("sn")
        strSA = objRecordSet.Fields("sAMAccountName")
        strUN = objRecordSet.Fields("userPrincipalName")
    End If
    objRecordSet.Close
    



    
    ' Loop required for Win9x clients during logon.
'    On Error Resume Next
'    Do While strNTName = ""
'      Err.Clear
'      If wscript.Version > 5 Then
'        wscript.Sleep 100
'      End If
'    Loop
    'strNTName = "e31155"
    On Error GoTo 0
    
    ' Determine DNS domain name from RootDSE object.
    
    ' Use the NameTranslate object to find the NetBIOS domain name from the
    ' DNS domain name.
    ''Set objTrans = CreateObject("NameTranslate")
    ''objTrans.init ADS_NAME_INITTYPE_GC, ""
    ''objTrans.Set ADS_NAME_TYPE_1779, strDNSDomain
    ''strNetBIOSDomain = objTrans.Get(ADS_NAME_TYPE_NT4)
    ' Remove trailing backslash.
    ''strNetBIOSDomain = Left(strNetBIOSDomain, Len(strNetBIOSDomain) - 1)
    
    ' Use the NameTranslate object to convert the NT user name to the
    ' Distinguished Name required for the LDAP provider.
    ''objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain & "\" & strNTName
    ''strUserDN = objTrans.Get(ADS_NAME_TYPE_1779)
    'strUserDN = "CN=BikS,OU=DevelopmentTeam,DC=central,DC=nbg,DC=gr"
    ' Bind to the user object in Active Directory with the LDAP provider.
    Set objUser = GetObject("LDAP://" & strDN)
    
    ' Use the NameTranslate object to convert the NT name of the computer to
    ' the Distinguished name required for the LDAP provider. Computer names
    ' must end with "$".
    'strComputer = objNetwork.ComputerName
    'objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain & "\" & strComputer & "$"
    'strComputerDN = objTrans.Get(ADS_NAME_TYPE_1779)
    
    ' Bind to the computer object in Active Directory with the LDAP
    ' provider.
    'Set objComputer = GetObject("LDAP://" & strComputerDN)
    
    
    LoadGroups objUser
    
    'Set objApp = CreateObject("WScript.Shell")
    'objApp.Run """ConnectToBasicDrives.cmd"
    'objApp.Run """Default Users.cmd"""
    'objApp.Run """GenerateInitialUserProfile.cmd"""
    '
    'If IsMember(objUser, "3270 Users") Then
    '  On Error Resume Next
    '  wscript.echo "3270 Users.cmd OK"
    'objApp.Run """3270 Users.cmd"""
    'End If
    '
    'If IsMember(objUser, "Asyk Users") Then
    '  On Error Resume Next
    '  wscript.echo "Asyk Users.cmd OK"
    'objApp.Run """Asyk Users.cmd"""
    'End If
    '
    'If IsMember(objUser, "Beta93 Users") Then
    '  On Error Resume Next
    '  wscript.echo "Beta93 Users.cmd OK"
    'objApp.Run """Beta93 Users.cmd"""
    'End If
    '
    'If IsMember(objUser, "Chief Teller") Then
    '  On Error Resume Next
    '  wscript.echo "Chief Teller.cmd OK"
    'objApp.Run """Chief Teller.cmd"""
    'End If
    '
    '
    'If IsMember(objUser, "CRA Users") Then
    '  On Error Resume Next
    '  wscript.echo "CRA Users.cmd OK"
    'objApp.Run """CRA Users.cmd"""
    'End If
    '
    '
    'If IsMember(objUser, "Import Users") Then
    '  On Error Resume Next
    '  wscript.echo "Import Users.cmd OK"
    'objApp.Run """Import Users.cmd"""
    'End If
    '
    'If IsMember(objUser, "Manager") Then
    '  On Error Resume Next
    '  wscript.echo "Manager.cmd OK"
    'objApp.Run """Manager.cmd"""
    'End If
    '
    '
    'If IsMember(objUser, "Online Info") Then
    '  On Error Resume Next
    '  wscript.echo "Online Info.cmd OK"
    'objApp.Run """Online Info.cmd"""
    'End If
    '
    '
    'If IsMember(objUser, "Pese Users") Then
    '  On Error Resume Next
    '  wscript.echo "Pese Users.cmd OK"
    'objApp.Run """Pese Users.cmd"""
    'End If
    '
    'If IsMember(objUser, "Teller") Then
    '  On Error Resume Next
    '  wscript.echo "Teller OK"
    'objApp.Run """Teller.cmd"""
    'End If
    
    ' Clean up.
'    If IsObject(objConnection) Then
'      objConnection.Close
'      Set objCommand = Nothing
'      Set objConnection = Nothing
'      Set objGroupList = Nothing
'    End If
'    Set objRootDSE = Nothing
'    Set objTrans = Nothing
'    Set objNetwork = Nothing
'    Set objUser = Nothing
'    Set objComputer = Nothing

End Sub

Public Function IsMember(userObject, ntGroupName) As Boolean
' Function to test for group membership.
' objADObject is a user or computer object.
' strGroupNTName is the NT name (sAMAccountName) of the group to test.
' objGroupList is a dictionary object, with global scope.
' Returns True if the user or computer is a member of the group.
' Subroutine LoadGroups is called once for each different objADObject.

' The first time IsMember is called, setup the dictionary object
' and objects required for ADO.
  'If Not objGroupList.Exists(userObject.sAMAccountName & "\") Then
    ' Dictionary object established, but group memberships for this
    ' user or computer must be added.
  '  LoadGroups userObject
  'End If
  ' Return True if this user or computer is a member of the group.
  ' IsMember = objGroupList.Exists(userObject.sAMAccountName & "\" & ntGroupName)
  IsMember = objGroupList.Exists(ntGroupName)
End Function

Sub LoadGroups(objADObject)
    If Not IsEmpty(objGroupList) Then
        Exit Sub
    End If
    
    Set objGroupList = CreateObject("Scripting.Dictionary")
    objGroupList.CompareMode = vbTextCompare
    Dim objCommand

    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection

    Set objRootDSE = GetObject("LDAP://RootDSE")
    strDNSDomain = objRootDSE.Get("defaultNamingContext")

    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Cache Results") = False


    ' Search entire domain.
    strbase = "<LDAP://" & strDNSDomain & ">"
    ' Retrieve NT name of each group.
    strattributes = "sAMAccountName"

    ' Load group memberships for this user or computer into dictionary
    ' object.

' Subroutine to populate dictionary object with group memberships.
' objGroupList is a dictionary object, with global scope. It keeps track
' of group memberships for each user or computer separately. ADO is used
' to retrieve the name of the group corresponding to each objectSid in
' the tokenGroup array. Based on an idea by Joe Kaplan.

  Dim arrbytGroups, k, strfilter, objRecordSet, strGroupName, strQuery

  ' Add user name to dictionary object, so LoadGroups need only be
  ' called once for each user or computer.
  'objGroupList(objADObject.sAMAccountName & "\") = True

  ' Retrieve tokenGroups array, a calculated attribute.
  objADObject.GetInfoEx Array("tokenGroups"), 0
  arrbytGroups = objADObject.Get("tokenGroups")

  ' Create a filter to search for groups with objectSid equal to each
  ' value in tokenGroups array.
  strfilter = "(|"
  If TypeName(arrbytGroups) = "Byte()" Then
    ' tokenGroups has one entry.
    strfilter = strfilter & "(objectSid=" _
      & OctetToHexStr(arrbytGroups) & ")"
  ElseIf UBound(arrbytGroups) > -1 Then
    ' TokenGroups is an array of two or more objectSid's.
    For k = 0 To UBound(arrbytGroups)
      strfilter = strfilter & "(objectSid=" & OctetToHexStr(arrbytGroups(k)) & ")"
    Next
  Else
    ' tokenGroups has no objectSid's.
    Exit Sub
  End If
  strfilter = strfilter & ")"

  ' Use ADO to search for groups whose objectSid matches any of the
  ' tokenGroups values for this user or computer.
  strQuery = strbase & ";" & strfilter & ";" _
    & strattributes & ";subtree"
  objCommand.CommandText = strQuery
  Set objRecordSet = objCommand.Execute

  ' Enumerate groups and add NT name to dictionary object.
  Do Until objRecordSet.EOF
    strGroupName = objRecordSet.Fields("sAMAccountName")
    'objGroupList(objADObject.sAMAccountName & "\" & strGroupName) = True
    objGroupList(strGroupName) = True
    UserGroups.Add strGroupName
    objRecordSet.MoveNext
  Loop

  Set objRecordSet = Nothing
End Sub

Function OctetToHexStr(arrbytOctet)
' Function to convert OctetString (byte array) to Hex string,
' with bytes delimited by \ for an ADO filter.

  Dim k
  OctetToHexStr = ""
  For k = 1 To LenB(arrbytOctet)
    OctetToHexStr = OctetToHexStr & "\" _
      & Right("0" & Hex(AscB(MidB(arrbytOctet, k, 1))), 2)
  Next
End Function





