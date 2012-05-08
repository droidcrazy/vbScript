' DocumentForest.vbs
' VBScript program to document an Active Directory forest. Program
' documents domains, containers, organizational units, and groups.
' Program also documents the number of user and computer objects in
' containers and groups, including users and computers that have the
' group as their "primary" group.
'
' ----------------------------------------------------------------------
' Copyright (c) 2003 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - April 19, 2003
' Version 1.1 - September 19, 2003 - Bug fix.
' Version 1.2 - July 6, 2007 - Modify use of Fields collection of
'                              Recordset object.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Const ADS_GROUP_TYPE_BUILTIN = &H1
Const ADS_GROUP_TYPE_GLOBAL = &H2
Const ADS_GROUP_TYPE_LOCAL = &H4
Const ADS_GROUP_TYPE_UNIVERSAL = &H8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &H80000000

Dim objRootDSE, strForest, objForest
Dim adoCommand, adoConnection

Set objRootDSE = GetObject("LDAP://RootDSE")
strForest = objRootDSE.Get("rootDomainNamingContext")
Set objForest = GetObject("LDAP://" & strForest)

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

Call EnumDomains(objForest, "")

' Clean up.
adoConnection.Close
Set objRootDSE = Nothing
Set objForest = Nothing
Set adoCommand = Nothing
Set adoConnection = Nothing

Sub EnumDomains(ByVal objParent, ByVal strOffset)
    ' Recursive subroutine to enumerate domains.
    Dim objGroup, objContainer, objChild

    ' Output domain name.
    Wscript.Echo strOffset & "Domain: " & objParent.name

    ' Count user and computer objects in domain.
    Call CountUsersComputers(objParent, "--" & strOffset)

    ' Enumerate groups in domain.
    objParent.Filter = Array("group")
    For Each objGroup In objParent
        Call EnumGroups(objGroup, objParent.distinguishedName, "--" & strOffset)
    Next

    ' Enumerate containers in domain.
    objParent.Filter = Array("container","organizationalUnit","builtinDomain")
    For Each objContainer In objParent
        Call EnumContainers(objContainer, objParent.distinguishedName, "--" & strOffset)
    Next

    ' Enumerate child domains.
    objParent.Filter = Array("domain")
    For Each objChild In objParent
        Call EnumDomains(objChild, "--" & strOffset)
    Next

    Set objGroup = Nothing
    Set objContainer = Nothing
    Set objChild = Nothing
End Sub

Sub EnumContainers(ByVal objParent, ByVal strDNSDomain, ByVal strOffset)
    ' Recursive subroutine to enumerate containers.
    Dim objGroup, objChild

    ' Output container name
    Wscript.Echo strOffset & "Container/OU: " & objParent.name

    ' Count user and computer objects in container.
    Call CountUsersComputers(objParent, "--" & strOffset)

    ' Enumerate groups in container.
    objParent.Filter = Array("group")
    For Each objGroup In objParent
        Call EnumGroups(objGroup, strDNSDomain, "--" & strOffset)
    Next

    ' Enumerate child containers.
    objParent.Filter = Array("container","organizationalUnit","builtinDomain")
    For Each objChild In objParent
        Call EnumContainers(objChild, strDNSDomain, "--" & strOffset)
    Next

    Set objGroup = Nothing
    Set objChild = Nothing
End Sub

Sub EnumGroups(ByVal objParent, ByVal strDNSDomain, ByVal strOffset)
    ' Recursive subroutine to enumerate groups.
    Dim lngUsers, lngComputers, lngGroups, objMember
    Dim lngPriUsers, lngPriComputers, intGroupToken
    Dim strFilter, strAttributes, strQuery, adoRecordset
    Dim strNTName, strCategory

    ' Output group sAMAccountName and type.
    Wscript.Echo strOffset & "Group: " & objParent.sAMAccountName _
        & " (" & GetType(objParent.GroupType) & ")"

    ' Enumerate group members.
    lngUsers = 0
    lngComputers = 0
    lngGroups = 0
    For Each objMember In objParent.Members
        If (LCase(objMember.Class) = "user") Then
            lngUsers = lngUsers + 1
        ElseIf (LCase(objMember.Class) = "computer") Then
            lngComputers = lngComputers + 1
        ElseIf (LCase(objMember.Class) = "group") Then
            lngGroups = lngGroups + 1
        End If
    Next
    If (lngUsers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of user members: " & lngUsers
    End If
    If (lngComputers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of computer members: " & lngComputers
    End If
    If (lngGroups > 0) Then
        Wscript.Echo "--" & strOffset & "Number of group members: " & lngGroups
    End If

    lngPriUsers = 0
    lngPriComputers = 0

    objParent.GetInfoEx Array("primaryGroupToken"), 0
    intGroupToken = objParent.Get("primaryGroupToken")
    strFilter = "(primaryGroupID=" & intGroupToken & ")"
    strAttributes = "sAMAccountName,objectCategory"
    strQuery = "<LDAP://" & strDNSDomain & ">;" & strFilter & ";" _
        & strAttributes & ";subtree"
    adoCommand.CommandText = strQuery
    Set adoRecordset = adoCommand.Execute
    Do Until adoRecordset.EOF
        strNTName = adoRecordset.Fields("sAMAccountName").Value
        strCategory = adoRecordset.Fields("objectCategory").Value
        If (Left(LCase(strCategory), 9) = "cn=person") Then
            lngPriUsers = lngPriUsers + 1
        End If
        If (Left(LCase(strCategory), 11) = "cn=computer") Then
            lngPriComputers = lngPriComputers + 1
        End If
        adoRecordset.MoveNext
    Loop
    adoRecordset.Close

    If (lngPriUsers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of user primary members: " & lngPriUsers
    End If

    If (lngPriComputers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of computer primary members: " & lngPriComputers
    End If

    ' Enumerate child groups.
    For Each objMember In objParent
        If (LCase(objMember.Class) = "group") Then
            Call EnumGroups(objMember, "--" & strOffset)
        End If
    Next

    Set objMember = Nothing
    Set adoRecordset = Nothing
End Sub

Sub CountUsersComputers(ByVal objContainer, ByVal strOffset)
    ' Subroutine to count computer objects in container.
    Dim lngUsers, lngComputers, objMember

    objContainer.Filter = Array("user")
    lngUsers = 0
    lngComputers = 0
    For Each objMember In objContainer
        If (LCase(objMember.Class) = "user") Then
            lngUsers = lngUsers + 1
        ElseIf (LCase(objMember.Class) = "computer") Then
            lngComputers = lngComputers + 1
        End If
    Next
    If (lngUsers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of users: " & lngUsers
    End If
    If (lngComputers > 0) Then
        Wscript.Echo "--" & strOffset & "Number of computers: " & lngComputers
    End If
    Set objMember = Nothing
End Sub

Function GetType(ByVal lngFlag)
    ' Function to determine group type.
    If ((lngFlag And ADS_GROUP_TYPE_BUILTIN) <> 0) Then
        GetType = "Built-in"
    ElseIf ((lngFlag And ADS_GROUP_TYPE_GLOBAL) <> 0) Then
        GetType = "Global"
    ElseIf ((lngFlag And ADS_GROUP_TYPE_LOCAL) <> 0) Then
        GetType = "Local"
    ElseIf ((lngFlag And ADS_GROUP_TYPE_UNIVERSAL) <> 0) Then
        GetType = "Universal"
    End If
    If ((lngFlag And ADS_GROUP_TYPE_SECURITY_ENABLED) <> 0) Then
        GetType = GetType & "/Security"
    Else
        GetType = GetType & "/Distribution"
    End If
End Function

