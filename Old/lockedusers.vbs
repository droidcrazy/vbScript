If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

' LockedUsers.vbs
' VBScript program to find user accounts in Active Directory that are
' locked out, then determine when they were locked out and on which
' Domain Controller.
'
' ----------------------------------------------------------------------
' Copyright (c) 2003 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - March 17, 2003
' Version 1.1 - May 9, 2003 - Account for error in IADsLargeInteger
'                             property methods HighPart and Lowpart.
' Version 1.2 - January 25, 2004 - Modify error trapping.
' Version 1.3 - March 18, 2004 - Modify NameTranslate constants.
' Version 1.4 - July 6, 2007 - Modify how IADsLargeInteger interface
'                              is invoked.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

'Option Explicit

Dim objRootDSE, strConfig, adoConnection, adoCommand, strQuery
Dim adoRecordset, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs()
Dim strDN, dtmDate, objDate, strUser, strNTName
Dim objList1, objList2, objList3, j, intBadCount
Dim strBase, strFilter, strAttributes, objWinNTUser
Dim objTrans, strNetBIOSDomain, objDomain, arrstrNTNames()
Dim lngHigh, lngLow

' Constants for the NameTranslate object.
Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use the NameTranslate object to convert the DNS domain name
' to the NetBIOS domain name.
Set objTrans = CreateObject("NameTranslate")
objTrans.Init ADS_NAME_INITTYPE_GC, ""
objTrans.Set ADS_NAME_TYPE_1779, strDNSDomain
strNetBIOSDomain = objTrans.Get(ADS_NAME_TYPE_NT4)
' Remove trailing backslash.
strNetBIOSDomain = Left(strNetBIOSDomain, Len(strNetBIOSDomain) - 1)

' Find locked out user accounts in domain
' create array of sAMAccountName's
Set objDomain = GetObject("WinNT://" & strNetBIOSDomain)
objDomain.Filter = Array("user")
k = 0
For Each objWinNTUser In objDomain
    If (objWinNTUser.IsAccountLocked = True) Then
        ReDim Preserve arrstrNTNames(k)
        arrstrNTNames(k) = objWinNTUser.name
        k = k + 1
    End If
Next

If (k = 0) Then
    Wscript.Echo "No user accounts locked out in domain"
    Wscript.Quit
End If

' Use dictionary objects to track latest badPasswordTime,
' badPwdCount, and Domain Controller for each locked out user.
Set objList1 = CreateObject("Scripting.Dictionary")
objList1.CompareMode = vbTextCompare
Set objList2 = CreateObject("Scripting.Dictionary")
objList2.CompareMode = vbTextCompare
Set objList3 = CreateObject("Scripting.Dictionary")
objList3.CompareMode = vbTextCompare

' Obtain local Time Zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If

' Determine configuration context.
strConfig = objRootDSE.Get("configurationNamingContext")

' Use ADO to search Active Directory for ObjectClass nTDSDSA.
' This will identify all Domain Controllers.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open = "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

strBase = "<LDAP://" & strConfig & ">"
strFilter = "(objectClass=nTDSDSA)"
strAttributes = "AdsPath"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 60
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

' Enumerate parent objects of class nTDSDSA. Save Domain Controller
' DNS host names in dynamic array arrstrDCs.
k = 0
Do Until adoRecordset.EOF
    Set objDC = _
        GetObject(GetObject(adoRecordset.Fields("AdsPath").Value).Parent)
    ReDim Preserve arrstrDCs(k)
    arrstrDCs(k) = objDC.DNSHostName
    k = k + 1
    adoRecordset.MoveNext
Loop
adoRecordset.Close

' Use ADO to retrieve all user objects from each Domain Controller.
strFilter = "(&(objectCategory=person)(objectClass=user))"
strAttributes = "distinguishedName,sAMAccountName," _
    & "badPasswordTime,badPwdCount"
For k = 0 To Ubound(arrstrDCs)
	WScript.Echo "Querying " & arrstrDCs(k)
    strBase = "<LDAP://" & arrstrDCs(k) & "/" & strDNSDomain & ">"
    strQuery = strBase & ";" & strFilter & ";" & strAttributes _
        & ";subtree"
    adoCommand.CommandText = strQuery
    On Error Resume Next
    Set adoRecordset = adoCommand.Execute
    If (Err.Number <> 0) Then
        On Error GoTo 0
        Wscript.Echo "Domain Controller not available: " & arrstrDCs(k)
    Else
        On Error GoTo 0
        Do Until adoRecordset.EOF
            strNTName = adoRecordset.Fields("sAMAccountName").Value
            ' Check each user to see if in array of locked out accounts.
            For j = 0 To UBound(arrstrNTNames)
                If (UCase(strNTName) = UCase(arrstrNTNames(j))) Then
                    ' User locked out, retrieve badPasswordTime.
                    strDN = adoRecordset.Fields("distinguishedName").Value
                    intBadCount = adoRecordset.Fields("badPwdCount").Value
                    On Error Resume Next
                    Set objDate = adoRecordset.Fields("badPasswordTime").Value
                    If (Err.Number <> 0) Then
                        On Error GoTo 0
                        dtmDate = #1/1/1601#
                    Else
                        On Error GoTo 0
                        lngHigh = objDate.HighPart
                        lngLow = objDate.LowPart
                        If (lngLow < 0) Then
                            lngHigh = lngHigh + 1
                        End If
                        If (lngHigh = 0) And (lngLow = 0 ) Then
                            dtmDate = #1/1/1601#
                        Else
                            dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
                                + lngLow)/600000000 - lngBias)/1440
                        End If
                    End If
                    If (objList1.Exists(strDN) = True) Then
                        If (dtmDate > objList1.Item(strDN)) Then
                            ' Later badBadPasswordTime found, save info from this DC.
                            objList1.Item(strDN) = dtmDate
                            objList2.Item(strDN) = intBadCount
                            objList3.Item(strDN) = arrstrDCs(k)
                        End If
                    Else
                        ' First time user found, save info from this DC.
                        objList1.Add strDN, dtmDate
                        objList2.Add strDN, intBadCount
                        objList3.Add strDN, arrstrDCs(k)
                    End If
                End If
            Next
            adoRecordset.MoveNext
        Loop
        adoRecordset.Close
    End If
Next

' Output information on each locked out user.
For Each strUser In objList1.Keys
    Wscript.Echo strUser & " ; " & objList1.Item(strUser) & " ; " _
        & objList2.Item(strUser) & " ; " & objList3.Item(strUser)
Next

' Clean up.
adoConnection.Close
Set objRootDSE = Nothing
Set adoConnection = Nothing
Set adoCommand = Nothing
Set adoRecordset = Nothing
Set objTrans = Nothing
Set objDomain = Nothing
Set objWinNTUser = Nothing
Set objDC = Nothing
Set objDate = Nothing
Set objList1 = Nothing
Set objList2 = Nothing
Set objList3 = Nothing
Set objShell = Nothing

