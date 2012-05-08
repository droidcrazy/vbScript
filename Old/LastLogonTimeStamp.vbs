If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

'region Script Settings
'<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
'  <ScriptPackager>
'    <process>cscript.exe</process>
'    <arguments />
'    <extractdir>%TEMP%</extractdir>
'    <files />
'    <usedefaulticon>true</usedefaulticon>
'    <showinsystray>false</showinsystray>
'    <altcreds>false</altcreds>
'    <efs>true</efs>
'    <ntfs>true</ntfs>
'    <local>false</local>
'    <abortonfail>true</abortonfail>
'    <product />
'    <version>1.0.0.1</version>
'    <versionstring />
'    <comments />
'    <includeinterpreter>false</includeinterpreter>
'    <forcecomregistration>false</forcecomregistration>
'    <consolemode>true</consolemode>
'    <EnableChangelog>false</EnableChangelog>
'    <AutoBackup>false</AutoBackup>
'    <snapinforce>false</snapinforce>
'    <snapinshowprogress>false</snapinshowprogress>
'    <snapinautoadd>0</snapinautoadd>
'    <snapinpermanentpath />
'  </ScriptPackager>
'</ScriptSettings>
'endregion

' LastLogonTimeStamp.vbs
' VBScript program to determine when each user in the domain last logged
' on. Domain must be at Windows Server 2003 Functional Level.
'
' ----------------------------------------------------------------------
' Copyright (c) 2007 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - March 24, 2007
' Version 1.1 - July 6, 2007 - Modify how IADsLargeInteger interface
'                              is invoked.
'
' The lastLogonTimeStamp attribute is Integer8, a 64-bit number
' representing the date as the number of 100 nanosecond intervals since
' 12:00 am January 1, 1601. This value is converted to a date. The last
' logon date is in UTC (Coordinated Univeral Time). It must be adjusted
' by the Time Zone bias in the machine registry to convert to local
' time.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objRootDSE, adoConnection, adoCommand, strQuery
Dim adoRecordset, strDNSDomain, objShell, lngBiasKey
Dim lngBias, k, strDN, dtmDate, objDate, objCreateDate, dtmCreateDate
Dim strBase, strFilter, strAttributes, lngHigh, lngLow, PropFlag
' Obtain local Time Zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
& "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
	lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
	lngBias = 0
	For k = 0 To UBound(lngBiasKey)
		lngBias = lngBias + (lngBiasKey(k) * 256 ^ k)
	Next
End If
Set objShell = Nothing

' Determine DNS domain from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Set objRootDSE = Nothing

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on all user objects.
strFilter = "(&(objectCategory=person)(objectClass=user))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "Name, distinguishedname, cn, lastLogonTimeStamp, whenCreated, userAccountControl"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Run the query.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 60
adoCommand.Properties("Cache Results") = False
Set adoRecordset = adoCommand.Execute

' Enumerate resulting recordset.
Do Until adoRecordset.EOF
	' Retrieve attribute values for the user.
	strDN = adoRecordset.Fields("Name").Value
	' Convert Integer8 value to date/time in current time zone.
	On Error Resume Next
	Set objDate = adoRecordset.Fields("lastLogonTimeStamp").Value
	If (Err.Number <> 0) Then
		On Error GoTo 0
		dtmDate = # 1 / 1 / 1601 #
	Else
		On Error GoTo 0
		lngHigh = objDate.HighPart
		lngLow = objDate.LowPart
		If (lngLow < 0) Then lngHigh = lngHigh + 1
		If (lngHigh = 0) And ( lngLow = 0 ) Then dtmDate = # 1 / 1 / 1601 # Else dtmDate = # 1 / 1 / 1601 # + (((lngHigh * (2 ^ 32)) + lngLow) / 600000000 - lngBias) / 1440
	End If
	On Error Resume Next
	dtmCreateDate = adoRecordset.Fields("whenCreated").Value
	    
	If CLng(adoRecordset.Fields("userAccountControl").Value) And 2 Then PropFlag = PropFlag & " Account is disabled"
	If CLng(adoRecordset.Fields("userAccountControl").Value) And 16 Then PropFlag = PropFlag & " Account is locked out"
	If PropFlag = "" Then PropFlag = "Account is Active"
	    
	    
	' Display values for the user.
	If (dtmDate = # 1 / 1 / 1601 #) Then
		WScript.Echo strDN & ";" & adoRecordset.Fields("distinguishedName").Value & ";" & adoRecordset.Fields("cn").Value & ";Never;" & dtmCreateDate & ";" & PropFlag
	Else
		WScript.Echo strDN & ";" & adoRecordset.Fields("distinguishedName").Value & ";" & adoRecordset.Fields("cn").Value & ";" & dtmDate & ";" & dtmCreateDate & ";" & PropFlag
	End If
	propflag = ""
	adoRecordset.MoveNext
Loop

' Clean up.
adoRecordset.Close
adoConnection.Close
Set adoConnection = Nothing
Set adoCommand = Nothing
Set adoRecordset = Nothing
Set objDate = Nothing