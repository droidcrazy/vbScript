' DocumentUsers.vbs
' VBScript program to document all users in Active Directory. Can be
' used to create a comma delimited file that can be read into a
' spreadsheet program.
'
' ----------------------------------------------------------------------
' Copyright (c) 2007 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - August 6, 2007
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objRootDSE, strDNSDomain, adoCommand, adoConnection
Dim strBase, strFilter, strAttributes, strQuery, adoRecordset
Dim strDN, strNTName, strFirst, strLast, arrDesc
Dim strItem, strDesc, strPhone, arrOtherPhone, strOtherPhone
Dim lngFlags, strFlags, objPwdLastSet, dtmPwdLastSet
Dim objShell, lngBiasKey, lngTZBias, k, arrAttrValues

' Obtain local Time Zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngTZBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngTZBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngTZBias = lngTZBias + (lngBiasKey(k) * 256^k)
    Next
End If
Set objShell = Nothing

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Search for all users.
strFilter = "(&(objectCategory=person)(objectClass=user))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName,sAMAccountName,givenName,sn," _
    & "description,userAccountControl,pwdLastSet,telephoneNumber," _
    & "otherTelephone"

' Construct the LDAP query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Run the query.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False
Set adoRecordset = adoCommand.Execute

' Output heading line.
Wscript.Echo """Distinguished Name"",""NT Name"",""First Name""," _
    & """Last Name"",""Description"",""Flags"",""Password Last Set""," _
    & """Telephone Number"",""Other Telephone Numbers"""

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
    ' Retrieve single-valued strings.
    strDN = adoRecordset.Fields("distinguishedName").Value
    ' Escape any forward slash characters, "/", with the backslash
    ' escape character. All other characters that should be escaped are.
    strDN = Replace(strDN, "/", "\/")

    strNTName = adoRecordset.Fields("sAMAccountName").Value
    strFirst = adoRecordset.Fields("givenName").Value
    strLast = adoRecordset.Fields("sn").Value
    strPhone = adoRecordset.Fields("telephoneNumber").Value

    ' The description attribute is multi-valued, but
    ' there is never more than one item in the array.
    arrDesc = adoRecordset.Fields("description").Value
    If IsNull(arrDesc) Then
        strDesc = ""
    Else
        For Each strItem In arrDesc
            strDesc = strItem
        Next
    End If

    ' The otherTelephone attribute is multi-valued.
    ' List numbers delimited by semicolons.
    arrOtherPhone = adoRecordset.Fields("otherTelephone").Value
    If IsNull(arrOtherPhone) Then
        strOtherPhone = ""
    Else
        strOtherPhone = ""
        For Each strItem In arrOtherPhone
            If (strOtherPhone = "") Then
                strOtherPhone = strItem
            Else
                strOtherPhone = strOtherPhone & ";" & strItem
            End If
        Next
    End If

    ' Test bits of userAccountControl.
    lngFlags = CLng(adoRecordset.Fields("userAccountControl").Value)
    strFlags = GetFlags(lngFlags)

    ' Convert Integer8 value to date in current time zone.
    Set objPwdLastSet = adoRecordset.Fields("pwdLastSet").Value
    dtmPwdLastSet = Integer8Date(objPwdLastSet, lngTZBias)

    ' Create array of string values to display.
    arrAttrValues = Array(strDN, strNTName, strFirst, strLast, _
        strDesc, strFlags, CStr(dtmPwdLastSet), strPhone, _
        strOtherPhone)

    ' Display array of values in a comma delimited line, with each
    ' value enclosed in quotes.
    Wscript.Echo CSVLine(arrAttrValues)

    ' Move to next record in recordset.
    adoRecordset.MoveNext
Loop

' Clean up.
adoRecordset.Close
adoConnection.Close

Function GetFlags(ByVal lngFlag)
    ' Function to test bits of userAccountControl attribute.
    ' Settings delimited by semicolons.

    ' Define bit masks.
    Const ADS_UF_ACCOUNTDISABLE = &H02
    Const ADS_UF_HOMEDIR_REQUIRED = &H08
    Const ADS_UF_LOCKOUT = &H10
    Const ADS_UF_PASSWD_NOTREQD = &H20
    Const ADS_UF_PASSWD_CANT_CHANGE = &H40
    Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80
    Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H100
    Const ADS_UF_NORMAL_ACCOUNT = &H200
    Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
    Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H1000
    Const ADS_UF_SERVER_TRUST_ACCOUNT = &H2000
    Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
    Const ADS_UF_MNS_LOGON_ACCOUNT = &H20000
    Const ADS_UF_SMARTCARD_REQUIRED = &H40000
    Const ADS_UF_TRUSTED_FOR_DELEGATION = &H80000
    Const ADS_UF_NOT_DELEGATED = &H100000
    Const ADS_UF_USE_DES_KEY_ONLY = &H200000
    Const ADS_UF_DONT_REQUIRE_PREAUTH = &H400000
    Const ADS_UF_PASSWORD_EXPIRED = &H800000
    Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = &H1000000

    GetFlags = ""

    If (lngFlag And ADS_UF_ACCOUNTDISABLE) <> 0 Then
        GetFlags = GetFlags & ";" & "User account disabled"
    End If
    If (lngFlag And ADS_UF_HOMEDIR_REQUIRED) <> 0 Then
        GetFlags = GetFlags & ";" & "Home directory required"
    End If
    If (lngFlag And ADS_UF_LOCKOUT) <> 0 Then
        GetFlags = GetFlags & ";" & "Account currently locked out"
    End If
    If (lngFlag And ADS_UF_PASSWD_NOTREQD) <> 0 Then
        GetFlags = GetFlags & ";" & "No password required"
    End If
    If (lngFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0 Then
        GetFlags = GetFlags & ";" & "User cannot change password"
    End If
    If (lngFlag And ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED) <> 0 Then
        GetFlags = GetFlags & ";" & "User can send an encrypted password"
    End If
    If (lngFlag And ADS_UF_TEMP_DUPLICATE_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "Account for user in another domain (local user account)"
    End If
    If (lngFlag And ADS_UF_NORMAL_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "Default account for typical user"
    End If
    If (lngFlag And ADS_UF_INTERDOMAIN_TRUST_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "A ""permit to trust"" account for a domain that ""trusts"" other domains"
    End If
    If (lngFlag And ADS_UF_WORKSTATION_TRUST_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "Computer account"
    End If
    If (lngFlag And ADS_UF_SERVER_TRUST_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "Computer account for system backup domain controller"
    End If
    If (lngFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0 Then
        GetFlags = GetFlags & ";" & "Password does not expire"
    End If
    If (lngFlag And ADS_UF_MNS_LOGON_ACCOUNT) <> 0 Then
        GetFlags = GetFlags & ";" & "MNS logon account"
    End If
    If (lngFlag And ADS_UF_SMARTCARD_REQUIRED) <> 0 Then
        GetFlags = GetFlags & ";" & "User must logon using a smart card"
    End If
    If (lngFlag And ADS_UF_TRUSTED_FOR_DELEGATION) <> 0 Then
        GetFlags = GetFlags & ";" & "Service account under which a service runs, trusted for Kerberos"
    End If
    If (lngFlag And ADS_UF_NOT_DELEGATED) <> 0 Then
        GetFlags = GetFlags & ";" & "Security context will not be delegated to a service"
    End If
    If (lngFlag And ADS_UF_USE_DES_KEY_ONLY) <> 0 Then
        GetFlags = GetFlags & ";" & "Must use DES encryption types for keys"
    End If
    If (lngFlag And ADS_UF_DONT_REQUIRE_PREAUTH) <> 0 Then
        GetFlags = GetFlags & ";" & "Account does not require Kerberos preauthenication for logon"
    End If
    If (lngFlag And ADS_UF_PASSWORD_EXPIRED) <> 0 Then
        GetFlags = GetFlags & ";" & "User password has expired"
    End If
    If (lngFlag And ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) <> 0 Then
        GetFlags = GetFlags & ";" & "Account enabled for delegation"
    End If

    If (Len(GetFlags) > 1) Then
        GetFlags = Mid(GetFlags, 2)
    End If

End Function

Function Integer8Date(ByVal objDate, ByVal lngBias)
    ' Function to convert Integer8 (64-bit) value to a date, adjusted for
    ' local time zone bias.
    Dim lngAdjust, lngDate, lngHigh, lngLow
    lngAdjust = lngBias
    lngHigh = objDate.HighPart
    lngLow = objdate.LowPart
    ' Account for error in IADslargeInteger property methods.
    If (lngLow < 0) Then
        lngHigh = lngHigh + 1
    End If
    If (lngHigh = 0) And (lngLow = 0) Then
        lngAdjust = 0
    End If
    lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
        + lngLow) / 600000000 - lngAdjust) / 1440
    ' Trap error if lngDate is ridiculously huge.
    On Error Resume Next
    Integer8Date = CDate(lngDate)
    If (Err.Number <> 0) Then
        On Error GoTo 0
        Integer8Date = #1/1/1601#
    End If
    On Error GoTo 0
End Function

Function CSVLine(ByVal arrValues)
    ' Function to convert array of values into comma delimited
    ' values enclosed in quotes.
    Dim strItem

    CSVLine = ""
    For Each strItem In arrValues
        ' Replace any embedded quotes with two quotes.
        If (strItem <> "") Then
            strItem = Replace(strItem, """", """" & """")
        End If
        ' Append string values, enclosed in quotes,
        ' delimited by commas.
        If (CSVLine = "") Then
            CSVLine = """" & strItem & """"
        Else
            CSVLine = CSVLine & ",""" & strItem & """"
        End If
    Next

End Function

