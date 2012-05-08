On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalShareSecuritySetting",,48)
For Each objItem in colItems

strShareName = objItem.name

Set wmiFileSecSetting = GetObject("winmgmts:Win32_LogicalShareSecuritySetting.name='" & strShareName & "'")

RetVal = wmiFileSecSetting.GetSecurityDescriptor(wmiSecurityDescriptor)
If Err <> 0 Then
WScript.Echo "GetSecurityDescriptor failed on " & strShareName & vbCrLf & Err.Number & VBCRLF & Err.Description
Else
WScript.Echo "GetSecurityDescriptor suceeded for " & strShareName
End If

' Retrieve the DACL array of Win32_ACE objects.
DACL = wmiSecurityDescriptor.DACL

For each wmiAce in DACL

WScript.Echo "-----------------------"
WScript.Echo "Found ACE"
WScript.Echo "-----------------------"
wscript.echo "Access Mask: " & wmiAce.AccessMask
wscript.echo "ACE Type: " & wmiAce.AceType

' Get Win32_Trustee object from ACE
Set Trustee = wmiAce.Trustee
wscript.echo "Trustee Domain: " & Trustee.Domain
wscript.echo "Trustee Name: " & Trustee.Name

' Get SID as array from Trustee
SID = Trustee.SID

For i = 0 To UBound(SID) - 1
strsid = strsid & SID(i) & ","
Next
strsid = strsid & SID(i)
wscript.echo "Trustee SID: {" & strsid & "}"

Next
WScript.Echo "==========================================================================="
Next