' NAME: ReadRegistry.vbs
' VERSION: 1.0 2/7/2008
' AUTHOR: Jeffery Hicks jhicks@sapien.com
' USAGE: cscript ReadRegistry [/s:computername] 

' DESCRIPTION: Use the WMI registry provider to read a remote
' registry key value. If you don’t specify a computer name the script
' will default to localhost. 

' *********************************************************************************
' * THIS PROGRAM IS OFFERED AS IS AND MAY BE FREELY MODIFIED OR ALTERED AS *
' * NECESSARY TO MEET YOUR NEEDS. THE AUTHOR MAKES NO GUARANTEES OR WARRANTIES, *
' * EXPRESS, IMPLIED OR OF ANY OTHER KIND TO THIS CODE OR ANY USER MODIFICATIONS. *
' * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED IN A SECURED LAB *
' * ENVIRONMENT. USE AT YOUR OWN RISK. *
' ********************************************************************************* 

On Error Resume Next 

Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7 

'the registry path to query
strKeyPath ="Software\Microsoft\Windows NT\CurrentVersion\"
'the registry key value to get
strValueName="RegisteredOwner" 

If WScript.Arguments.Named.Exists("S") Then
strComputer=UCase(WScript.Arguments.Named.Item("S"))
Else
strComputer = "LOCALHOST"
End If 

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\" &_
strComputer & "\root\default:StdRegProv")
If Err.Number = 0 Then 
strMsg="There was a problem connecting to " &_
strComputer & "\root\default:StdRegProv" & VbCrLf &_
"Error " & Err.Number & " " & Err.description
WScript.Echo strMsg
WScript.Quit
End If
objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValues, arrTypes
For x=0 To UBound(arrValues)-1
if Ucase(arrValues(x)) = UCase(strValueName) Then 

Select Case arrTypes(x)
Case REG_SZ
objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Case REG_EXPAND_SZ
objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Case REG_BINARY
objReg.GetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Case REG_DWORD
objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Wscript.Echo
Case REG_MULTI_SZ
objReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
End Select 
End If 

Next 

If VarType(strValue)=0 Then
WScript.Echo "Failed to get a value for " & strKeyPath & strValueName &_
" on " & strComputer & ". Check your registry keys and permissions."
Else
If IsArray (strValue) Then
WScript.Echo strComputer & " - " & strValueName
For Each value In strValue
WScript.Echo value
Next
Else
WScript.Echo strComputer & " - " & strValueName & " = " & strValue
End If
End If 

'EOF
