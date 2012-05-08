const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\System Admin Scripting Guide"
strValueName = "String Value Name"
strValue = "string value"
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue