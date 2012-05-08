Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "pxhousdc02"

Set objRegProv = GetObject("winmgmts:{impersonationLevel=Impersonate}" & _
 "!\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows Script Host\Settings"
objRegProv.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"Remote","1"

