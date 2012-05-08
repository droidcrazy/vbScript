const HKEY_LOCAL_MACHINE = &H80000002
const REG_SZ = 1
const REG_EXPAND_SZ = 2
const REG_BINARY = 3
const REG_DWORD = 4
const REG_MULTI_SZ = 7
 
strComputer = "."
Set StdOut = WScript.StdOut
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
 
strKeyPath = "software\policies\microsoft\windows\windowsupdate"


strValueName = "TargetGroup"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
StdOut.WriteLine "Target Group(s): " & strValue

