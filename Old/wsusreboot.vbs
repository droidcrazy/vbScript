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
 
strKeyPath = "software\microsoft\windows\currentversion\windowsupdate\auto update"
 
oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath,_
 arrValueNames, arrValueTypes
 
For i=0 To UBound(arrValueNames)
    StdOut.WriteLine "Value Name: " & arrValueNames(i) 
    
    Select Case arrValueTypes(i)
        Case REG_SZ
            StdOut.WriteLine "Data Type: String"
            StdOut.WriteBlankLines(1)
        Case REG_EXPAND_SZ
            StdOut.WriteLine "Data Type: Expanded String"
            StdOut.WriteBlankLines(1)
        Case REG_BINARY
            StdOut.WriteLine "Data Type: Binary"
            StdOut.WriteBlankLines(1)
        Case REG_DWORD
            StdOut.WriteLine "Data Type: DWORD"
            StdOut.WriteBlankLines(1)
        Case REG_MULTI_SZ
            StdOut.WriteLine "Data Type: Multi String"
            StdOut.WriteBlankLines(1)
    End Select 
Next