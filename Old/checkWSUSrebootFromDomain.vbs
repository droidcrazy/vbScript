Select Case Right(UCase(WScript.FullName), Len("CScript.exe"))
Case UCase("CScript.exe")
c = True
w = False
Case UCase("WScript.exe")
c = False
w = True
Case Else
WScript.Quit
End Select

'For debugging.
d = False

'If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
'    Set objShell = CreateObject("wscript.shell")
'    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
'    WScript.Quit
'End If

If w Then OutputFile = "./needsreboot.txt" : Set fso = CreateObject("Scripting.FileSystemObject") : Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True)

ContainerName = "houston"
Set Container = getobject("WinNT://" & ContainerName)
Container.Filter = Array("Computer")
For Each Computer In Container
Call chkreboot(computer.name)
Next
printout "Done."

Sub chkreboot(strComputer)
If c And d Then printOut "Checking " & strComputer & "..."
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002
err.clear 
	Set colgroups = GetObject("WinNT://" & strComputer & ",computer") 
	colGroups.GetInfo 
	If err.number <> 0 then 
	If c And d Then printout "Error binding to computer " & strComputer & "."
	Else  
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "software\microsoft\windows\currentversion\windowsupdate\auto update"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
For Each subkey In arrSubKeys
If subkey = "RebootRequired" Then
	printout strcomputer & " needs to be rebooted."
 End If
 Next
End If 
End Sub 'chkreboot

Sub printOut (data) 
If c Then WScript.Echo data
If w Then txtStreamOut.writeline data 
End Sub 'printOut 