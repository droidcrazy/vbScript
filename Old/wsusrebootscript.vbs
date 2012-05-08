On Error Resume Next 
Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./houston.txt"
 OutputFile = "./needsreboot.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	server = cleanme(server)
	chkreboot server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 
 
Sub chkreboot(strComputer)
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002
err.clear 
	Set colgroups = GetObject("WinNT://" & strComputer & ",computer") 
	colGroups.GetInfo 
	If err.number <> 0 then 
		printout "Error binding to computer " & strComputer & "."
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
'	WScript.Echo data
	txtStreamOut.writeline data 
End Sub 'printOut 

Function cleanme(str)
Dim regEx
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.Pattern = "\s"
str = Trim(regEx.Replace(str, ""))
return str 
End Function 'cleanme