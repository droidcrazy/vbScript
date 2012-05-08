Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "H:/vprocs.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine
	process = "IntegraSys_v28.exe"
	checkProc process,server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 
 
Sub checkProc(strProcess,strComputer)
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & strProcess & "'")
	For Each objProcess in colProcessList
    	WScript.Echo(objProcess.name & " is on " & strComputer)
Next
End Sub