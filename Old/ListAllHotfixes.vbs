' List Installed Hot Fixes

Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./houston.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	gethotfixes server
  Loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If
Set fso = Nothing 

Sub gethotfixes(strComputer)
'strComputer = "."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
If Err.Number <> 0 Then WScript.Echo("Error " & Err.Number & " occured connecting to " & strComputer & ". Description: " & Err.Description):Err.Clear:Exit Sub
On Error Goto 0

Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")

For Each objQuickFix in colQuickFixes
If objQuickFix.HotFixID <> "File 1" Then
	strID = objQuickFix.HotFixID
	dateInstallDate = objQuickFix.InstalledOn
    Wscript.Echo "Computer: " & objQuickFix.CSName & ", Hot Fix ID: " & objQuickFix.HotFixID & ", Installation Date: " & objQuickFix.InstalledOn
    'Wscript.Echo "Description: " & objQuickFix.Description 
    'Wscript.Echo "Hot Fix ID: " & objQuickFix.HotFixID
    'Wscript.Echo "Installation Date: " & objQuickFix.InstallDate
    'Wscript.Echo "Installed By: " & objQuickFix.InstalledBy
End If
Next
End Sub