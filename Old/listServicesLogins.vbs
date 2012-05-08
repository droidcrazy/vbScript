If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

' List Service logins
Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./houston.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile)
  	WScript.Echo """Server Name"",""Service Name"",""State"",""Logon Account"""
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	chkServices server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 

Sub chkServices(strComputer)
'strComputer = "."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
If Err.Number <> 0 Then WScript.Echo """" & strComputer & """,""Error binding to " & strComputer & """" : Err.Clear:Exit Sub
On Error Goto 0 
Set colListOfServices = objWMIService.ExecQuery _
        ("Select * from Win32_Service")

For Each objService in colListOfServices
Select Case UCase(objService.StartName)
Case "LOCALSYSTEM"
Case "NT AUTHORITY\NETWORKSERVICE"
Case "NT AUTHORITY\LOCALSERVICE"
Case "NT AUTHORITY\NETWORK SERVICE"
Case "NT AUTHORITY\LOCAL SERVICE"
Case Else
    wscript.echo("""" & objService.SystemName & """,""" & objService.Name & """,""" & objService.State & """,""" & objService.StartName & """")
End Select
Next

End Sub