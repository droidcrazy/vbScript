If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

' List DCOM Application Settings
Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./houston.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	listDCOMlogons server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 


Sub listDCOMlogons(strComputer)
On Error Resume Next

'strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_DCOMApplicationSetting")

For Each objItem in colItems
Select Case UCase(CStr(objItem.RunAsUser))
Case ""
Case "INTERACTIVE USER"
Case "NT AUTHORITY\LOCALSERVICE"
Case Else
    Wscript.Echo """" & strcomputer & """,""" & objItem.AppID & """,""" & objItem.Description & """,""" & objItem.RunAsUser & """"
End Select
Next
End Sub