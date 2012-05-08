If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

ContainerName = "houston"
Set Container = getobject("WinNT://" & ContainerName)
Container.Filter = Array("Computer")
For Each Computer In Container
if wscript.Arguments.Named.Exists("user") Then strUser = WScript.Arguments.Named("user") Else strUser = ""
Call getSessions(ContainerName,computer.name,strUser)
Next
wscript.Echo("Done.")

Sub getSessions(strDomain,strComputer,strUser)
	On Error Resume Next
	Set container = getobject("WinNT://" & strDomain & "/" & strComputer & "/LanmanServer")
	If Err.Number <> 0 Then Err.Clear : Exit Sub
	For Each sessions In container.sessions
		If strUser = "" Then
		wscript.echo strComputer & "," & sessions.user & "," & sessions.computer & "," & int(sessions.connecttime/60) & "," & int(sessions.idletime/60)
		Else
		If ucase(sessions.user) = UCase(strUser) Then
		WScript.Echo strComputer & "," & sessions.user & "," & sessions.computer & "," & int(sessions.connecttime/60) & "," & int(sessions.idletime/60)
		End If
		End If 
	Next
End Sub 'getSessions