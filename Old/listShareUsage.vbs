Dim objStdOut, args
Dim strComputer, strDomain
Dim strSeconds, blnRepeat
Dim counter

Set objStdOut = Wscript.stdOut
Set args = Wscript.Arguments.Named

strComputer = args.Item("computer")
strDomain = args.Item("adcon")
If args.Exists("timer") Then strSeconds = args.Item("timer"):blnRepeat = True
	On Error Resume Next
	Set container = getobject("WinNT://" & strDomain & "/" & strComputer & "/LanmanServer")

Call main()

Sub main
If blnRepeat Then
Call repeater(strSeconds)
Else
Call list(strComputer, strDomain)
End If
End Sub

Sub list(strComputer, strDomain)
 On Error Resume Next
	strOutput = strOutput & "Accessed By" & vbTab & "# Locks" & vbtab & "Open File" & vbCrLf
	strOutput = strOutput & "===========" & vbTab & "=======" & vbtab & "=========" & vbCrLf
counter = 0
	for each resource in container.resources
	counter = counter + 1
'		wscript.echo resource.user & vbtab & resource.lockcount & vbtab & resource.path
		strOutput = strOutput & resource.user
		strOutput = strOutput & vbTab
		If Len(resource.user) < 8 Then strOutput = strOutput & vbTab
		strOutput = strOutput & vbTab
		strOutput = strOutput & resource.lockcount
		strOutput = strOutput & vbTab
		strOutput = strOutput & resource.path
		If Err.Number <> 0 Then strOutput = strOutput & "error 0x" & Hex(Err.Number) & ":" & Err.Description:Err.Clear
		strOutput = strOutput & vbCrLf
	Next
	objStdOut.Write strOutput
	objStdOut.WriteLine "Total Number: " & counter
	Set container = Nothing
End Sub

Sub repeater(strSeconds)
On Error Resume Next
Do While blnRepeat
counter = 0
for each resource in container.resources
counter = counter + 1
Next
objStdOut.WriteLine Now & " -- Count: " & counter
WScript.Sleep strSeconds * 1000
Loop
End Sub