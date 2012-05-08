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

Dim CPUSink,ProcSink,objShell
On Error Resume Next

strCPUQuery="Select * from __InstanceCreationEvent WITHIN 10 WHERE " &_
"TargetInstance ISA 'Win32_Process'" 'AND " &_
'"TargetInstance.Name='IntegraSys_v28.exe'"

strCPU2Query="Select * from __InstanceDeletionEvent WITHIN 10 WHERE " &_
"TargetInstance ISA 'Win32_Process'" 'AND " &_
'"TargetInstance.Name='IntegraSys_v28.exe'"

'check every 10 seconds for existence of charmap.exe process
strProcQuery="Select * from __InstanceCreationEvent WITHIN 10 WHERE " &_
"TargetInstance ISA 'Win32_Process' AND " &_
"TargetInstance.Name='charmap.exe'"

Set objShell=CreateObject("Wscript.Shell")
Set CPUSink=WScript.CreateObject("WBemScripting.SWbemSink","CPUSINK_")
Set CPU2Sink=WScript.CreateObject("WBemScripting.SWbemSink","CPU2SINK_")
Set ProcSink=WScript.CreateObject("WBemScripting.SWbemSink","PROCSINK_")

Set oWMILocal=GetObject("winmgmts://")
oWMILocal.ExecNotificationQueryAsync ProcSink,strProcQuery
If err.number<>0 Then
 msgbox "Oops! There was an error creating process event sink " &_
 "locally." & vbCrlf & "Error #" &err.number & vbCrlf &_
 "Description (if available): " & vbCrlf & " " & err.description &_
  vbCrlf & "Source (If available): " & VbCrLf & " " &_
   err.source,-1,"CPU Monitoring"
 WScript.quit
Else
 Err.Clear
End If
blnLoop=True

Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./vproc.txt"
 If w Then OutputFile = "./processlog.txt"
 If w Then set txtStreamOut = fso.OpenTextFile(outputfile, 8, True)
 printout Now & " -- Monitoring started ******************************************************************************************************"
 

 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  If w Then
  MsgBox "Starting Process Monitor on systems in " & Right(inputfile, Len(inputfile) -2) & ".",0,"Process Monitor"
  MsgBox "Launch CHARMAP to stop monitoring",0,"Process Monitor"
  dq = MsgBox("Do you want to debug?",vbYesNo,"Debug?")
  ElseIf c Then
  printout "Starting Process Monitor on systems in " & Right(inputfile, Len(inputfile) -2) & "."
  printout "Launch CHARMAP to stop monitoring"
  End If
  Select Case dq
  Case 6
  d = True
  Case 7
  d = False
  Case Else
  d = False
  End Select
  
'	WScript.Sleep 10000
  Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	If blnLoop Then SetupRemoteMonitoring(strComputer)
  Loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 

'for debugging
' d = false


'Check if trigger process has been run, sleeping every 20 miliseconds.
While blnLoop
 WScript.sleep 20
Wend

WScript.Echo "Cancelling monitoring.  You can go ahead and close " &_
"the trigger application."
objShell.AppActivate("Character Map")

CPUSink.Cancel()
CPu2Sink.Cancel()
ProcSink.Cancel()
WScript.DisconnectObject(CPUSink)
WScript.DisconnectObject(CPU2Sink)
WScript.DisconnectObject(ProcSink)

Set oWMILocal=Nothing
Set oWMIRemote=Nothing
Set CPUSink=Nothing
Set ProcSink=Nothing
WScript.quit

'*******************************************************************
Sub CPUSINK_OnObjectReady(objEvent,objContext)
On Error Resume Next
strSystem=objEvent.Path_.Server
display = True
Select Case UCase(objEvent.TargetInstance.name)
Case "XCOPY.EXE"
display = True
Case "CMD.EXE"
display = False
Case "SLEEP.EXE"
display = false
Case Else
display = True
End Select
Call objEvent.TargetInstance.getowner(strUser,strDomain)
strProcess = objEvent.TargetInstance.name
dtStarted = WMIDateStringToDate(objEvent.TargetInstance.creationdate)
strCommandLine = objEvent.TargetInstance.commandline
strPID = objEvent.TargetInstance.ProcessId
If display Then printout Now & " - Process " & strProcess & "(" & strPID & ") started on " & strSystem & " by " & strDomain & "\" & strUser & " at " & dtStarted & ". Command Line: " & strCommandLine
strUser = ""
strDomain = ""
Err.Clear
End Sub

Sub CPU2SINK_OnObjectReady(objEvent,objContext)
On Error Resume Next
strSystem=objEvent.Path_.Server
display = True
Select Case UCase(objEvent.TargetInstance.name)
Case "XCOPY.EXE"
display = True
Case "CMD.EXE"
display = False
Case "SLEEP.EXE"
display = false
Case Else
display = True
End Select
Call objEvent.TargetInstance.getowner(strUser,strDomain)
strProcess = objEvent.TargetInstance.name
dtDeleted = WMIDateStringToDate(objEvent.TargetInstance.terminationdate)
strPID = objEvent.TargetInstance.ProcessId
If display Then printout Now & " - Process " & strProcess & "(" & strPID & ") ended on " & strSystem & " by " & strDomain & "\" & strUser & " at " & dtDeleted & "."
strDomain = ""
Err.Clear
End Sub

Sub PROCSINK_OnObjectReady(objEvent,objContext)
'trigger has been detected to close out this script
  blnLoop=False
End Sub

Sub SetupRemoteMonitoring(strComputer)
Err.Clear
Set oWMIRemote=GetObject("winmgmts://" & strComputer)
If err.number<>0 Then
 msgbox "Oops!  There was an error connecting to " &_
  UCase(strComputer) & vbCrlf & "Error #" &err.number & VbCrLf &_
  "Description (if available): " & VbCrLf & " " &_
  err.description & VbCrLf & "Source (If available): " & _
  vbCrlf & " " & Err.source,-1,"CPU Monitoring":Err.Clear
 Exit Sub
Else
 oWMIRemote.ExecNotificationQueryAsync CPUSink,strCPUQuery
 oWMIRemote.ExecNotificationQueryAsync CPU2Sink,strCPUQuery
  If err.number<>0 Then
   msgbox "Oops! There was an error creating CPU sink for " &_
    UCase(strComputer) & vbCrlf & "Error #" &err.number & vbCrlf &_
     "Description (if available): " & vbCrlf & " " &_
      err.description & vbCrlf & "Source (If available): " & _
   vbCrlf & " " & err.source,-1,"CPU Monitoring":Err.Clear
   Exit Sub
  Else
   If d Then printout "Monitoring: " & strComputer
   err.Clear
  End If
End If
End Sub

Function printout(msg)
If w Then txtStreamOut.WriteLine msg
If c Then WScript.Echo msg
End Function

Function WMIDateStringToDate(dtmWMI)
 WMIDateStringToDate = CDate(Mid(dtmWMI, 5, 2) & "/" & _
 Mid(dtmWMI, 7, 2) & "/" & Left(dtmWMI, 4) _
 & " " & Mid (dtmWMI, 9, 2) & ":" & _
 Mid(dtmWMI, 11, 2) & ":" & Mid(dtmWMI, _
 13, 2))
End Function

'EOF
