'region Script Settings
'<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
'  <ScriptPackager>
'    <process>cscript.exe</process>
'    <arguments />
'    <extractdir>%TEMP%</extractdir>
'    <files />
'    <usedefaulticon>true</usedefaulticon>
'    <showinsystray>false</showinsystray>
'    <altcreds>false</altcreds>
'    <efs>true</efs>
'    <ntfs>true</ntfs>
'    <local>false</local>
'    <abortonfail>true</abortonfail>
'    <product />
'    <version>1.0.0.1</version>
'    <versionstring />
'    <comments />
'    <includeinterpreter>false</includeinterpreter>
'    <forcecomregistration>false</forcecomregistration>
'    <consolemode>false</consolemode>
'    <EnableChangelog>false</EnableChangelog>
'    <AutoBackup>false</AutoBackup>
'    <snapinforce>false</snapinforce>
'    <snapinshowprogress>false</snapinshowprogress>
'    <snapinautoadd>0</snapinautoadd>
'    <snapinpermanentpath />
'  </ScriptPackager>
'</ScriptSettings>
'endregion

'USAGE: cscript|wscript wmiprocessorquery.vbs server

Dim CPUSink,ProcSink,objShell
On Error Resume Next

'check every 30 seconds for processor load > 30%.
strCPUQuery="Select * from __InstanceModificationEvent WITHIN 30 WHERE " &_
"TargetInstance ISA 'Win32_Processor' AND " &_
"TargetInstance.LoadPercentage >30"

'check every 10 seconds for existence of charmap.exe process
strProcQuery="Select * from __InstanceCreationEvent WITHIN 10 WHERE " &_
"TargetInstance ISA 'Win32_Process' AND " &_
"TargetInstance.Name='charmap.exe'"

Set objShell=CreateObject("Wscript.Shell")
Set CPUSink=WScript.CreateObject("WBemScripting.SWbemSink","CPUSINK_")
Set ProcSink=WScript.CreateObject("WBemScripting.SWbemSink","PROCSINK_")

Set oWMILocal=GetObject("winmgmts://")
oWMILocal.ExecNotificationQueryAsync ProcSink,strProcQuery
If err.number<>0 Then
 WScript.Echo "Oops! There was an error creating process event sink " &_
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

 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  WScript.Echo "Starting CPU Monitor"
'	WScript.Sleep 10000
  Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	If blnLoop Then SetupRemoteMonitoring(strComputer)
  Loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 

WScript.Echo "Launch CHARMAP to stop monitoring"


'Check if trigger process has been run, sleeping every 5 seconds.
While blnLoop
 WScript.sleep 5
Wend

WScript.Echo "Cancelling monitoring.  You can go ahead and close " &_
"the trigger application."
objShell.AppActivate("Character Map")

CPUSink.Cancel()
ProcSink.Cancel()
WScript.DisconnectObject(CPUSink)
WScript.DisconnectObject(ProcSink)

Set oWMILocal=Nothing
Set oWMIRemote=Nothing
Set CPUSink=Nothing
Set ProcSink=Nothing
WScript.quit

'*******************************************************************
Sub CPUSINK_OnObjectReady(objEvent,objContext)
strSystem=objEvent.Path_.Server
WScript.Echo Now & " - Processor on " & strSystem & " is " & objEvent.TargetInstance.LoadPercentage & "%."
End Sub

Sub PROCSINK_OnObjectReady(objEvent,objContext)
'trigger has been detected to close out this script
  blnLoop=False
End Sub

Sub SetupRemoteMonitoring(strComputer)
Err.Clear
Set oWMIRemote=GetObject("winmgmts://" & strComputer)
If err.number<>0 Then
 WScript.Echo "Oops!  There was an error connecting to " &_
  UCase(strComputer) & vbCrlf & "Error #" &err.number & VbCrLf &_
  "Description (if available): " & VbCrLf & " " &_
  err.description & VbCrLf & "Source (If available): " & _
  vbCrlf & " " & Err.source,-1,"CPU Monitoring":Err.Clear
 Exit Sub
Else
 oWMIRemote.ExecNotificationQueryAsync CPUSink,strCPUQuery
  If err.number<>0 Then
   WScript.Echo "Oops! There was an error creating CPU sink for " &_
    UCase(strComputer) & vbCrlf & "Error #" &err.number & vbCrlf &_
     "Description (if available): " & vbCrlf & " " &_
      err.description & vbCrlf & "Source (If available): " & _
   vbCrlf & " " & err.source,-1,"CPU Monitoring":Err.Clear
   Exit Sub
  Else
   WScript.Echo "Monitoring: " & strComputer
   err.Clear
  End If
End If
End Sub

'EOF