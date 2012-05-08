Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strProcQuery="Select * from __InstanceCreationEvent WITHIN 10 WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name='charmap.exe'"

On Error Resume Next
Set ProcSink=WScript.CreateObject("WBemScripting.SWbemSink","PROCSINK_")

Set oWMILocal=GetObject("winmgmts://")
oWMILocal.ExecNotificationQueryAsync ProcSink,strProcQuery

If err.number<>0 Then
 WScript.quit
Else
 Err.Clear
 On Error Goto 0
End If
blnLoop=True

kill "procmon.exe"
FilePath = "h:\testing\"

logfile = FilePath & tdStamp & ".pml"
objShell.Run FilePath & "procmon\procmon.exe /acceptlicense /profiling /nofilter /backingfile " & logfile,2

While blnLoop
WScript.Sleep 5000
Set objFile = objFSO.getFile(logfile)
Do While objfile.size < 1024*1024*20
WScript.Sleep 5000
Loop
objshell.AppActivate "process monitor"
WScript.Sleep 500
objshell.SendKeys "%fx"
WScript.Sleep 500
oldlog = logfile
logfile = FilePath & tdStamp & ".pml"
objShell.Run FilePath & "procmon\procmon.exe /acceptlicense /profiling /nofilter /backingfile " & logfile,2
WScript.Sleep 500
On Error Resume Next
objFSO.CopyFile oldlog , "\\pxhouscorp02\networking\logarchive\"
objfso.DeleteFile oldlog
Err.Clear
On Error Goto 0
Wend

kill "charmap.exe"
kill "procmon.exe"

ProcSink.Cancel()
WScript.DisconnectObject(ProcSink)
Set ProcSink=Nothing
WScript.quit


Sub kill(procname)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\localhost\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & procname & "'")
For Each objProcess in colProcessList
    objProcess.terminate()
Next
End Sub

Function tdstamp()
tdStamp = Right(Hour(Now)+100,2)&Right(Minute(Now)+100,2)&Right(Second(Now)+100,2)&"-"&Right(Month(Now)+100,2)&Right(Day(Now)+100,2)&Right(Year(Now)+10000,2)
End Function

Sub PROCSINK_OnObjectReady(objEvent,objContext)
'trigger has been detected to close out this script
  blnLoop=False
kill "charmap.exe"
kill "procmon.exe"
ProcSink.Cancel()
WScript.DisconnectObject(ProcSink)
Set ProcSink=Nothing
WScript.Quit
End Sub