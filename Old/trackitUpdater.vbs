'Trackit8LogonScriptAddin.wbt rewrite
On Error Resume Next 'this is like the errormode(@off) line

'check to see if called by cscript or wscript
Select Case Right(UCase(WScript.FullName), 11)
Case UCase("CScript.exe")
c = True
w = False
Case UCase("WScript.exe")
c = False
w = True
Case Else
WScript.Quit
End Select

'Define constants used in script
'Constants for file system object:
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2

'check for versions from command line -- if they don't exist use hardcoded values
If WScript.Arguments.Named.Exists("tiremote") Then
correctVer1 = WScript.Arguments.Named("tiremote")
Else
correctVer1 = "8.0.50.124"
End If

If WScript.Arguments.Named.Exists("tiremoteservice") Then
correctVer2 = WScript.Arguments.Named("tiremoteservice")
Else
correctVer2 = "8.0.50.124"
End If

'Create objects that we will work through
Set fso = CreateObject("scripting.filesystemobject")
Set WshShell = CreateObject("wscript.Shell")
'Today = Now 'I'm showing you this, but in the script, I'm going to use the vb function "now" -- it's easier, so I'm commenting this out
WinDir = WshShell.ExpandEnvironmentStrings("%windir%") 'this uses the shell object we created to get the environment variable %windir% 
logfile = WinDir & "\TIUpdate.txt"
'check the log file size and rotate if need be
If fso.FileExists(logfile) Then
	If CInt(fso.GetFile(logfile).Size) > CInt("10240") Then rotateLog(logfile)
End If

Set l = fso.OpenTextFile(logfile, ForAppending, True) 'This opens the TIUpdate.txt in %windir% in an appending state as a 'textstream' named 'l' (named 'handle' in the winbatch script.) -- the true at the end tells it to create the file if it doesn't exist
printOut "Starting " & WScript.ScriptFullName
InstallDir = windir & "\TIREMOTE"
TargetFile1 = InstallDir & "\TIRemote.exe"
TargetFile2 = InstallDir & "\TIRemoteService.exe"
Installer="\\pxhoustrackit\TrackIt8\Installers\WorkstationManager\TIWSMgr.exe"

Call main()
Call CleanupAndExit(0)

Sub main()
If fso.FolderExists(InstallDir) Then
	If fso.FileExists(TargetFile1) And fso.FileExists(TargetFile2) Then
		If Not chkVersions(correctVer1, correctVer2, False) Then
			printOut "Versions are not up to date, attempting to remove and install updated version."
			If uninstall() Then
				If removeDir(windir & "\TIREMOTE") Then
					install
				End If
			End If
		Else
			printOut "File Versions are up to date."
		End If
	Else
		printOut "One or both files are missing. Reinstalling."
		If removeDir(windir & "\TIREMOTE") Then
			install
		End If
	End If	
Else
	printOut "Installation directory is missing. Reinstalling."
	install
End If
End Sub 'main

Function chkVersions(correctVer1, correctVer2, silent)
	chkVersions = False
	Ver1 = fso.GetFileVersion(TargetFile1)
	Ver2 = fso.GetFileVersion(TargetFile2)
	If Not silent Then printOut "Current file versions: " & TargetFile1 & ": " & Ver1 & " & " & TargetFile2 & ": " & Ver2
    If Ver1 = correctVer1 And Ver2 = correctVer2 Then chkVersions = True
End Function 'chkVersions

Function printOut(msg)
	If c Then WScript.StdOut.WriteLine Now & " -- " & msg
	l.WriteLine Now & " -- " & msg
End Function 'printOut

Function uninstall()
	uninstall = False
	ReturnCode = WshShell.Run(TargetFile1 & " /UNINSTALL", 0, True)
	If ReturnCode = 0 Then
		uninstall = true
		printOut "Uninstallation was successful."
	Else
		uninstall = False
		printOut "Uninstallation was unsuccessful. The ReturnCode was " & ReturnCode & "."
	End If
End Function 'uninstall

Function install()
	install = False
	printOut "Attempting to install current version of TrackIt."
	returncode = WshShell.Run(installer, 0, True)
	counter = 0
	Do While Not chkVersions(correctVer1, correctVer2, True)
		counter = counter +1
		If counter > 10 Then returncode = 1:Exit Do
		WScript.Sleep 5000
	loop
	If returncode = 0 Then
		install = True
		printOut "Installation was successful."
	Else
		install = False
		printOut "Installation was unsuccessful. The return code was " & returncode & "."
	End If
End Function 'install

Function removeDir(strDir)
	On Error Resume Next
	printOut "Attempting to remove " & strDir & " directory."
	WScript.Sleep 5000
	counter = 0
	Do While fso.FolderExists(strDir)
		counter = counter + 1
		fso.DeleteFolder strDir, True
		If Err.Number <> 0 Then 
			printOut "Error 0x" & Hex(Err.Number) & " occured: " & Err.Description & ", Source: " & Err.Source
			Err.Clear 
		Else 
			Exit Do
		End If
		If counter > 10 Then Exit Do
		WScript.Sleep 5000
	Loop
	If fso.FolderExists(strDir) Then
		printOut "After trying " & counter & " time(s), deletion of " & strDir & " was unsuccessful."
		removeDir = False
	Else
		printOut "After trying " & counter & " time(s), deletion of " & strDir & " was successful."
		removeDir = True
	End If
End Function 'remove

Function rotatelog(logfile)
	If fso.FileExists(logfile & ".old2") Then fso.DeleteFile logfile & ".old2"
	If fso.FileExists(logfile & ".old1") Then fso.MoveFile logfile & ".old1", logfile & ".old2"
	fso.MoveFile logfile, logfile & ".old1" 
End Function 'rotatelog

Sub CleanupAndExit(ExitStatus)
	printOut "Finishing script and closing."
	l.Close
	WScript.Quit(ExitStatus Mod 255)
End Sub 'CleanupAndExit
