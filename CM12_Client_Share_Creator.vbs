'This is where you want to put the files.
strDest = "c:\SCCM_SHARE\"

Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")


Set objFile = objFSO.GetFile(strPath)
strMyFolder = objFSO.GetParentFolderName(objFile) 


If Not objFSO.FolderExists(strDest) Then
objFSO.CreateFolder strDest
End If

strCMD = "xcopy """ & strMyFolder & """ """ & strDest & """ /E /Y"
objShell.Run strCMD
strCMD = "icacls " & strDest & " /grant ""NOBLE\Domain Users:(OI)(CI)RX"""
objShell.Run strCMD
strCMD = "icacls " & strDest & " /grant ""Everyone:(OI)(CI)RX"""
objShell.Run strCMD

For Each strArgument In WScript.Arguments 
	If LCase(strArgument) = "install" Then
		objShell.run "ccmsetup.exe /forceinstall /source:""" & strMyFolder """ CCMLOGMAXHISTORY=5 SMSCACHESIZE=15360 SMSSITECODE=HOU SMSMP=HOUSCCM.NOBLE.CC FSP=HOUSCCM CCMLOGMAXSIZE=512000 DISABLESITEOPT=TRUE DISABLECACHEOPT=TRUE"
	End If
Next
