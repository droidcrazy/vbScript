'This is where you want to put the files.
strDest = "c:\SCCM_SHARE\"

Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(strPath)
strMyFolder = objFSO.GetParentFolderName(objFile) 
WScript.StdOut.Write strMyFolder
WScript.Quit 0

'If Not objFSO.FolderExists(strDest) Then
'objFSO.CreateFolder strDest
'End If

'strCMD = "xcopy """ & strMyFolder & """ """ & strDest & """ /E"
'objShell.Run strCMD
'strCMD = "icacls " & strDest & " /grant ""NOBLE\Domain Users:(OI)(CI)RX"""
'objShell.Run strCMD
'strCMD = "icacls " & strDest & " /grant ""Everyone:(OI)(CI)RX"""
'objShell.Run strCMD
