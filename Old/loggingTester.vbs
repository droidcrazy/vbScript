Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Const log_SUCCESS = 0
Const log_ERROR = 1
Const log_WARNING = 2
Const log_INFORMATION = 4
Const log_AUDIT_SUCCESS = 8
Const log_AUDIT_FAILURE = 16



strLogLocation = objShell.CurrentDirectory
strLogSystem = "houmwibigbro"

strScript = Wscript.ScriptName
strLogName = Left(strScript, Len(strScript)-4) & "Log.txt"

Set objLogFile = objFSO.OpenTextFile(strLogLocation & strLogName, ForAppending, True)


strMsg = "Script Started"
objLogFile.WriteLine(Now & vbTab & strMsg)

strMsg = "Script Ended"
objLogFile.WriteLine(Now & vbTab & strMsg)


If Err.Number <> 0 Then
objShell.LogEvent log_WARNING,  "Error: " & strScript & vbCRLF & Err.Description, strLogSystem
Else
objShell.LogEvent log_SUCCESS,  "Success: " & strScript, strLogSystem
End If

objLogFile.Close

Set objLogFile = Nothing
Set objFSO = Nothing
Set objShell = Nothing