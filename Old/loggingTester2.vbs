' Write to a Custom Event Log Using EventCreate
Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Echo wshshell.ExpandEnvironmentStrings("%COMPUTERNAME%")
version = runme("%comspec% /c ver")
WScript.Echo version
strCommand = "eventcreate /s houmwibigbro /so " & WScript.ScriptName & " /T Warning /ID 100 /L Scripts /D """ & version & """"
WScript.Echo runme(strCommand)
WScript.Quit

Function runme(strCommand)
Dim WshShell, oExec
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
Set oExec    = WshShell.Exec(strCommand)
If Err.Number <> 0 Then runme = "Error " & Hex(Err.Number) & " occured. Description: " & Err.Description & " Source: " & Err.Source :Err.Clear : Exit Function
On Error Goto 0
Dim allInput, tryCount
allInput = ""
tryCount = 0
Do While True
     Dim input
     input = ReadAllFromAny(oExec)
     If -1 = input Then
          If tryCount > 10 And oExec.Status = 1 Then
               Exit Do
          End If
          tryCount = tryCount + 1
          WScript.Sleep 100
     Else
          allInput = allInput & input
          tryCount = 0
     End If
Loop
runme = allInput
End Function

Function ReadAllFromAny(oExec)
     If Not oExec.StdOut.AtEndOfStream Then
          ReadAllFromAny = oExec.StdOut.ReadAll
          Exit Function
     End If
     If Not oExec.StdErr.AtEndOfStream Then
          ReadAllFromAny = "STDERR: " + oExec.StdErr.ReadAll
          Exit Function
     End If
     ReadAllFromAny = -1
End Function