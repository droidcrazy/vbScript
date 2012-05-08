' Ping Multiple Computers


Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./servers.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  'Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	'server = cleanme(server)
	pingme server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 


Sub pingme(machine)
'For Each machine in aMachines
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & machine & "'")
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
        	'WScript.Echo objStatus.StatusCode
            WScript.Echo("Computer " & machine & " did not respond.") 
        Else
        	WScript.Echo("Computer " & machine & " responded on " & objStatus.ProtocolAddress & " in " & objStatus.ResponseTime & "ms.")
        End If
    Next
'Next
End Sub

Function cleanme(str)
Dim regEx
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.Pattern = "\s"
str = Trim(regEx.Replace(str, ""))
'return str 
End Function 'cleanme