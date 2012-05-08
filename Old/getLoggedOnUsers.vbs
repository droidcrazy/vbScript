If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./houston.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	server = txtStreamIn.ReadLine 
	GetLoggedOnUser server
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 



' List the User Logged on to a Remote Computer
Sub GetLoggedOnUser(strComputer)
'strComputer = "."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
If Err.Number <> 0 Then WScript.Echo "Error binding to computer " & strComputer & ":" & Err.Description : Err.Clear : Exit Sub
On Error Goto 0
Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo strComputer & " logged-on user: " & objComputer.UserName
Next

Set colSessions = objWMIService.ExecQuery _ 
    ("Select * from Win32_LogonSession") 'where logontype = 10") 
On Error Resume Next
count = 0
count = colSessions.count
On Error Goto 0
If Count = 0 Then 
   Wscript.Echo "No Logon Sessions found." 
Else 
   WScript.Echo "Logon Sessions:"
   For Each objSession in colSessions 
   On Error Resume Next  
     Set colList = objWMIService.ExecQuery("Associators of " _ 
         & "{Win32_LogonSession.LogonId=" & objSession.LogonId & "} " _ 
         & "Where AssocClass=Win32_LoggedOnUser Role=Dependent" ) 
     For Each objItem in colList
     display = False 
     Select Case objsession.logontype
     Case 0
     logontype = "System Use Only."
     logontypedescrition = "Used only by the System account."
     display = False
     Case 2
     logontype = "Interactive"
     logontypedescrition = "Intended for users who are interactively using the machine," _
     & " such as a user being logged on by a terminal server, remote shell, or similar process."
     display = True
     Case 3
     logontype = "Network"
     logontypedescrition = "Intended for high-performance servers to authenticate clear text passwords." _
     & " LogonUser does not cache credentials for this logon type."
     display = True
     Case 4
     logontype = "Batch"
     logontypedescrition = "Intended for batch servers, where processes can be executed on behalf of " _
     & "a user without their direct intervention; or for higher performance servers that process many " _
     & "clear-text authentication attempts at a time, such as mail or Web servers. LogonUser does not " _
     & "cache credentials for this logon type."
     display = False
     Case 5
     logontype = "Service"
     logontypedescrition = "Indicates a service-type logon. The account provided must have the service " _
     & "privilege enabled."
     display = False
     Case 6
     logontype = "Proxy"
     logontypedescrition = "Indicates a proxy-type logon."
     display = False
     Case 7
     logontype = "Unlock"
     logontypedescrition = "This logon type is intended for GINA DLLs logging on users who are " _
     & "interactively using the machine. This logon type allows a unique audit record to be " _
     & "generated that shows when the workstation was unlocked."
     display = False
     Case 8
     logontype = "Network Cleartext"
     logontypedescrition = "Windows Server 2003, Windows 2000, and Windows XP:  " _
     & "Preserves the name and password in the authentication packages, allowing the " _
     & "server to make connections to other network servers while impersonating the client. " _
     & "This allows a server to accept clear text credentials from a client, call LogonUser, " _
     & "verify that the user can access the system across the network, and still communicate " _
     & "with other servers."
     display = False
     Case 9
     logontype = "New Credentials"
     logontypedescrition = "Windows Server 2003, Windows 2000, and Windows XP:  " _
     & "Allows the caller to clone its current token and specify new credentials for " _
     & "outbound connections. The new logon session has the same local identify, but " _
     & "uses different credentials for other network connections."
     display = False
     Case 10
     logontype = "Remote Interactive"
     logontypedescrition = "Terminal Services session that is both remote and interactive."
     display = True
     Case 11
     logontype = "Cached Interactive"
     logontypedescrition = "Terminal Services session that is both remote and interactive."
     display = False
     Case 12
     logontype = "Cached Remote Interactive"
     logontypedescrition = "Attempt cached credentials without accessing the network."
     display = True
     Case 13
     logontype = "Cached Unlock"
     logontypedescrition = "Workstation logon."
     display = False
     Case Else
     logontype = "Unknown"
     logontypedescrition = "No Description."
     display = False
     End Select
If Err.Number <> 0 Then 
If display Then WScript.Echo "Could not find LogonID:" & objSession.LogonId & " on " & strComputer & "." & vbTab & "Logon Type: " & logontype & vbTab & "Authentication Package: " & objsession.authenticationpackage
Err.Clear
End If
starttime = WMIDateStringToDate(objsession.starttime)
If Err.Number <> 0 Then WScript.Echo Err.Description
       If display Then WScript.Echo "Username: " & objItem.Name & vbTab & "FullName: " & objItem.FullName & vbTab & "Start Time: " & starttime & vbTab & "Logon type: " & logontype & vbTab & "Authentication Package: " & objsession.authenticationpackage
     Next 
If Err.Number <> 0 Then Err.Clear
On Error Goto 0
   Next 
End If 

End Sub

Function WMIDateStringToDate(dtmWMI)
 WMIDateStringToDate = CDate(Mid(dtmWMI, 5, 2) & "/" & _
 Mid(dtmWMI, 7, 2) & "/" & Left(dtmWMI, 4) _
 & " " & Mid (dtmWMI, 9, 2) & ":" & _
 Mid(dtmWMI, 11, 2) & ":" & Mid(dtmWMI, _
 13, 2))
End Function