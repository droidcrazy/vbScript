Set fso = CreateObject("Scripting.FileSystemObject") 
 InputFile = "./servers.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	stroutput = getDriveLettersAndSize(strComputer)
	WScript.Echo """" & strComputer & """,""" & Left(stroutput,Len(stroutput)-2) & """"
	stroutput = ""
  loop
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 

Function getDriveLettersAndSize(strComputer)
    On Error Resume Next
    Set objWMIService = GetObject("winmgmts://" & strComputer & "/root/cimv2")
    If Err.Number Then
        getDriveLettersAndSize = "Error # " & CStr(Err.Number) & " " & Err.Description & "  "
        Err.Clear
    Else
    On Error GoTo 0
    Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3", , 48)
    For Each objItem In colItems
        getDriveLettersAndSize = getDriveLettersAndSize & objItem.Name & " " & Round(getDriveSizeTotal(strComputer,objItem.Name)/1024/1024/1024,2) & "GB, "
    Next
    End If
End Function

Function getDriveSizeTotal(strComputer, drvLetter)
On Error Resume Next
    Set objWMIService = GetObject("winmgmts://" & strComputer & "/root/cimv2")
    strTemp = strComputer
    If Err.Number Then
        getDriveSizeTotal = "0"
        Err.Clear
    Else
    On Error GoTo 0
    Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3", , 48)
    For Each objItem In colItems
        If UCase(objItem.Name) = UCase(drvLetter) Then
        getDriveSizeTotal = objItem.Size
        End If
    Next
    End If
End Function