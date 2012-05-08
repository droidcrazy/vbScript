on error resume next 
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject") 
OutputFile = "./whos_logged_on.csv" 
 inputfile = "./workstations.txt"
 If fso.FileExists(InputFile) Then 
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True) 
  printout "username,computer" 
  Do While Not (txtStreamIn.AtEndOfStream) 
	wkstn = txtStreamIn.ReadLine 
	wkstn = cleanme(wkstn)
'	wscript.echo "connecting to " & wkstn & "..."
	chkwho wkstn 
  loop
	wscript.echo "Who's logged in enumeration is done. output file is " & OutputFile 
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the workstations.txt file exists in the directory you are running this from." 
  usage 
 End If 


On Error Resume Next 


Sub chkwho(strComputer)
On Error Resume Next
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
  For Each objItem in colItems
'	Wscript.Echo "UserName: " & objItem.UserName & " is logged in at computer " & strComputer
	printout objItem.UserName & "," & strComputer
  Next
End Sub 'chkwho


Sub printOut (data) 
' WScript.Echo data 
 txtStreamOut.writeline data 
End Sub 'printOut 


Function cleanme(str)
Dim regEx
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.Pattern = "\s"
str = Trim(regEx.Replace(str, ""))
return str 
End Function 'cleanme