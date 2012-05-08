pathBackTmp = ".\BACKUPFOLDER" 

backupSSLcerts getComputer

'Backup SSL certs 
Sub backupSSLcerts(strComputer)
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC") 
For Each objweb in objIIS 
If lCase(objweb.Class) = "iiswebserver" Then 
set iiscertobj = CreateObject("IIS.CertObj") 
iiscertobj.InstanceName = "W3SVC/" & objweb.Name 
'On Error Resume Next 
iiscertobj.Export _ 
pathBackTmp & "\" & objweb.ServerComment & ".pfx", _ 
objweb.ServerComment, _ 
true, true, false 
If err.number = 0 Then 
call printLine("Backup SSL Certificate", objweb.ServerComment & " (" & objweb.Name & ")") 
End If 
err.Clear 
On Error Goto 0 
Set iiscertobj = nothing 
End If 
Next 
Set objIIS = Nothing 
End Sub 

'Get computer name 
Function getComputer() 
Set objNet = WScript.CreateObject("WScript.Network") 
getComputer= objNet.ComputerName 
Set objNet = Nothing 
End Function 

'Print message line 
Function printLine(strLabel, strMessage) 
strLabel = trim(left(strLabel,30)) 
strLabel = strLabel & Replace(Space(30-len(strLabel))," ",".") 
WScript.Echo "> " & strLabel & ": " & strMessage 
End Function
