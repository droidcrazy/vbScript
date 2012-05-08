on error resume next 
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject") 
OutputFile = "./local_groups_all.csv" 

   inputfile = "./servers.txt"
   If fso.FileExists(InputFile) Then 
       Set txtStreamIn = fso.OpenTextFile(InputFile) 
       Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True) 
       printout "Server,Group,Members that are users, Members that are groups" 
       Do While Not (txtStreamIn.AtEndOfStream) 
           server = txtStreamIn.ReadLine 
           chkgroups server 
       loop 
	wscript.echo "all groups enumeration is done. output file is " & OutputFile
   Else 
       WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
       usage 
   End If 

On Error Resume Next 

Sub chkgroups(strComputer) 
   line = "" 
   err.clear 
   Set colgroups = GetObject("WinNT://" & strComputer & ",computer") 
   colGroups.GetInfo 
   if err.number <> 0 then 
       printout strComputer & ",error binding to computer" 
   Else 
       colgroups.Filter = Array("Group") 
       For Each objgroup In colGroups
          line = strcomputer & "," & objgroup.name 
          
	For Each objuser in objgroup.members 
		if objuser.name <> "" Then 
		if objuser.class = "User" Then
				if users = "" then
					users = getFullName(objuser.adspath)
				else
					users = users & ", " & getFullName(objuser.adspath)
				end if
			else
				if memgroup = "" then 
					memgroup = getFullName(objuser.adspath)
				else
					memgroup = memgroup & ", " & getFullName(objuser.adspath)
				end if
			end if
		Else 
                       users = ",no accounts in this group"
	exit for 
       end If 
           Next
           printout line & ",""" & users & """" & ",""" & memgroup & """"         
       Next
       err.clear 
   End If 
End Sub 'chkgroups 

Sub printout (data) 
   WScript.Echo data 
'   txtStreamOut.writeline data 
End Sub 'printout 

Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName 
