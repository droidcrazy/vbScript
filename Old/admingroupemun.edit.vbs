on error resume next 
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject")
Set grpdict = CreateObject("Scripting.Dictionary")
Set usrdict = CreateObject("Scripting.Dictionary")
OutputFile = "./local_groups_all.csv" 

   inputfile = "./servers.txt"
   If fso.FileExists(InputFile) Then 
       Set txtStreamIn = fso.OpenTextFile(InputFile) 
       Set txtStreamOut = fso.OpenTextFile(OutputFile, 2, True) 
       printout "Server,Group,Members that are users, Members that are groups" 
       Do While Not (txtStreamIn.AtEndOfStream) 
           server = txtStreamIn.ReadLine 
           chkgroups server
       Loop 
           printout ""
           printout ""
           printout "Group Name, Members"
           getusers
           printout ""
           printout ""
           printout "Login,Full Name,Last Login Date,Disabled"
           getuserinfos
	wscript.echo "all groups enumeration is done. output file is " & OutputFile
   Else 
       WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
       usage 
   End If 

On Error Resume Next 

Sub chkgroups(strComputer) 
On Error Resume Next
	Line = "" 
	err.clear 
	Set colgroups = GetObject("WinNT://" & strComputer & ",computer") 
	colGroups.GetInfo 
	If err.number <> 0 then 
		printout strComputer & ",error binding to computer" 
	Else 
		colgroups.Filter = Array("Group") 
		For Each objGroup in colGroups
			If objGroup.name = "Administrators" Then
				Line = strcomputer & "," & objgroup.name 
				For Each objUser in objgroup.members 
					If objUser.name <> "" Then 
						If objUser.class = "User" Then
							thisuser = getFullName(objuser.adspath)
							usrdict.add thisuser, thisuser
							If users = "" Then
									users = thisuser
								Else
									users = users & ", " & thisuser
							End If
						Else
							thisgroup = getFullName(objuser.adspath)
							grpdict.add thisgroup, thisgroup
							If memgroup = "" then 
									memgroup = thisgroup
								Else
									memgroup = memgroup & ", " & thisgroup
							End If
						End If
					Else 
						users = ",no accounts in this group"
					Exit For 
					End If 
				Next
				printout line & ",""" & users & """" & ",""" & memgroup & """"         
			End If
		Next
		err.clear 
	End If 
End Sub 'chkgroups 

Sub getusers()
On Error Resume Next
	For Each group in grpdict
	Line = ""
	memberout = ""
	 Set grp = GetObject("WinNT://" & group)
		Line = group & ","
For Each member In grp.Members
	If member.class = "User" Then
		usrdict.add getfullname(member.adspath), getfullname(member.adspath)
	End If
	If memberout = "" Then
		memberout = getfullname(member.adspath) & " (" & member.class & ")"
	Else
		memberout = memberout & ", " & getfullname(member.adspath) & " (" & member.class & ")"
	End If
	If (member.Class = "Group") Then
		For Each obj In member.Members
			If obj.class = "User" Then
				usrdict.add getfullname(obj.adspath), getfullname(obj.adspath)
			End If
			If objout = "" Then
				objout = "[" & getfullname(obj.adspath) & " (" & obj.Class & ")"
			Else
				objout = objout & ", " & getfullname(obj.adspath) & " (" & obj.Class & ")"
		End If
		Next
		objout = objout & "]"
		memberout = memberout & ", " & objout
	End If
Next
	printout Line & """" & memberout & """"
Next
End Sub 'getusers

Sub getuserinfos
On Error Resume Next
Set userlist = CreateObject("System.Collections.ArrayList")
For each user in usrdict
userlist.add user
Next
userlist.sort
For each user in userlist
Set userobj = GetObject("WinNT://" & user & ",user")
printout """" & user & """,""" & userobj.fullname & """,""" & userobj.lastlogin & """,""" & userobj.accountdisabled & """"
Next
End Sub 'getuserinfos

Sub printout (data) 
'	WScript.Echo data 
	txtStreamOut.writeline data 
End Sub 'printout 

Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName 
