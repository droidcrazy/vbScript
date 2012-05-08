on error resume Next
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject")
Set grpdict = CreateObject("Scripting.Dictionary")
Set usrdict = CreateObject("Scripting.Dictionary")

Set WshShell = wscript.createobject("wscript.shell")
answer = MsgBox("Click OK to choose an input file"  & Chr(10) & Chr(13) & "or click Cancel to use the default of servers.txt", 65, "Admin Group Enumeration Tool")
If answer = 1 Then
Set ObjFSO = CreateObject("UserAccounts.CommonDialog")
ObjFSO.Filter = "Text Documents|*.txt"
ObjFSC.Title = "Select an Input File"
ObjFSO.FilterIndex = 3
ObjFSO.InitialDir = wshshell.currentdirectory
InitFSO = ObjFSO.ShowOpen
If InitFSO = False Then
    Wscript.Echo "Script Error: Please select a file!"
    Wscript.Quit
Else
    inputfile = ObjFSO.FileName
End If

temparray = split(inputfile, "\")
inputfilename = temparray(UBound(temparray))
inputfilename = Left(inputfilename, Len(inputfilename) - 4)
'wscript.echo inputfilename

'wscript.echo inputfile
'wscript.quit
Else

inputfile = "./servers.txt"
End If


	If fso.FileExists(InputFile) Then 
			Set objExcel = CreateObject("Excel.Application") 
			objExcel.Visible = true
			'objExcel.Workbooks.Add
			Set objWorkbook = objExcel.Workbooks.Add()
			Set objWorksheet = objWorkbook.Worksheets(1)
			objExcel.worksheets(2).delete
			objExcel.worksheets(2).delete
		
			chkgrpmaster
			getusers
			getuserinfos
		strTime = Right(100 + Month(Now), 2) & "-" & Right (100 + Day(Now), 2) & "-" & Year(Now) & "." & Right(100 + hour(now), 2) & Right( 100 + Minute(now), 2)
		OutputFile = wshshell.currentdirectory & "admingroupenum." & inputfilename & "." & strTime & ".xlsx"
		objWorkbook.SaveAs OutputFile
		objExcel.quit
		wscript.echo "Administrator group enumeration is done. Output file is " & OutputFile
	Else 
		WScript.Echo "Input file doesnt exist. Please make sure that the " & Right(inputfile, Len(inputfile) - 2) & " file exists in the directory you are running this from." 
	End If 

On Error Resume Next 

Sub chkgrpmaster()
On Error Resume Next
		objExcel.worksheets(1).Activate
		objExcel.worksheets(1).Name = "Computer Groups"
		objExcel.Cells(1, 1).Value = "Computer Name" 
		objExcel.Cells(1, 2).Value = "Group Name" 
		objExcel.Cells(1, 3).Value = "Members that are Users" 
		objExcel.Cells(1, 4).Value = "Members that are Groups"
		rowVar=2 
Set txtStreamIn = fso.OpenTextFile(InputFile)
	Do While Not (txtStreamIn.AtEndOfStream) 
			strComputer = txtStreamIn.ReadLine 
			'wscript.echo ">" & strComputer & "<"
			strComputer = cleanme(strComputer)
			If Left(strComputer, 2) = "\\" Then
				strcomputer = Right(strcomputer, Len(strcomputer) - 2)
			End If
			'wscript.echo ">" & strComputer & "<"
err.clear 
			objExcel.Cells(rowVar, 1).Value = strComputer
	Set colgroups = GetObject("WinNT://" & strComputer & ",computer") 
	colGroups.GetInfo 
	If err.number <> 0 then 
		objExcel.Cells(rowVar, 2).Value = "error binding to computer" & "Error #" & Err.Number & " " & Err.Description
	Else 
		colgroups.Filter = Array("Group") 
		For Each objGroup in colGroups
			If objGroup.name = "Administrators" Then
					objExcel.Cells(rowVar, 2).Value = objgroup.name 
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
				objExcel.Cells(rowVar, 3).Value = users
				objExcel.Cells(rowVar, 4).Value = memgroup
				users = ""
				memgroup = ""
			End If
		Next
		err.clear 
	End If 

			rowVar = rowVar +1
		Loop 
			objExcel.Cells.Select
			objExcel.Cells.EntireColumn.AutoFit
			objExcel.Range("B2").Select
			objExcel.ActiveWindow.FreezePanes = True
End Sub 'chkgrpmaster

Sub getusers()
On Error Resume Next
		rowVar = 2
		set objWorksheet = objExcel.Sheets.Add( , objExcel.WorkSheets(objExcel.WorkSheets.Count))
		objExcel.worksheets(2).Activate
		objExcel.worksheets(2).Name = "Group Enumeration"
		objExcel.Cells(1, 1).Value = "Group Name" 
		objExcel.Cells(1, 2).Value = "Members" 
For Each group in grpdict
	memberout = ""
	 Set grp = GetObject("WinNT://" & group)
		objExcel.Cells(rowVar, 1).Value = group
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
	objExcel.Cells(rowVar, 2).Value = memberout
	rowVar = rowVar +1
Next
	objExcel.Columns("A:A").EntireColumn.AutoFit
	objExcel.Range("B2").Select
	objExcel.ActiveWindow.FreezePanes = True
End Sub 'getusers


Sub getuserinfos
On Error Resume Next
		rowVar = 2
		set objWorksheet = objExcel.Sheets.Add( , objExcel.WorkSheets(objExcel.WorkSheets.Count))
		objExcel.worksheets(3).Activate
		objExcel.worksheets(3).Name = "User Enumeration"
		objExcel.Cells(1, 1).Value = "Login ID" 
		objExcel.Cells(1, 2).Value = "Full Name" 
		objExcel.Cells(1, 3).Value = "Last Login Date"
		objExcel.Cells(1, 4).Value = "Disabled"
Set userlist = CreateObject("System.Collections.ArrayList")
For each user in usrdict
userlist.add user
Next
userlist.sort
For each user in userlist
Set userobj = GetObject("WinNT://" & user & ",user")
	objExcel.Cells(rowVar, 1).Value = user
	objExcel.Cells(rowVar, 2).Value = userobj.fullname
	objExcel.Cells(rowVar, 3).Value = userobj.lastlogin
	objExcel.Cells(rowVar, 4).Value = userobj.accountdisabled
	rowVar = rowVar +1
Next
	objExcel.Cells.Select
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range("B2").Select
	objExcel.ActiveWindow.FreezePanes = True
End Sub 'getuserinfos

Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName 

Function cleanme(str)
'On Error Resume Next
Dim regEx
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.Pattern = "\s"
Str = Trim(regEx.Replace(str, ""))
'regEx.Pattern = "\\\\"
'Str = Trim(regEx.Replace(str, ""))
Return str 
End Function 'cleanme