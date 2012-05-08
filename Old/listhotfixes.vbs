Set WshShell = wscript.createobject("wscript.shell")
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject")
answer = MsgBox("Click OK to choose an input file"  & Chr(10) & Chr(13) & "or click Cancel to use the default of servers.txt", 65, "Admin Group Enumeration Tool")
If answer = 1 Then
Set ObjFSO = CreateObject("UserAccounts.CommonDialog")
ObjFSO.Filter = "Text Documents|*.txt"
'ObjFSO.Title = "Select an Input File"
ObjFSO.FilterIndex = 3
ObjFSO.InitialDir = wshshell.currentdirectory
InitFSO = ObjFSO.ShowOpen
If InitFSO = False Then
    Wscript.Echo "Script Error: Please select a file!"
    Wscript.Quit
Else
    inputfile = ObjFSO.FileName
End If
End If


If fso.FileExists(InputFile) Then 
	Set objExcel = CreateObject("Excel.Application") 
	objExcel.Visible = False
	'objExcel.Workbooks.Add
	Set objWorkbook = objExcel.Workbooks.Add()
	Set objWorksheet = objWorkbook.Worksheets(1)
	objExcel.worksheets(2).delete
	objExcel.worksheets(2).delete
	
	

		getUpdatesInfo
		
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 
 
Sub getUpdatesInfo()
On Error Resume Next
		objExcel.worksheets(1).Activate
		objExcel.worksheets(1).Name = "Hotfixes"
		objExcel.Cells(1, 1).Value = "Computer Name" 
		'objExcel.Cells(1, 2).Value = "Description" 
		objExcel.Cells(1, 2).Value = "Hotfix ID" 
		objExcel.Cells(1, 3).Value = "Installation Date"
		objExcel.Cells(1, 4).Value = "Installed By"
		rowVar=2 

  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	strComputer = ereg_replace(strComputer, "/s", "", True)
	If strComputer <> "" Then
			
		Set objWMIService = GetObject("winmgmts:" _
		    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		If Err.Number <> 0 Then
			objExcel.Cells(rowVar, 1).Value = strComputer
			objExcel.Cells(rowVar, 2).Value = "Error # " & CStr(Err.Number) & " " & Err.Description
			'printout "Error # " & CStr(Err.Number) & " " & Err.Description
			rowvar = rowvar+1
		    Err.Clear
		Else 
		Set colQuickFixes = objWMIService.ExecQuery _
		    ("Select * from Win32_QuickFixEngineering")
		
		For Each objQuickFix in colQuickFixes
		If objQuickFix.HotFixID <> "File 1" Then
			objExcel.Cells(rowVar, 1).Value = strComputer
			'objExcel.Cells(rowVar, 2).Value = objQuickFix.Description
			objExcel.Cells(rowVar, 2).Value = objQuickFix.HotFixID
			objExcel.Cells(rowVar, 3).Value = objQuickFix.InstalledOn
			objExcel.Cells(rowVar, 4).Value = objQuickFix.InstalledBy
			rowvar = rowvar+1
		    'printOut "Computer: " & objQuickFix.CSName
		    'printOut "Description: " & objQuickFix.Description
		    'printOut "Hot Fix ID: " & objQuickFix.HotFixID
		    'printOut "Installation Date: " & objQuickFix.InstalledOn
		    'printOut "Installed By: " & objQuickFix.InstalledBy
		    'printOut objQuickFix.InstalledOn & " -- HotfixID: " & objQuickFix.HotFixID & " Installed by: " & objQuickFix.InstalledBy
		    End If
		    
		Next
		End If
	End If
  Loop
  			objExcel.Cells.Select
			objExcel.Cells.EntireColumn.AutoFit
			objExcel.Range("B2").Select
			objExcel.ActiveWindow.FreezePanes = True
			objExcel.Visible = True

End Sub 'getUpdatesInfo

Sub printOut (data) 
	'WScript.Echo data
	txtStreamOut.writeline data 
End Sub 'printOut 

Function cleanme(str)
On Error Resume Next
Dim regEx
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.Pattern = "\s"
str = Trim(regEx.Replace(str, ""))
return str 
End Function 'cleanme

function ereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
  ' Function replaces pattern with replacement
  ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)  
  dim objRegExp : 
  set objRegExp = new RegExp  
  With objRegExp    
	  .Pattern = strPattern    
	  .IgnoreCase = varIgnoreCase    
	  .Global = True  
  end with  
  ereg_replace = objRegExp.replace(strOriginalString, strReplacement)  
  set objRegExp = Nothing
end function

