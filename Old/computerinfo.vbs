Set WshShell = wscript.createobject("wscript.shell")
Set objArgs = WScript.Arguments 
Set fso = CreateObject("Scripting.FileSystemObject")
answer = MsgBox("Click OK to choose an input file"  & Chr(10) & Chr(13) & "or click Cancel to use the default of servers.txt", 65, "Computer Info Tool")
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
	objExcel.Visible = true
	'objExcel.Workbooks.Add
	Set objWorkbook = objExcel.Workbooks.Add()
	Set objWorksheet = objWorkbook.Worksheets(1)
	objExcel.worksheets(2).delete
	objExcel.worksheets(2).delete
	strOut = ""
	getOSInfo
	

		
		
 Else 
  WScript.Echo "Input file doesnt exist. Please make sure that the servers.txt file exists in the directory you are running this from." 
 End If 
 
Sub getOSInfo()
On Error Resume Next
		objExcel.worksheets(1).Activate
		objExcel.worksheets(1).Name = "Computer Info"
		objExcel.Cells(1, 1).Value = "Computer Name"
		objExcel.Cells(1, 2).Value = "Ping"
		objExcel.Cells(1, 3).Value = "IP from Ping"
		objExcel.Cells(1, 4).Value = "Computer Name from system"
		objExcel.Cells(1, 5).Value = "IP(s) from system"
		objExcel.Cells(1, 6).Value = "Operating System"
		objExcel.Cells(1, 7).Value = "Last Bootup Time"
		objExcel.Cells(1, 8).Value = "Install Date" 
		objExcel.Cells(1, 9).Value = "Manufacturer"
		objExcel.Cells(1, 10).Value = "Serial Number"
		objExcel.Cells(1, 11).Value = "Model"
		objExcel.Cells(1, 12).Value = "LDAP DN"
		objExcel.Cells(1, 13).Value = "Member of Group(s)"
		rowVar=2 

  Set txtStreamIn = fso.OpenTextFile(InputFile) 
  Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	strComputer = ereg_replace(strComputer, "/s", "", True)
If strComputer <> "" Then

Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
    ExecQuery("select * from Win32_PingStatus where address = '"_
        & strComputer & "'")
For Each objStatus in objPing
    If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
    	'WScript.Echo objStatus.StatusCode
        'WScript.Echo("Computer " & strComputer & " did not respond.")
        objExcel.Cells(rowVar, 2).Value = "No response"
        objExcel.Cells(rowVar, 3).Value = objStatus.ProtocolAddress
    Else
    	'WScript.Echo("Computer " & strComputer & " responded on " & objStatus.ProtocolAddress & " in " & objStatus.ResponseTime & "ms.")
    	objExcel.Cells(rowVar, 2).Value = objStatus.ResponseTime & "ms"
    	objExcel.Cells(rowVar, 3).Value = objStatus.ProtocolAddress
    End If
    Next
	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	If Err.Number <> 0 Then
		objExcel.Cells(rowVar, 1).Value = strComputer
		objExcel.Cells(rowVar, 4).Value = "Error # " & CStr(Err.Number) & " " & Err.Description
		printout "Error # " & CStr(Err.Number) & " " & Err.Description
		rowvar = rowvar+1
	    Err.Clear
	Else 
Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Set colBIOS = objWMIService.ExecQuery ("Select * from Win32_BIOS")
Set colComputerSystem = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
Set colNetworkAdapterConfiguration = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration")

For Each objOS in colOperatingSystems
	objExcel.Cells(rowVar, 1).Value = strComputer
	Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
	dtmConvertedDate.Value = objOS.InstallDate
	dtmInstallDate = dtmConvertedDate.GetVarDate
	objExcel.Cells(rowVar, 8).Value = dtmInstallDate
	tempArray = Split(objOS.name, "|")
	objExcel.Cells(rowVar, 6).Value = tempArray(0)
	'dtmBootup = objOS.LastBootUpTime
	'dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
	'objExcel.Cells(rowVar, 7).Value = DateDiff("s", dtmLastBootUpTime, Now)
	dtmConvertedDate.Value = objOS.LastBootUpTime
	dtmBootTime = dtmConvertedDate.GetVarDate
	objExcel.Cells(rowVar, 7).Value = dtmBootTime
Next

For each objBIOS in colBIOS
	objExcel.Cells(rowVar, 10).Value = objBIOS.SerialNumber
Next

For Each objCS In colComputerSystem
    objExcel.Cells(rowVar, 9).Value = objCS.Manufacturer
    objExcel.Cells(rowVar, 11).Value = objCS.Model
    objExcel.Cells(rowVar, 4).Value = objCS.name
Next

For Each objNetAdapter In colNetworkAdapterConfiguration
	ipAddress = objNetAdapter.ipaddress
	For i = 0 To UBound(ipaddress)
		If iplist = "" Then
		iplist = ipaddress(i)
		Else
		iplist = iplist & ", " & ipaddress(i)
		End If
	Next
	objExcel.Cells(rowVar, 5).Value = iplist
	
Next


Err.Clear
strUser = strComputer & "$"
set objRoot = getobject("LDAP://RootDSE")
defaultNC = objRoot.get("defaultnamingcontext")
computerDN = FindUser(strUser, defaultNC)
ouarray = Split(computerDN,",")
For i = 1 To UBound(ouarray)
	If ou = "" Then
	ou = ouarray(i)
	Else 
	ou = ou & "," & ouarray(i)
	End If
Next	
objExcel.Cells(rowVar, 12).Value = ou
	
	set dicSeenGroup = CreateObject("Scripting.Dictionary")
	strGroups = DisplayGroups(computerDN,"",dicSeenGroup)
	aryGroups = Split(strGroups,"CN=")
	strGroups = ""
	For i = 2 To UBound(aryGroups)
		strGroups = strGroups & ", " & aryGroups(i)
	Next
	objExcel.Cells(rowVar, 13).Value = Right(strGroups,Len(strGroups) -2)
	Err.Clear

strOut = ""
iplist = ""
ou = ""
rowvar = rowvar+1
End If
End If

Loop
  			objExcel.Cells.Select
			objExcel.Cells.EntireColumn.AutoFit
			objExcel.Range("B2").Select
			objExcel.ActiveWindow.FreezePanes = True

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

Function FindUser(Byval UserName, Byval Domain) 
	on error resume next

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & Domain & _
			"' WHERE sAMAccountName = '" & UserName & "'"
	
	set rs = cmd.execute

	if err<>0 then
		FindUser="Error connecting to Active Directory Database:" & err.description
		'wscript.quit
	else
		if not rs.BOF and not rs.EOF then
     			rs.MoveFirst
     			FindUser = rs(0)
		else
			FindUser = "Not Found"
		end if
	end if
	cn.close
end Function

Function DisplayGroups ( strObjectADsPath, strSpaces, dicSeenGroup)
   set objObject = GetObject(strObjectADsPath)
   'strOut must be global variable
   strOut = strOut & strSpaces & objObject.Name
   on error resume next ' Doing this to avoid an error when memberOf is empty
   if IsArray( objObject.Get("memberOf") ) then
      colGroups = objObject.Get("memberOf")
   else
      colGroups = Array( objObject.Get("memberOf") )
   end if

   for each strGroupDN In colGroups
      if Not dicSeenGroup.Exists(strGroupDN) then
         dicSeenGroup.Add strGroupDN, 1
         DisplayGroups "LDAP://" & strGroupDN, strSpaces & " ", dicSeenGroup
      end if
   next
Err.Clear
DisplayGroups = strOut
End Function
