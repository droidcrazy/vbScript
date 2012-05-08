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

Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 
objExcel.Workbooks.Add
objExcel.worksheets(2).delete
objExcel.worksheets(2).delete


arrComputers = Split(strComputers, " ") 


rowvar = 2
counter = 1
  Set txtStreamIn = fso.OpenTextFile(InputFile) 
Do While Not (txtStreamIn.AtEndOfStream) 
	strComputer = txtStreamIn.ReadLine 
	If strComputer <> "" Then

'	objExcel.worksheets(counter).Activate
'	objExcel.worksheets(counter).Name = strComputer
objExcel.Cells(1, 1).Value = "Computer Name" 
objExcel.Cells(1, 2).Value = "OS Version" 
objExcel.Cells(1, 3).Value = "Service Pack" 
objExcel.Cells(1, 4).Value = "# of Procs" 
objExcel.Cells(1, 5).Value = "Proc Type" 
objExcel.Cells(1, 6).Value = "Max Clock Speed" 
objExcel.Cells(1, 7).Value = "Tot Phy Mem" 
objExcel.Cells(1, 8).Value = "Free Phy Mem" 
objExcel.Cells(1, 9).Value = "Tot Virtual Mem" 
objExcel.Cells(1, 10).Value = "Free Virtual Mem" 
objExcel.Cells(1, 11).Value = "Tot Visible Mem" 
objExcel.Cells(1, 12).Value = "Domain" 
objExcel.Cells(1, 13).Value = "Domain Role" 
objExcel.Cells(1, 14).Value = "Manufacturer" 
objExcel.Cells(1, 15).Value = "Model" 
objExcel.Cells(1, 16).Value = "Serial Number" 
objExcel.Cells(1, 17).Value = "User" 



' 
'===================================================================== 
' Insert your code here 
' 
'===================================================================== 
objExcel.Cells(rowvar, 1).Value = strComputer

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
Set colItems = objWMIService.ExecQuery _ 
("Select * From Win32_OperatingSystem") 
For Each objItem in ColItems 
Wscript.Echo strComputer & ": " & objItem.Caption 
objExcel.Cells(rowvar, 2).Value = objItem.Caption 
Next 
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem") 
For Each objOS in colOSes 


WScript.Echo "Service Pack: " & objOS.ServicePackMajorVersion & "." & _ 
objOS.ServicePackMinorVersion 
objExcel.Cells(rowvar, 3).Value = objOS.ServicePackMajorVersion & "." & _
objOS.ServicePackMinorVersion 
Next 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colCSes = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
For Each objCS In colCSes 
WScript.Echo "Number Of Processors: " & objCS.NumberOfProcessors 
objExcel.Cells(rowvar, 4).Value = objCS.NumberOfProcessors 
Next 
Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor") 
For Each objProcessor in colProcessors 
WScript.Echo "Name: " & objProcessor.Name 
WScript.Echo "Maximum Clock Speed: " & objProcessor.MaxClockSpeed 
objExcel.Cells(rowvar, 5).Value = objProcessor.Name 
objExcel.Cells(rowvar, 6).Value = objProcessor.MaxClockSpeed 
Next 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colCSItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
For Each objCSItem In colCSItems 
WScript.Echo "Total Physical Memory: " & objCSItem.TotalPhysicalMemory 
objExcel.Cells(rowvar, 7).Value = objCSItem.TotalPhysicalMemory 
Next 
Set colOSItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem") 
For Each objOSItem In colOSItems 
WScript.Echo "Free Physical Memory: " & objOSItem.FreePhysicalMemory 
WScript.Echo "Total Virtual Memory: " & objOSItem.TotalVirtualMemorySize 
WScript.Echo "Free Virtual Memory: " & objOSItem.FreeVirtualMemory 
WScript.Echo "Total Visible Memory Size: " & objOSItem.TotalVisibleMemorySize 
objExcel.Cells(rowvar, 8).Value = objOSItem.FreePhysicalMemory 
objExcel.Cells(rowvar, 9).Value = objOSItem.TotalVirtualMemorySize 
objExcel.Cells(rowvar, 10).Value = objOSItem.FreeVirtualMemory 
objExcel.Cells(rowvar, 11).Value = objOSItem.TotalVisibleMemorySize 
Next 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
For Each objItem In colItems 
WScript.Echo "Domain: " & objItem.Domain 
objExcel.Cells(rowvar, 12).Value = objItem.Domain 
Select Case objItem.DomainRole 
Case 0 strDomainRole = "Standalone Workstation" 
Case 1 strDomainRole = "Member Workstation" 
Case 2 strDomainRole = "Standalone Server" 
Case 3 strDomainRole = "Member Server" 
Case 4 strDomainRole = "Backup Domain Controller" 
Case 5 strDomainRole = "Primary Domain Controller" 
End Select 
WScript.Echo "Domain Role: " & strDomainRole 
objExcel.Cells(rowvar, 13).Value = strDomainRole 
strRoles = Join(objItem.Roles, ",") 


Next 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 
For Each objItem In colItems 
WScript.Echo "Manufacturer: " & objItem.Manufacturer 
WScript.Echo "Model: " & objItem.Model 
objExcel.Cells(rowvar, 14).Value = objItem.Manufacturer 
objExcel.Cells(rowvar, 15).Value = objItem.Model 
Next 
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime") 
Set objWMIService = GetObject("winmgmts:" _ 
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 


Set colOperatingSystems = objWMIService.ExecQuery _ 
("Select * from Win32_OperatingSystem") 


For Each objOperatingSystem in colOperatingSystems 
Wscript.Echo "Serial Number: " & objOperatingSystem.SerialNumber 
objExcel.Cells(rowvar, 16).Value = objOperatingSystem.SerialNumber 


Next 
Set objWMIService = GetObject("winmgmts:" _ 
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 


Set colComputer = objWMIService.ExecQuery _ 
("Select * from Win32_ComputerSystem") 


For Each objComputer in colComputer 
Wscript.Echo "Logged-on user: " & objComputer.UserName 
objExcel.Cells(rowvar, 17).Value = objComputer.UserName 
Next 
	objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit
    objExcel.Range("B2").Select
    objExcel.ActiveWindow.FreezePanes = True
		'counter = counter +1
		'set objWorksheet = objExcel.Sheets.Add( , objExcel.WorkSheets(objExcel.WorkSheets.Count))
End If
rowvar = rowvar+1
Loop
'objExcel.worksheets(counter).delete
objExcel.worksheets(1).Activate