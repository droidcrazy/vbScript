Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher
intHistoryCount = objSearcher.GetTotalHistoryCount
Set WshNetwork = CreateObject("wscript.network")
strComputer = WshNetwork.ComputerName
Set colHistory = objSearcher.QueryHistory(1, intHistoryCount)

For Each objEntry in colHistory
	Select Case objEntry.Operation
	Case 1
		Operation = "Installation"
	Case 2
		Operation = "Uninstallation"
	Case Else
		Operation = "Unknown"
	End Select
	
	Select Case objEntry.ResultCode
	Case 0
		ResultCode = "Not Started"
	Case 1
		ResultCode = "In Progress"
	Case 2
		ResultCode = "Suceeded"
	Case 3
		ResultCode = "Suceeded with errors"
	Case 4
		ResultCode = "Failed"
	Case 5
		ResultCode = "Aborted"
	Case Else
		ResultCode = "Unknown"
	End Select
	
	Select Case objEntry.ServerSelection
	Case 0
		ServerSelection = "Default"
	Case 1
		ServerSelection = "Managed Server"
	Case 2
		ServerSelection = "Windows Update"
	Case 3
		ServerSelection = "Others"
	Case Else
		ServerSelection = "Unknown"
	End Select
	WScript.Echo(strComputer & "," & objEntry.Date & "," & Operation & "," & ResultCode & ",""" & objEntry.Title & """," & ServerSelection)

Next

