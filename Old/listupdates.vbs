' List Update History
On Error Resume Next

Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher
intHistoryCount = objSearcher.GetTotalHistoryCount

Set colHistory = objSearcher.QueryHistory(1, intHistoryCount)

For Each objEntry in colHistory
	WScript.Echo objEntry.Date & " -- Name: " & objentry.Title
	'Wscript.Echo "Operation: " & objEntry.Operation
    'Wscript.Echo "Result code: " & objEntry.ResultCode
    'Wscript.Echo "Exception: " & objEntry.Exception
    'Wscript.Echo "Date: " & objEntry.Date
    'Wscript.Echo "Title: " & objEntry.Title
    'Wscript.Echo "Description: " & objEntry.Description
    'Wscript.Echo "Unmapped exception: " & objEntry.UnmappedException
    'Wscript.Echo "Client application ID: " & objEntry.ClientApplicationID
    'Wscript.Echo "Server selection: " & objEntry.ServerSelection
    'Wscript.Echo "Service ID: " & objEntry.ServiceID
    'i = 1
    'For Each strStep in objEntry.UninstallationSteps
    '    Wscript.Echo i & " -- " & strStep
    '    i = i + 1
    'Next
    'Wscript.Echo "Uninstallation notes: " & objEntry.UninstallationNotes
    'Wscript.Echo "Support URL: " & objEntry.SupportURL
    'Wscript.Echo
Next