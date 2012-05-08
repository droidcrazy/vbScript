On Error Resume Next
Set autoUpdateClient = CreateObject("microsoft.Update.AutoUpdate")
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'")

WScript.Echo(autoupdateclient.Results.LastInstallationSuccessDate)
autoUpdateClient.detectnow()
If Err.Number <> 0 Then WScript.Echo(Err.Number & ": Description:" & Err.Description):WScript.Quit
'------------------------------------------------------------------------------------
'report missing updates:
WScript.Echo("Missing " & searchResult.Updates.count & " updates:")

If searchResult.Updates.Count = 0 Then
	WScript.Echo  "There are no further updates needed for your PC at this time."
	wscript.quit
End If

Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1 
	Set update = searchResult.Updates.Item(I) 
	strUpdates = strUpdates & update.Title
 	WScript.echo("Update to be added to download list: " & update.Title) 
	If Not update.EulaAccepted Then update.AcceptEula 
	updatesToDownload.Add(update)
	WScript.Echo("Added.")
Next 

'------------------------------------------------------------------------------------
'download missing updates
Set downloader = updateSession.CreateUpdateDownloader() 
on error resume next
downloader.Updates = updatesToDownload
WScript.Echo("********** Downloading updates **********")

Set DLjob = downloader.BeginDownload()

Do While NOT DLjob.completed
	WScript.Echo dljob.getprogress.percentcomplete & "% complete."
	WScript.Sleep 5000
Loop

if err.number <> 0 then
	WScript.Echo("Error " & err.number & " has occured.  Error description: " & err.description)
End if


For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    If update.IsDownloaded Then
       WScript.Echo "downloaded: " & update.Title
    End If
       On Error GoTo 0
Next

'------------------------------------------------------------------------------------
'install missing updates
Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
Set installer = updateSession.CreateUpdateInstaller()
WScript.Echo("********** Adding updates to collection **********")
For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
    If update.IsDownloaded = true Then
       updatesToInstall.Add(update)
    End If
       WScript.Echo("Adding to collection: " & update.Title)
Next

installer.Updates = updatesToInstall
WScript.Echo("********** Installing updates **********")

on error resume next	
	Set installationResult = installer.Install()
WScript.Echo("Installation Result: " & installationResult.ResultCode)
WScript.Echo("Reboot Required: " & installationResult.RebootRequired)
WScript.Echo("Listing of updates installed and individual installation results:")
For i = 0 to updatesToInstall.Count - 1
		If installationResult.GetUpdateResult(i).ResultCode = 2 Then 
			strResult = "Installed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 1 Then 
			strResult = "In progress"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 3 Then 
			strResult = "Operation complete, but with errors"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 4 Then 
			strResult = "Operation failed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 5 Then 
			strResult = "Operation aborted"			
		End If
		WScript.Echo(updatesToInstall.Item(i).Title & ": " & strResult)
	Next

WScript.Echo("********** Rebooting Computer **********")
If installationResult.RebootRequired Then
WScript.Echo("********** Rebooting Computer **********")
strComputer = "."
Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem"_
	& " where Primary=true")

Const EWX_LOGOFF = 0 
Const EWX_SHUTDOWN = 1 
Const EWX_REBOOT = 2 
Const EWX_FORCE = 4 
Const EWX_POWEROFF = 8 

For each OpSys in OpSysSet 
	opSys.win32shutdown EWX_REBOOT + EWX_FORCE
Next 
Else
WScript.Echo("********** No Reboot Required **********")
End If 