On Error Resume Next
Set autoUpdateClient = CreateObject("microsoft.Update.AutoUpdate")
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'")
If Err.Number <> 0 Then WScript.Echo(Err.Number & ": Description:" & Err.Description):Err.Clear

autoUpdateClient.detectnow()
'------------------------------------------------------------------------------------
'report missing updates:
WScript.Echo("Missing " & searchResult.Updates.count & " updates:")
'If searchResult.Updates.count = 0 Then WScript.Quit
Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1 
	Set update = searchResult.Updates.Item(I) 
 	If Not update.EulaAccepted Then update.AcceptEula 
	If Not update.isdownloaded Then updatesToDownload.Add(update):WScript.echo("Update added to download list: " & update.Title) 
Next 
if err.number <> 0 then	WScript.Echo("Error " & err.number & " has occured.  Error description: " & err.description):Err.Clear
On Error Goto 0
'------------------------------------------------------------------------------------
'download missing updates
Set downloader = updateSession.CreateUpdateDownloader() 

downloader.Updates = updatesToDownload
WScript.Echo("********** Downloading updates **********")

'Dim DLjob 

'Set DLjob = downloader.BeginDownload() 'this and the loop below is what i want to do...
'more info at http://forums.techarena.in/showthread.php?t=359040
'more info at http://msdn.microsoft.com/en-us/library/aa386132(VS.85).aspx
'more info at http://www.swissdelphicenter.ch/de/forum/index.php/topic,10601.0.html

'Do While NOT DLjob.iscompleted
'	WScript.Echo dljob.getprogress.percentcomplete & "% complete."
'	WScript.Sleep 5000
'Loop
Set dlsink=wscript.CreateObject("WBemScripting.SWbemSink","DLSINK_")

downloader.begindownload dlsink,dlsink




For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    If update.IsDownloaded Then
       WScript.Echo "downloaded: " & update.Title
    End If
       On Error GoTo 0
Next

Sub DLSINK_onProgressChanged(objEvent,objContext)
WScript.Echo "called onprogresschanged"
End Sub

Sub DLINK_onCompleted(objEvent,objContext)
WScript.Echo "called oncompleted"
End Sub

