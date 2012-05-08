On Error Resume Next
version = "1.06"
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const ForWriting = 2
Const ForReading = 1
Const ForAppending = 8

Set autoUpdateClient = CreateObject("microsoft.Update.AutoUpdate")
Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'")
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
Set objADInfo = CreateObject("ADSystemInfo")
Set WshShell = WScript.CreateObject("WScript.Shell")
'Set WshSysEnv = WshShell.Environment("PROCESS")
Set ws = wscript.CreateObject("Scripting.FileSystemObject")

'Script Configuration----------------------------------------------------
'------------------------------------------------------------------------
scriptroot = "\\hou20017\batch$"
strDateStamp =Year(Now) & Right(100 + Month(Now), 2) & Right (100 + Day(Now), 2)
logfile = scriptroot & "\log\MissingSecAndCritUpdates" & strDateStamp & ".log"
strComputer1 = objADInfo.ComputerName
If strComputer = "" Then strComputer = wshShell.ExpandEnvironmentStrings("%Computername%")
If InStr(ucase(WScript.FullName),"CSCRIPT.EXE") Then blnCScript = TRUE Else blnCScript = False
OODList = scriptroot & "\log\OutOfDateSecAndCritUpdates" & strDateStamp & ".txt"
OODCompList = scriptroot & "\log\OutOfDateSecAndCritComputers" & strDateStamp & ".txt"
OODcheck = False
ComputerOOD = False
OODUpdates = 0


'End Script Configuration------------------------------------------------
'------------------------------------------------------------------------
Set l = ws.OpenTextFile (logfile, ForAppending, True)
Set OODfile = ws.OpenTextFile (OODList, ForAppending, True)
Set OODcompfile = ws.OpenTextFile (OODCompList, ForAppending, True)
If Err.Number <> 0 Then WriteLog(Err.Number & ": Description:" & Err.Description)
Err.Clear
autoUpdateClient.detectnow()

If searchResult.Updates.count = 0 Then
WriteLog("Up to date.")
Else
WriteLog("Missing " & searchResult.Updates.count & " updates.")
End If

'Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

objConnection.Open _
	"Provider=SQLOLEDB;Data Source=houmwinetop01;" & _
		"Trusted_Connection=Yes;Initial Catalog=SUSDB;"


For I = 0 to searchResult.Updates.Count-1 
	Set update = searchResult.Updates.Item(I) 
	'strUpdates = strUpdates & update.Title
	'objRecordSet.Open "SELECT UpdateId,CreationDate FROM PUBLIC_VIEWS.vUpdate where updateid='" & update.identity.updateid & "'", _
    '    objConnection, adOpenStatic, adLockOptimistic
    strSQL = "Declare @updateid varchar(100);Set @updateid = '" & update.identity.updateid & "';" & _
		"SELECT UpdateId,CreationDate,InstallationRebootBehavior FROM PUBLIC_VIEWS.vUpdate Where updateid = @updateid"
			objRecordSet.Open strSQL, objConnection, adOpenStatic, adLockOptimistic

	releasedate = objRecordSet.Fields("CreationDate").Value
	rebootbehavior = objRecordSet.Fields("InstallationRebootBehavior").Value
	On Error Goto 0
		For counter = 0 To update.categories.count -1
		If category = "" Then
		category = update.categories.item(counter).name
		Else
		category = category & "; " & update.categories.item(counter).name
		End If
		checkcat = update.categories.item(counter).name
		If checkcat = "Security Updates" Or checkcat = "Critical Updates" Then categorymatch = True
		Next
	If checkOOD(releasedate) And categorymatch Then 
	OODfile.WriteLine("[" & time & "] - " & strComputer & ",script version: " & version & "," & """Missing: " & update.Title & """, Update ID:" & update.identity.updateid & ", Release Date:" & releasedate & ", Reboot Behavior:" & rebootbehavior)
	OODupdates = OODupdates +1
	End If
	If checkOOD(releasedate) And categorymatch Then ComputerOOD = True
	WriteLog("""Missing: " & update.Title & """, Update ID:" & update.identity.updateid & ", Release Date:" & releasedate & ", Reboot Behavior:" & rebootbehavior)
	If Not update.EulaAccepted Then update.AcceptEula

	objRecordset.Close
	category = ""
	categorymatch = False
Next
If ComputerOOD Then OODfile.WriteLine("[" & time & "] - " & strComputer & ",script version: " & version & "," & "This computer has " & OODupdates & " updates out of date.")
If ComputerOOD Then OODcompfile.WriteLine(strComputer & ",""" & OODupdates & " updates out of date.""") Else OODcompfile.WriteLine(strComputer & ","" Up to date.""")
objConnection.Close
Set autoUpdateClient = Nothing
Set updateSession = Nothing
Set objConnection = Nothing
Set objRecordSet = Nothing
Set objADInfo = Nothing

Function WriteLog(strMsg) 
l.writeline "[" & time & "] - " & strComputer & ",script version: " & version & "," & strMsg
' Output to screen if cscript.exe 
If blnCScript Then WScript.Echo "[" & time & "] " & strMsg 
End Function

Function checkOOD(releasedate)
	d = CDate(releasedate)
	date0 = DateAdd("m",-3,Now)
	If d < date0 Then
	checkOOD = True
	Else
	checkOOD = False
	End If
End Function