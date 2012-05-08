Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
'This script enumerates the AD and creates a record set of all machine accounts
'It then loops through the record set checking the machine accounts password age
'and comparing it to the age that was entered. Any machine account that has a
'password age older than that which is specified will be listed.
'Keep in mind that by default a machine will renew it's password after 30 days so
'if a password is 60 days old then it has not been powered on and attached to the 
'network for at least 30 days.

'The second inputbox is to decide if you only wish to run this in list mode, which
'will only list the machines in a text file but will not take any other action.
'If you select Yes to run in delete mode, then it will still write to the text
'file but it will also prompt you with the machine account name and its password
'age for each machine account that is greater than the age specified and ask if
'you wish to delete it from AD or not.

'Don't forget to replace "yourdom.com" with your domain name.
DomainName = "houston.personix.local"
Do While limit < 60
limit = inputbox("Enter max age in days." & vbCrLf & "Above 60 please.")
On Error Resume Next
If limit < 60 Then If MsgBox("Number of days needs to be higher than 60." & vbCrLf & "Cancel to quit.",vbOKCancel,"Error") = vbcancel Then WScript.Quit
If Err.Number <> 0 Then Err.Clear : If MsgBox("Numbers only please." & vbCrLf & "Cancel to quit.",vbOKCancel,"Error") = vbcancel Then WScript.Quit
If limit = "" Then If MsgBox("You need to put in a limit or you can cancel to exit.",vbOKCancel,"Error") = vbcancel Then WScript.Quit
Loop


'****************Setup Log file******************************************************

Set fso = CreateObject("Scripting.FileSystemObject")
'The 8 in this line will append to an existing file, replace with a 2 to override
set txtStream = fso.OpenTextFile("oldaccounts.csv", ForWriting, True)
txtStream.WriteLine """Name"",""Last Password Set"",""Days Old"",""Address From Ping"",""Ping Result"",""Name From System (from WMI)"""


'****************Setup ADSI connection and populate ADSI Collection******************

Set objADOconnADSI = CreateObject("ADODB.Connection")
objADOconnADSI.Open "Provider=ADsDSOObject;"
Set objCommandADSI = CreateObject("ADODB.Command")
objCommandADSI.ActiveConnection = objADOconnADSI
'there is a 1000 object default if these next 2 lines are omited.
objCommandADSI.Properties("Size Limit")= 10000
objCommandADSI.Properties("Page Size")= 10000
objCommandADSI.Properties("Sort on") = "sAMAccountName"
objCommandADSI.CommandText = "<LDAP://" & DomainName & ">;(objectClass=computer);sAMAccountName,pwdLastSet,name,distinguishedname;subtree"
Set objRSADSI = objCommandADSI.Execute

'Loop through record set and compare password age*************************************

do while NOT objRSADSI.EOF
	if not isnull(objRSADSI.Fields("distinguishedname")) and objRSADSI.Fields("distinguishedname") <> "" then
		objDate = objRSADSI.Fields("PwdLastSet")
		'Go to function to make sense of the PwdLastSet value from AD for the machine account.
		dtmPwdLastSet = Integer8Date(objDate, lngBias)
		'calculate the current age of the password.
		DiffADate = DateDiff("d", dtmPwdLastSet, Now)
		'Is the password older than the specified age.
		if DiffADate > int(limit) then
			'get the ping address
			strPingAddress = pingAddress(objRSADSI.Fields("name"))
			'if there's an address, ping it
			If Not strPingAddress = "No Address" Then strPingResult = pingResult(objRSADSI.Fields("name")) Else strPingResult = "Skipping"
			'if it responds, try to connect to WMI and ask it its name
			If Not strPingResult = "Skipping" Then strWMIComputerName = getWMIComputerName(strPingAddress) Else strWMIComputerName = "Skipping"
			txtStream.WriteLine """" & objRSADSI.Fields("name") & """,""" & dtmPwdLastSet & """,""" & DiffADate & """,""" & strPingAddress & """,""" & strPingResult & """,""" & strWMIComputerName & """"
		end if
	end if
	objRSADSI.MoveNext
Loop
txtstream.Close
ToAddress = "zeitz, scott"
MessageSubject = "old computer account (auto email from script)"
MessageBody = "testing this script that uses outlook to email"
MessageAttachment = "h:\oldaccounts.csv"

' connect to Outlook
Set ol = WScript.CreateObject("Outlook.Application")
Set ns = ol.getNamespace("MAPI")

Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf

' validate the recipient, just in case...
Set myRecipient = ns.CreateRecipient(ToAddress)
myRecipient.Resolve
If Not myRecipient.Resolved Then
  MsgBox "Unknown recipient"
Else
  newMail.Recipients.Add(ToAddress)
  newMail.recipients.add("becker, tommy")
  newMail.Attachments.Add(MessageAttachment).Displayname = "Check this out"
  newMail.Send
End If

Set ol = Nothing

wscript.echo "Done!"

'I found this function and it seems to work greate. I don't pretend to fully understand it though.
'I don't know who wrote it or I would give them credit.
Function Integer8Date(objDate, lngBias)
' Function to convert Integer8 (64-bit) value to a date, adjusted for
' local time zone bias.
Dim lngAdjust, lngDate, lngHigh, lngLow
lngAdjust = lngBias
lngHigh = objDate.HighPart
lngLow = objdate.LowPart
' Account for bug in IADslargeInteger property methods.
If lngLow < 0 Then
lngHigh = lngHigh + 1
End If
If (lngHigh = 0) And (lngLow = 0) Then
lngAdjust = 0
End If
lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
+ lngLow) / 600000000 - lngAdjust) / 1440
Integer8Date = CDate(lngDate)
End Function

Function pingAddress(strComputer)
On Error Resume Next
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
    ExecQuery("select * from Win32_PingStatus where address = '"_
        & strComputer & "'")
For Each objStatus in objPing
    If Err.Number <> 0 Then WScript.Echo "Error 0x" & Hex(Err.Number) & ":" & Err.Description & " occured connecting to " & strComputer : Err.Clear : pingAddress = "No Address" : Exit Function
    If objStatus.protocoladdress = "" Then
        pingAddress = "No Address"
    Else
        pingAddress = objStatus.protocoladdress
    End If
Next
End Function

Function pingResult(strComputer)
On Error Resume Next
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
    ExecQuery("select * from Win32_PingStatus where address = '"_
        & strComputer & "'")
For Each objStatus in objPing
    If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
        pingResult = "No response"
    Else
        pingResult = objStatus.ResponseTime & "ms"
    End If
Next
End Function

Function getWMIComputerName(strComputer)
	On Error Resume Next
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    If Err.Number <> 0 Then getWMIComputerName = "Error 0x" & Hex(Err.Number) & ":" & Err.Description & " occured connecting to " & strComputer : Err.Clear : Exit Function
	Set colComputerSystem = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
		For Each objCS In colComputerSystem
    		getWMIComputerName = objCS.name
		Next
End Function