Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const AGING_TOKEN = "[Aging:"

Const DDNS_NO_REFRESH = 7 ' The dynamic DNS no refresh period, where an update classified as a refresh will not be accepted for the record
Const DDNS_REFRESH = 7 ' The dynamic DNS refresh period, during which an update will be accepted for the record
Const GMT_OFFSET = +6 ' Offset in hours to adjust the resultant times based on the current GMT timezone

Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count = 1 Then 
strFileName = WScript.Arguments(0) 
Else
wscript.echo "Specify a filename containing the output of dnscmd. eg DNSScavengeTest.vbs DNS_Records.txt"
wscript.quit(2)
End If

If Not objFSO.FileExists(strFileName) Then
WScript.Echo "Error: " & strFileName & " file not found."
wscript.quit(2)
End If

Set objTextStream = objFSO.OpenTextFile(strFileName, ForReading)
strZoneRecords = objTextStream.ReadAll

For Each strLine in Split(strZoneRecords, vbCRLF)
intStart = InStr(1, strLine, AGING_TOKEN, 1) 
If intStart <> 0 Then ' Does this line contain an aging value?
intStart = intStart + Len(AGING_TOKEN)
intEnd = InStr(intStart, strLine, "]")
If intEnd <> 0 Then intLength = intEnd - intStart
strHost = Right (strline, Len(strline) - InStrRev(strline, vbTab)) ' Yes, extract the host

intAging = Mid(strLine, intStart, intLength) ' Extract the aging value, expressed in the decimal number of hours since 01/01/1601

dtmDate = DateAdd("h", intAging, "01/01/1601 00:00:00 AM") ' Convert to a date timestamp
dtmDate = DateAdd("h", GMT_OFFSET, dtmDate) ' Add the current GMT offset

intDiff = DateDiff("h", dtmDate, Now) ' The difference between now and the timestampe
intHourDiff = intDiff - ((DDNS_NO_REFRESH * 24) + (DDNS_REFRESH * 24)) ' Based on the dynamic DNS no-refresh and refresh periods combined
If intHourDiff > 0 Then ' Is this a positive number, indicating the record will be scavenged
intDay = CInt(intHourDiff / 24) ' Yes, convert to a number of days for output
WScript.Echo strHost & ", " & dtmDate & ", " & intDay + DDNS_NO_REFRESH + DDNS_REFRESH & " days ago" 'This record would be scavenged
Else
WScript.Echo "*" & strHost & ", " & dtmDate ' This record won't be scavenged
intDay = 0
End If
End If
Next
