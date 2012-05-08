'region Script Settings
'<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
'  <ScriptPackager>
'    <process>cscript.exe</process>
'    <arguments />
'    <extractdir>%TEMP%</extractdir>
'    <files />
'    <usedefaulticon>true</usedefaulticon>
'    <showinsystray>false</showinsystray>
'    <altcreds>false</altcreds>
'    <efs>true</efs>
'    <ntfs>true</ntfs>
'    <local>false</local>
'    <abortonfail>true</abortonfail>
'    <product />
'    <version>1.0.0.1</version>
'    <versionstring />
'    <comments />
'    <includeinterpreter>false</includeinterpreter>
'    <forcecomregistration>false</forcecomregistration>
'    <consolemode>false</consolemode>
'    <EnableChangelog>false</EnableChangelog>
'    <AutoBackup>false</AutoBackup>
'    <snapinforce>false</snapinforce>
'    <snapinshowprogress>false</snapinshowprogress>
'    <snapinautoadd>0</snapinautoadd>
'    <snapinpermanentpath />
'  </ScriptPackager>
'</ScriptSettings>
'endregion

' BEGIN CALLOUT A
' Establish the AD connection.
Set Conn = CreateObject("ADODB.Connection")
conn.Provider = "ADSDSOObject"
conn.Open "ADs Provider"
' END CALLOUT A

Dim RS
Dim query
baseDN = "dc=houston,dc=personix,dc=local"
' BEGIN CALLOUT B
' Set up the query string and execute it.
query = "<LDAP://" & baseDN & _
">;(&(ObjectClass=User)(userAccountControl=*));" & _
"distinguishedName,samAccountName,userAccountControl;subtree"
' END CALLOUT B
Set RS = conn.Execute(CStr(query))

' Create the output file for writing.
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set fsOut = fs.OpenTextFile(".\userStatusReport.txt", ForWriting, True)

' Output the header for output file.
WScript.Echo Chr(34) & "dn" & Chr(34) & "," & _
Chr(34) & "username" & Chr(34) & "," & _
Chr(34) & "status" & Chr(34) 

' BEGIN CALLOUT C
Dim outputStr
' Loop through the result set and output to file.
While Not RS.EOF
	' Create the output string by first retrieving the DN and 
	' username values from the result set.
	outputStr = Chr(34) & RS("distinguishedName") & Chr(34) & "," & _
	Chr(34) & RS("samAccountName") & Chr(34) & "," 
	
	Dim accountControl, status
	' Get the value for userAccountControl, convert to a 
	' Long because the value of userAccountControl
	' might be greater than the maximum value for an Int.
	accountControl = CLng(RS("userAccountControl"))
	
	' Perform bitwise AND operation on the userAccountControl 
	' value to determine the state.
	If (accountControl And ACCOUNTDISABLE) Then
		status = "Disabled"
	Else
		If (accountControl And LOCKOUT) Then
			status = "Locked Out"
		Else
			status = "Active"
		End If
	End If
	
	' Write the output string to a file.
	outputStr = outputStr & Chr(34) & status & Chr(34)
	WScript.Echo outputStr
	RS.MoveNext
Wend
' END CALLOUT C

' Close the output file.
'fsOut.Close


