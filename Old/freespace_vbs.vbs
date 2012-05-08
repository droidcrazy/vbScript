' FreeSpace.vbs,  Version 1.00
' Display free disk space for all local drives.
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com


' Check command line parameters
Select Case WScript.Arguments.Count
	Case 0
		' Default if none specified is local computer (".")
		Set objWMIService = GetObject( "winmgmts://./root/cimv2" )
		Set colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
		For Each objItem in colItems
			strComputer = objItem.Name
		Next
	Case 1
		' Command line parameter can either be a computer name
		' or "/?" to request online help
		strComputer = Wscript.Arguments(0)
		if InStr( strComputer, "?" ) > 0 Then Syntax
	Case Else
		' Maximum is 1 command line parameter
		Syntax
End Select

Display( strComputer )
WScript.Quit(0)


Function Display( strComputer )
	strMsg = vbCrLf & "Name:" & vbTab & "Drive:" & vbTab & "Size:" & _
	         vbTab & "Free:" & vbTab & "% Free:" & vbCrLf & "=====" & _
	         vbTab & "======" & vbTab & "=====" & vbTab & "=====" & _
	         vbTab & "=======" & vbCrLf
	On Error Resume Next
	Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
	If Err.Number Then
		WScript.Echo vbCrLf & "Error # " & CStr( Err.Number ) & _
		             " " & Err.Description
		Err.Clear
		Syntax
	End If
	On Error GoTo 0
	' Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where MediaType=12",,48)
	Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3",,48)
	For Each objItem in colItems
		strMsg = strMsg & strComputer & vbTab & _
		         objItem.Name & vbTab & _
		         CStr( Int( 0.5 + ( objItem.Size / 1073741824 ) ) ) & _
		         vbTab & _
		         CStr( Int( 0.5 + ( objItem.FreeSpace / 1073741824 ) ) ) & _
		         vbTab & _
		         CStr( Int( 0.5 + ( 100 * objItem.FreeSpace / objItem.Size) ) ) & _
		         vbCrLf
	Next
	WScript.Echo strMsg
End Function


Sub Syntax
	strMsg = vbCrLf & "FreeSpace.vbs,  Version 1.00" & vbCrLf & _
	         "Display free disk space for all local drives." & vbCrLf & _
	         vbCrLf & _
	         "Usage:  CSCRIPT  FREESPACE.VBS  [ computer_name ]" & _
	         vbCrLf & vbCrLf & _
	         "Where:  " & Chr(34) & "computer_name" & Chr(34) & _
	         " is the name of a WMI enabled computer on the network" & _
	         vbCrLf & vbCrLf & _
	         "Written by Rob van der Woude" & vbCrLf & _
	         "http://www.robvanderwoude.com" & vbCrLf
	WScript.Echo strMsg
	WScript.Quit(1)
End Sub
