Option Explicit


Dim strComputer, member
Dim grp, obj
On Error Resume Next
strComputer = "."

Set grp = GetObject("WinNT://" & strComputer & "/Administrators,group")

For Each member In grp.Members
	WScript.Echo member.Name & " (" & member.Class & ")"

	If (member.Class = "Group") Then
		For Each obj In member.Members
			WScript.Echo " " & obj.adspath & " (" & obj.Class & ")"
			WScript.Echo " Last Logon: " & obj.LastLogin
		Next
	End If

Next
