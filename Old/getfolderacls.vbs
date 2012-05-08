' Read a Folder DACL



strFolderName = "z:\pricing db"

Dim Perms_LStr, Perms_SStr, Perms_Const
    'Permission LongNames
    Perms_LStr=Array("Synchronize"			, _
		"Take Ownership"					, _
		"Change Permissions"				, _
		"Read Permissions"					, _
		"Delete"							, _
		"Write Attributes"					, _
		"Read Attributes"					, _
		"Delete Subfolders and Files"			, _
		"Traverse Folder / Execute File"		, _
		"Write Extended Attributes"			, _
		"Read Extended Attributes"			, _
		"Create Folders / Append Data"		, _
		"Create Files / Write Data"			, _
		"List Folder / Read Data"	)
    'Permission Single Character codes
    Perms_SStr=Array("E"		, _
		"D"		, _
		"C"		, _
		"B"		, _
		"A"		, _
		"9"		, _
		"8"		, _
		"7"		, _
		"6"		, _
		"5"		, _
		"4"		, _
		"3"		, _
		"2"		, _
		"1"		)
    'Permission Integer
    Perms_Const=Array(&H100000	, _
		&H80000		, _
		&H40000		, _
		&H20000		, _
		&H10000		, _
		&H100		, _
		&H80		, _
		&H40		, _
		&H20		, _
		&H10		, _
		&H8			, _
		&H4			, _
		&H2			, _
		&H1		)

SE_DACL_PRESENT = &h4
ACCESS_ALLOWED_ACE_TYPE = &h0
ACCESS_DENIED_ACE_TYPE  = &h1

FILE_ALL_ACCESS         = &h1f01ff
FOLDER_ADD_SUBDIRECTORY = &h000004
FILE_DELETE             = &h010000
FILE_DELETE_CHILD       = &h000040
FOLDER_TRAVERSE         = &h000020
FILE_READ_ATTRIBUTES    = &h000080
FILE_READ_CONTROL       = &h020000
FOLDER_LIST_DIRECTORY   = &h000001
FILE_READ_EA            = &h000008
FILE_SYNCHRONIZE        = &h100000
FILE_WRITE_ATTRIBUTES   = &h000100
FILE_WRITE_DAC          = &h040000
FOLDER_ADD_FILE         = &h000002
FILE_WRITE_EA           = &h000010
FILE_WRITE_OWNER        = &h080000

Set objWMIService = GetObject("winmgmts:")
Set objFolderSecuritySettings = _
objWMIService.Get("Win32_LogicalFileSecuritySetting='" & strFolderName & "'")
intRetVal = objFolderSecuritySettings.GetSecurityDescriptor(objSD)

intControlFlags = objSD.ControlFlags

If intControlFlags AND SE_DACL_PRESENT Then
   arrACEs = objSD.DACL
   For Each objACE in arrACEs
   	If objACE.aceflags And 16 Then
   	Else
   	If acllist = "" Then
		acllist = "[" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name
	Else
		acllist = acllist & ", " & objACE.Trustee.Domain & "\" & objACE.Trustee.Name
	End If	
      If objACE.AceType = ACCESS_ALLOWED_ACE_TYPE Then
         acllist = acllist & " - Allowed:" & objACE.aceflags & ";" & SECString(objACE.AccessMask,true)
         'WScript.Echo objACE.aceflags
      ElseIf objACE.AceType = ACCESS_DENIED_ACE_TYPE Then
         acllist = acllist & " - Denied:" & objACE.aceflags & ";" & objACE.AccessMask
      End If
      End If
   Next
   If acllist <> "" Then acllist = acllist & "]"
   WScript.Echo acllist
Else
   WScript.Echo "No DACL present in security descriptor"
End If

Function SECString(byval intBitmask, byval ReturnLong)
'debug_on = true
    On Error Resume Next
    Dim LongName, X

    If debug_on then wscript.echo("SECString: Enter")

    SECString = ""

    Do
	If debug_on then wscript.echo("SECString: intBitmask = " & intBitmask)
		
	For X = LBound(Perms_LStr) to UBound(Perms_LStr)
    		If ((intBitmask And Perms_Const(X)) = Perms_Const(X)) then
			If Perms_SStr(X) <> "" then
				SECString = SECString & Perms_SStr(X)
			End if
    		End if
	Next

	If debug_on then wscript.echo("SECString: SECString = " & SECString)
	Select Case SECString
	Case "DCBA987654321", "EDCBA987654321"
		SECString = "F"								'Full control
		LongName = "Full Control"	
	Case "BA98654321", "EBA98654321"
		SECString = "M"								'Modify
		LongName = "Modify"
	Case "B98654321", "EB98654321"
		SECString = "XW"								'Read, Write and Execute
		LongName = "Read, Write and Execute"
	Case "B9854321", "EB9854321"
		SECString = "RW"								'Read and Write
		LongName = "Read and Write"
	Case "B8641", "EB8641"
		SECString = "X"								'Read and Execute
		LongName = "Read and Execute"
	Case "B841", "EB841"
		SECString = "R"								'Read
		LongName = "Read"
	Case "9532", "E9532"
		SECString = "W"								'Write
		LongName = "Write"
	Case Else
If SECString = "" Then
		Select Case intBitmask
		Case "-1610612736"							'custom Read and Exceute
			SECString = "X"
			LongName = "Read and Execute"
		Case Else	
			LongName = "Unknown (" & intBitmask & ")"
		End Select		
		Else
			If LEN(SECString) = 1 then
				For X = LBound(Perms_SStr) to UBound(Perms_SStr)
					If StrComp(SECString,Perms_SStr(X),1) = 0 Then
						LongName = "Advanced (" & Perms_LStr(X) & ")"
						Exit For
					End if
				Next
			Else
				LongName = "Special (" & SECString & ")"
			End if
		End if
	End Select

	Exit Do
    Loop

    If ReturnLong Then SECString = LongName

    If debug_on then wscript.echo("SECString: Return = " & SECString)

    Call blnErrorOccurred(" occurred while in the SECString routine. (Msg#2001)")
    If debug_on then wscript.echo("SECString: Exit")

End Function