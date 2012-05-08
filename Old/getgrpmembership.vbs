	' This code displays the group membership of a user.
	' It avoids infinite loops due to circular group nesting by
	' keeping track of the groups that have already been seen.
	' ------ SCRIPT CONFIGURATION ------
	strUserDN = "<UserDN>" ' e.g. cn=jsmith,cn=Users,dc=rallencorp,dc=com
	' ------ END CONFIGURATION --------

	set objUser = GetObject("LDAP://" & strUserDN)
	Wscript.Echo "Group membership for " & objUser.Get("cn") & ":"
	strSpaces = ""
	set dicSeenGroup = CreateObject("Scripting.Dictionary")
	DisplayGroups("LDAP://" &  
strUserDN, strSpaces, dicSeenGroup)

	Function DisplayGroups ( strObjectADsPath, strSpaces, dicSeenGroup)

	   set objObject = GetObject(strObjectADsPath)
	   WScript.Echo strSpaces & objObject.Name
	   on error resume next ' Doing this to avoid an error when memberOf is empty
	   if IsArray( objObject.Get("memberOf") ) then
	      colGroups = objObject.Get("memberOf")
	   else
	      colGroups = Array( objObject.Get("memberOf") )
	   end if

	   for each strGroupDN In colGroups
	      if Not dicSeenGroup.Exists(strGroupDN) then
	         dicSeenGroup.Add strGroupDN, 1
	         DisplayGroups "LDAP://" & strGroupDN, strSpaces & " ", dicSeenGroup
	      end if
	   next

	End Function
