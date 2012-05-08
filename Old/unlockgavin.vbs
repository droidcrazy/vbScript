Set objUser = GetObject("LDAP://cn=(SA) Gavin Fuller,ou=networking,ou=departments,dc=houston,dc=personix,dc=local")
If objuser.isaccountlocked Then
objUser.IsAccountLocked = False
objUser.SetInfo
WScript.Echo objuser.fullname & " is now unlocked."
Else
WScript.Echo objuser.fullname & " is not locked at this time."
End If