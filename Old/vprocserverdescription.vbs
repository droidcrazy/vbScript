Set objSysInfo = CreateObject("ADSystemInfo")

Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
Set objComputer = GetObject("LDAP://" & objSysInfo.ComputerName)

strMessage = "Last deployed on " & Now & "."

objComputer.Description = strMessage
objComputer.SetInfo