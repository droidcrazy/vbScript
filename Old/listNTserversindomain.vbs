Set objAdRootDSE = GetObject("LDAP://RootDSE")
Set objRS = CreateObject("adodb.recordset")
 
  varConfigNC = objAdRootDSE.Get("defaultNamingContext")
  strConnstring = "Provider=ADsDSOObject"
  strWQL = "SELECT * FROM 'LDAP://" & varConfigNC & "' WHERE objectCategory= 'Computer'"
  objRS.Open strWQL, strConnstring
    Do until objRS.eof
       Set objServer = GetObject(objRS.Fields.Item(0))
      strServerName = objServer.CN
      strOperatingSystem = objServer.OperatingSystem
      If InStr(strOperatingSystem, "NT") Then WScript.Echo  strServerName' & " is running " & strOperatingSystem
       objRS.movenext
       Set objServer = Nothing
    Loop
  objRS.close
 
Set objRS = Nothing
Set objAdRootDSE = Nothing