Option Explicit
Dim iiscertobj, targetServer, targetServers, pfxbasename, pfxpassword, InstanceName, info
Set iiscertobj = WScript.CreateObject("IIS.CertObj")
pfxbasename = WScript.Arguments(0)
pfxpassword = WScript.Arguments(1)
InstanceName = WScript.Arguments(2)
targetServers = split(WScript.Arguments(3), ",")
iiscertobj.UserName = WScript.Arguments(4)
'iiscertobj.UserPassword = WScript.Arguments(5)
'iiscertobj.InstanceName = InstanceName
For Each targetServer in targetServers
'	info = iiscertobj.getcertinfo
  iiscertobj.ServerName = targetServer
'  WScript.Echo iiscertobj.isexportable
  iiscertobj.Export pfxbasename + targetServer + ".pfx", pfxpassword, true, false, false
Next
