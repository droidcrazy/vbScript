On Error Resume Next
 
strComputer = "."
Set objUser = GetObject("WinNT://houston/sa_tbecker,user")
Set objClass = GetObject(objUser.Schema)
 
WScript.Echo "Mandatory properties for " & objUser.Name & ":"
For Each property In objClass.Properties
    WScript.Echo property, objUser.Get(property)
Next
 
WScript.Echo "Optional properties for " & objUser.Name & ":"
For Each property In objClass.OptionalProperties
    WScript.Echo property, objUser.Get(property)
Next
