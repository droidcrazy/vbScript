Set container = GetObject("WinNT://personix")
container.filter = Array("User")
For Each user In container
WScript.Echo getFullName(user.adspath):intCount = intCount +1
Next
WScript.Echo


Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName