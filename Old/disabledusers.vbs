Set container = GetObject("WinNT://personix")
container.filter = Array("User")
For Each user In container
If user.accountdisabled Then WScript.Echo user.fullname & " : " & getFullName(user.adspath):intCount = intCount +1
Next
WScript.Echo
WScript.Echo "Total of " & intCount & " accounts disabled."


Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName