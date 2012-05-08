Dim WshShell, BtnCode
Set WshShell = WScript.CreateObject("WScript.Shell")

BtnCode = WshShell.Popup("Do you feel alright?", 7, "Answer This Question:", 4 + 32)

Select Case BtnCode
   case 6      WScript.Echo "Glad to hear you feel alright."
   case 7      WScript.Echo "Hope you're feeling better soon."
   case -1     WScript.Echo "Is there anybody out there?"
End Select
