strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colQuickFixes = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering")

For Each objQuickFix in colQuickFixes
    Wscript.Echo "Computer: " & objQuickFix.CSName
    Wscript.Echo "Description: " & objQuickFix.Description
    Wscript.Echo "Hot Fix ID: " & objQuickFix.HotFixID
    Wscript.Echo "Installation Date: " & objQuickFix.InstallDate
    Wscript.Echo "Installed By: " & objQuickFix.InstalledBy
Next
wscript.echo colQuickFixes.count