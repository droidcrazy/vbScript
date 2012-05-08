Const HKEY_LOCAL_MACHINE = &H80000002

On Error Resume Next
Dim Controller, RemoteScript
strComputer = "pxhousdc02"
Set Controller = WScript.CreateObject("WSHController")

Set objRegProv = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows Script Host\Settings"
If Err.Number <> 0 Then WScript.Echo "Error 0x" & Hex(Err.Number) & " occured. Description: " & Err.Description & " Source: " & Err.Source : Err.Clear : WScript.Quit

objRegProv.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"Remote","1"
If Err.Number <> 0 Then WScript.Echo "Error 0x" & Hex(Err.Number) & " occured. Description: " & Err.Description & " Source: " & Err.Source : Err.Clear : WScript.Quit

Set RemoteScript = Controller.CreateScript("CertExpiryCheck.vbs", strComputer)
If Err.Number <> 0 Then WScript.Echo "Error 0x" & Hex(Err.Number) & " occured. Description: " & Err.Description & " Source: " & Err.Source : Err.Clear : WScript.Quit
WScript.ConnectObject RemoteScript, "remote_"
RemoteScript.Execute

Do While RemoteScript.Status <> 2 
    WScript.Sleep 100
Loop

WScript.DisconnectObject RemoteScript

Sub remote_Error
    Dim theError
    Set theError = RemoteScript.Error
    WScript.Echo "Error 0x" & Hex(theError.Number) & " - Line: " & theError.Line & ", Char: " & theError.Character & vbCrLf & "Description: " & theError.Description
    WScript.Quit -1
End Sub
