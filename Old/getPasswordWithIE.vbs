Call getPassword()



Function getPassword()
Set objExplorer = WScript.CreateObject("InternetExplorer.Application","IE_")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width=400
objExplorer.Height = 250 
objExplorer.Left = 0
objExplorer.Top = 0

Do While (objExplorer.Busy)
    Wscript.Sleep 200
Loop

strBody = "<HTML><SCRIPT LANGUAGE=""VBScript"">" & vbCrLf & _
"Sub RunScript" & vbCrLf & _
"    OKClicked.Value = ""OK""" & VbCrLf & _
"End Sub" & VbCrLf & _
"Sub CancelScript" & vbCrLf & _
"    OKClicked.Value = ""Cancelled""" & vbCrLf & _
"End Sub" & vbCrLf & _
"</SCRIPT>" & vbCrLf & _
"<BODY>" & vbCrLf & _
"<font size=""2"" face=""Arial"">" & vbCrLf & _
"Password:&nbsp;&nbsp;&nbsp; </font><font face=""Arial"">" & vbCrLf & _
"<input type=""password"" name=""UserPassword"" size=""40""></font></p>" & vbCrLf & _
"<input type=""hidden"" name=""OKClicked"" size = ""20"">" & vbCrLf & _
"<input id=runbutton class=""button"" type=""button"" value="" OK """ & vbCrLf & _ 
"name=""ok_button"" onClick=""RunScript"">" & vbCrLf & _
"&nbsp;&nbsp;&nbsp;" & vbCrLf & _
"<input id=runbutton class=""button"" type=""button"" value=""Cancel""" & vbCrLf & _
"name=""cancel_button"" onClick=""CancelScript"">" & vbCrLf & _
"</BODY></HTML>"

Set WshShell = CreateObject("wscript.shell")
Set fso = CreateObject("scripting.filesystemobject")
temp = wshshell.ExpandEnvironmentStrings("%temp%")
Set html = fso.OpenTextFile(temp & "my.html", 2, True)
html.Write(strbody)
html.Close

On Error Resume Next
objExplorer.Visible = 1             
objExplorer.Navigate "file:///" & temp & "my.html"
If Err.Number <> 0 Then WScript.StdErr.WriteLine "Error 0x" & Hex(Err.Number) & ": " & Err.Description & ": " & Err.Source:objExplorer.Quit:Exit Function
On Error Goto 0

Do While (objExplorer.Busy)
    Wscript.Sleep 200
Loop    
Wscript.Sleep 200
Do While (objExplorer.Document.Body.All.OKClicked.Value = "")
    Wscript.Sleep 250                 
Loop 
fso.DeleteFile temp & "my.html"
getPassword = objExplorer.Document.Body.All.UserPassword.Value
objExplorer.Quit
End Function