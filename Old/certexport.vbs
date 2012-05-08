' ------------------------------------------------------------- 
' Export IS Certificates 
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
' Exports all the available certificates, used in IIS. 
' ------------------------------------------------------------ 
' © Copyright 2005 Warren Ashcroft (Helm Plus) 
' All Rights Reserved 
' This script must not be distributed without permission 
' ------------------------------------------------------------ 
' Usage: Save this script as a VBS file and run like this: 
' cscript.exe "c:\path\to\file.vbs" 
' ------------------------------------------------------------ 
' For support email support@helmplus.com 
' Results of this are not my responsibility 
'------------------------------------------------------------- 

'------------------------------------------------------------- 
' Configuration variables 
'------------------------------------------------------------- 
' Export Folder - The folder must exist 
Const ExportFolder = "C:\CertificateExports" 

' Export Password - Required on re-import 
Const ExportPassword = "personix" 

' IIS Server - The IIS Server to exports certificates of 
Const IISServer = "PXHOUSWEB01" 

'------------------------------------------------------------- 
' DO NOT edit below here 
'------------------------------------------------------------- 
Dim oIIS, oWeb, oCert 
Dim iCount 

Set oIIS = GetObject("IIS://" & IISServer & "/W3SVC") 

For Each oWeb in oIIS 
If LCase(oWeb.Class) = "iiswebserver" Then
On Error Goto 0
Set oCert = CreateObject("IIS.CertObj") 

oCert.InstanceName = "W3SVC/" & oWeb.Name 

'On Error Resume Next 
oCert.Export ExportFolder & "\" & oWeb.ServerComment & ".pfx", ExportPassword, True, True, False 

If Err.Number = 0 Then 
iCount = iCount + 1 
WriteLogLine "Exporting SSL Certificate", oWeb.ServerComment 
End If 

Err.Clear 
On Error Goto 0 

Set oCert = nothing 
End If 
Next 

Set oIIS = Nothing 
WriteLogLine "Total Certificates Exported", iCount 

Sub WriteLogLine(ByVal LogLineLabel, ByVal LogLineContent) 
If (Len(LogLineLabel) = 0) And (Len(LogLineContent) = 0) Then 
WScript.Echo "" 
Else 
LogLineLabel = Trim(Left(LogLineLabel, 30)) 
LogLineLabel = LogLineLabel & Replace(Space(30 - Len(LogLineLabel)), " ", ".") 
WScript.Echo "> " & LogLineLabel & ": " & LogLineContent 
End If 
End Sub
