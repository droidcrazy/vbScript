Option Explicit 
On error resume next 
Const ADS_SCOPE_SUBTREE = 2 
'******************************************************************************* 
'*                                                                                                                                                          * 
'*    This Script list all the Users, their Group membership and the                                                      * 
'*   last time they logged on the computer, in a ~ delimited text file                                                  * 
'*         for easy import into database or Excel spreadsheet.                                                              * 
'*                                                                                                                                                          * 
'******************************************************************************* 
' Copyright (c) 2006 Vicky Desjardins 
' Version 1.0 - November 28, 2006 
' 
' You have a royalty-free right to use, modify, reproduce, and 
' distribute this script file in any way you find useful, provided that 
' you agree that the copyright owner above has no warranty, obligations, 
' or liability for such use. 

Dim colCSItems 
Dim dtmEventDate 
Dim Groups 
Dim Message 
Dim objFSO 
Dim objStream 
Dim objConnection 
Dim objCommand 
Dim objRecordSet 
Dim objComputer 
Dim objCSItem 
Dim objWMIService 
Dim PrimaryDate 
Dim strComputer 
Dim strDateLastlog 
Dim strGenerated 
Dim strLine 
Dim Utilisateur 
Dim UserID 

Set objFSO = createobject("scripting.filesystemobject") 
Set objStream = objFSO.CreateTextFile("C:\RetrieveUserSecurity.txt", True) 
Set objConnection = CreateObject("ADODB.Connection") 
Set objCommand =   CreateObject("ADODB.Command") 
Set objComputer = CreateObject("Shell.LocalMachine") 

Message = "Generating Text File C:\RetrieveUserSecurity.txt." & VbCrLF 
Message = Message & "Depending on your system this could take a while." & VbCrLF 
Message = Message & "                     Please Wait..." 
Msgbox (Message) 
' ****************************************************************************** 
' *                                     Creating Titles                                                                                            * 
' *   UserID     Full Name     Description     Group Membership     Last logon                                       * 
' *                                                                                                                                                        * 
' ******************************************************************************                                                                             
StrLine = "UserID~Full Name~Description~Group Membership~Last Logon" 
objStream.Writeline strLine 
strLine= "" 
' ****************************************************************************** 
' *                                                                                                                                                        * 
' *                 Main Module Query computer for users                                                                         * 
' *                                                                                                                                                        * 
' ****************************************************************************** 
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2") 
Set colCSItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount") 
For Each objCSItem In colCSItems 
 Utilisateur = objCSItem.Name 
 strLine = strLine & objCSItem.Name & "~" 
 strLine = strLine & objCSItem.FullName & "~" 
 strLine = strLine & objCSItem.Description & "~" 
   Groups = GroupMembership (Utilisateur) 
   strLine = strLine & Groups & "~" 
   
 UserID = objComputer.MachineName & "\\" & Utilisateur 
 strDateLastlog = Lastlog(UserID) 
   PrimaryDate = 19700830074757 
 if strDateLastlog = PrimaryDate then 
  strLine = strLine & "Never Logged on " 
 Else 
  dtmEventDate = strDateLastlog 
  strGenerated = WMIDateStringToDate(dtmEventDate) 
  strLine = strLine & "Last loggin : " & strGenerated 
 End if 

 strDateLastlog = "" 
 Utilisateur ="" 
 objStream.Writeline strLine 
 strLine= "" 
Next 

'******************************************************************************* 
'*                                                                                                                                                         * 
'*                           Convert Date Function                                                                                          * 
'*                                                                                                                                                         * 
'******************************************************************************* 
Function WMIDateStringToDate(dtmEventDate) 
   WMIDateStringToDate = CDate(Mid(dtmEventDate, 5, 2) & "/" & _ 
       Mid(dtmEventDate, 7, 2) & "/" & Left(dtmEventDate, 4) _ 
           & " " & Mid (dtmEventDate, 9, 2) & ":" & _ 
               Mid(dtmEventDate, 11, 2) & ":" & Mid(dtmEventDate, _ 
                   13, 2)) 
End Function 

'******************************************************************************* 
'*                                                                                                                                                         * 
'*                     Retreive User Last logon Function                                                                              * 
'*                                                                                                                                                         * 
'******************************************************************************* 
Function Lastlog (StrUserID) 
Dim strComputer 
Dim TestDate 
Dim StrQuery 
Dim objWMIService 
Dim collectionItems 
Dim objItem 
Dim PreDate 

strComputer = "." 
TestDate = 19700830074757 
StrQuery = "select * from win32_ntlogEvent Where LogFile='Security' " 
StrQuery = StrQuery &  "and eventCode = 528 and User =  '" & StrUserID  & "'" 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,(Security)}!\\" & strComputer & "\root\CIMV2") 
Set collectionItems = objWMIService.ExecQuery(StrQuery) 
    For Each objItem In collectionItems 
    PreDate = objItem.TimeGenerated 
    If TestDate <= PreDate then 
    TestDate = objItem.TimeGenerated 
    End if 
    Next 
Lastlog = TestDate 
End Function 

'******************************************************************************* 
'*                                                                                                                                                         * 
'*                       Retreive User's Group membership                                                                           * 
'*                                                                                                                                                         * 
'******************************************************************************* 
Function GroupMembership (UserName) 
Dim StrGroups 
Dim strComputer 
Dim Virgule 
Dim colGroups 
Dim objGroup 
Dim objUser 

StrGroups = "" 
strComputer = "." 
Virgule = 0 
Set colGroups = GetObject("WinNT://" & strComputer & "") 
colGroups.Filter = Array("group") 
For Each objGroup In colGroups 
   For Each objUser in objGroup.Members 
       If objUser.name = Username Then 
          if Virgule <> 0 then StrGroups = StrGroups & "," End if 
          StrGroups = StrGroups & objGroup.Name 
          Virgule = Virgule + 1 
       End If 
   Next 
Next 
GroupMembership = StrGroups 
End Function 
