If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    temp = MsgBox("Script can be processed with CScript.exe only." & vbCrLf & "Examples:" & vbCrLf & "cscript //nologo moveToTrackItOU.vbs" & vbCrLf & "cscript //nologo moveToTrackItOU.vbs /undo", 48, "Trackit OU Mover")
    WScript.Quit
End If

If WScript.Arguments.Named.Exists("undo") Then
WScript.Echo "Undoing!"
undo
WScript.Quit
End If

'strComputer = "hou2020if002"
If Not strComputer = "" Then
GetOUs strComputer
WScript.Quit
End If

Set fso = CreateObject("Scripting.FileSystemObject")
 InputFile = "./trackitdeploy.txt"
 If fso.FileExists(InputFile) Then
  Set txtStreamIn = fso.OpenTextFile(InputFile)
  Do While Not (txtStreamIn.AtEndOfStream)
    server = txtStreamIn.ReadLine
    GetOUs server
  loop
 Else
  WScript.Echo "Input file doesnt exist. Please make sure that the " & Right(inputfile,Len(inputfile)-2) & " file exists in the directory you are running this from."
 End If

Sub GetOUs(strComputer)
strUser = strComputer & "$"
set objRoot = getobject("LDAP://RootDSE")
defaultNC = objRoot.get("defaultnamingcontext")
computerDN = FindUser(strUser, defaultNC)
If computerDN = "" Then WScript.Echo "Computer not found: """ & strComputer & """":Exit Sub
ouarray = Split(computerDN,",")
For i = 1 To UBound(ouarray)
    If ou = "" Then
    ou = ouarray(i)
    Else
    ou = ou & "," & ouarray(i)
    End If
Next

WScript.Echo "Current OU for " & strComputer & " is: " & ou
SetDescription ComputerDN,OU
moveOU ComputerDN

End Sub

Sub SetDescription(ComputerDN,OU)
On Error Resume Next
Set objComputer = GetObject(ComputerDN)
strComputer = Right(objComputer.name,Len(objComputer.name)-3)

strDescription = objComputer.description
'Wscript.Echo "Description: " & strDescription
place = InStr(strDescription,"OriginalOU = ")
If InStr(strDescription,"OriginalOU = ") = 0 Then
objComputer.description = strDescription & ": OriginalOU = " & OU
objComputer.SetInfo
WScript.Echo "Description on " & strComputer & " has been set to " & objComputer.description & "."
Else
WScript.Echo "Original OU already set on " & strComputer
End If

End Sub

Sub MoveOU(ComputerDN)

Set objNewOU = GetObject("LDAP://OU=Track-It,OU=Server Devices,OU=Machines,DC=houston,DC=personix,DC=local")
ouarray = Split(computerDN,",")
ComputerCN = ouarray(0)
ComputerCN = Right(ComputerCN,Len(ComputerCN)-7)
strComputer = Right(ComputerCN,Len(ComputerCN)-3)
Set objMoveComputer = objNewOU.MoveHere(ComputerDN,ComputerCN)

WScript.Echo strComputer & " has been moved to the TrackIt OU."
End Sub

Function FindUser(Byval UserName, Byval Domain)
    on error resume next

    set cn = createobject("ADODB.Connection")
    set cmd = createobject("ADODB.Command")
    set rs = createobject("ADODB.Recordset")

    cn.open "Provider=ADsDSOObject;"

    cmd.activeconnection=cn
    cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & Domain & _
            "' WHERE sAMAccountName = '" & UserName & "'"

    set rs = cmd.execute

    if err<>0 then
        FindUser="Error connecting to Active Directory Database:" & err.description
        'wscript.quit
    else
        if not rs.BOF and not rs.EOF then
                 rs.MoveFirst
                 FindUser = rs(0)
        else
            FindUser = ""
        end if
    end if
    cn.close
end Function

Sub undo()
strUndoOU = "LDAP://OU=Track-It,OU=Server Devices,OU=Machines,DC=houston,DC=personix,DC=local"
Set objUndoOU = GetObject(strUndoOU)
objUndoOU.Filter = Array("computer")
For Each objComputer In objUndoOU
strComputer = Right(objComputer.name,Len(objComputer.name)-3)
strDescription = objComputer.description
On Error Resume Next
arryDescription = Split(strDescription,": OriginalOU = ")
strDescription = arryDescription(0)
strOriginalOU = arryDescription(1)
'WScript.Echo Err.Number
If Err.Number = "9" Then
WScript.Echo "Original OU is not set for " & strComputer & " -- cannot move to Original OU."
Err.Clear
Else
On Error Goto 0
Set objNewOU = GetObject("LDAP://" & strOriginalOU)
computerCN = objComputer.name
arryUndoOU = Split(strUndoOU,"//")
computerDN = arryUndoOU(0) & "//" & computerCN & "," & arryUndoOU(1)

If strDescription = "" Then objComputer.description = " " Else objComputer.description = strDescription
objComputer.setInfo
Set objMoveComputer = objNewOU.MoveHere(ComputerDN,ComputerCN)
WScript.Echo strComputer & " has been moved back to OU: " & strOriginalOU
End If
Next
End Sub