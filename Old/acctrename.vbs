If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K title " & WScript.ScriptFullName & " Output:&cscript //nologo " & WScript.ScriptFullName,4,False
    WScript.Quit
End If

Do While strOldName = ""
strOldName = InputBox("Please Enter the old account name:","User Rename Script","administrator")
If strOldName = "" Then If MsgBox("You need to type in a name.",vbOKCancel,"Username Error.") = vbCancel Then WScript.Quit
Loop

Do While strNewName = ""
strNewName = inputbox("What would you like it to be?","User Rename Script")
If strNewName = "" Then If MsgBox("You need to type in a name.",vbOKCancel,"Username Error.") = vbCancel Then WScript.Quit
Loop


Set WshShell = WScript.CreateObject("WScript.Shell")
buttoncode = WshShell.Popup("Do you want to reverse?" & vbCrLf & "This is to undo." & vbCrLf & "You have 10 seconds to choose.", 10, "Undo.", 4 + 32)
Select Case buttoncode
   Case 6      undo = True
   Case 7      undo = False
   Case -1     undo = False
   Case Else   undo = False
End Select

Set fso = CreateObject("Scripting.FileSystemObject")

Set fso = CreateObject("Scripting.FileSystemObject")
answer = MsgBox("Click OK to choose an input file"  & Chr(10) & Chr(13) & "or click Cancel to use the default of rename.txt", 65, "User Rename Script")
If answer = 1 Then
Set ObjFSO = CreateObject("UserAccounts.CommonDialog")
ObjFSO.Filter = "Text Documents|*.txt"
'ObjFSO.Title = "Select an Input File"
ObjFSO.FilterIndex = 3
ObjFSO.InitialDir = wshshell.currentdirectory
InitFSO = ObjFSO.ShowOpen
If InitFSO = False Then
    Wscript.Echo "Script Error: Please select a file!"
    Wscript.Quit
Else
    inputfile = ObjFSO.FileName
End If
Else
InputFile = "./rename.txt"
End If


If fso.FileExists(InputFile) Then
	Set txtStreamIn = fso.OpenTextFile(InputFile)
	Do While Not ( txtStreamIn.AtEndOfStream )
		server = txtStreamIn.ReadLine
		If Not undo Then rename server, strOldName, strNewName
		if undo then rename server,strNewName,strOldName
	Loop
Else
	WScript.Echo "Input file doesnt exist. Please make sure that the " & Right(InputFile,Len(InputFile) - 2) & " file exists in the directory you are running this from."
End If

Sub rename(strComputer, strOldName, strNewName)
	On Error Resume Next
	Set objComp = GetObject("WinNT://" & strComputer)
	Set objUser = GetObject("WinNT://" & strComputer & "/" & strOldName & ",user")
	If Err.Number = "-2147022675" Then WScript.echo "User """ & strOldName & """ not found on " & strComputer & "." : Err.Clear : Exit Sub
	Set objNewUser = objComp.MoveHere(objUser.ADsPath, strNewName)
	If Err.Number <> 0 Then WScript.echo "Error:" & Err.Number & ". Description:" & Err.Description & ". On " & strComputer & "." : Err.Clear : Exit Sub
	WScript.Echo "Successfully renamed account on " & strComputer & " to: " & strNewName
	On Error Goto 0
	Set ArgObj = Nothing
End Sub