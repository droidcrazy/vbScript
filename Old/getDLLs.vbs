If Right(UCase(WScript.FullName), Len("CScript.exe")) <> UCase("CScript.exe") Then
    'temp = MsgBox("Script can be processed with CScript.exe only.", 48, WScript.ScriptFullName)
    Set objShell = CreateObject("wscript.shell")
    objShell.Run "cmd /K cscript //nologo """ & WScript.ScriptFullName & """",4,False
    WScript.Quit
End If

'option Explicit
On Error Resume Next
If WScript.Arguments.Count = 0 Then
machine = "."
Else
machine = WScript.Arguments.Unnamed(0)
End If


'dim machine, HKCR, hkey, key, Reg, counter
counter = 0
'machine = "houmwpvproc033"
HKCR = &H80000000
hkey = HKCR
key = "TypeLib"
'on error resume next      

Set Reg = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\" _
    & machine & "\root\default:StdRegProv")

' returns an array containing names of subkeys
' under key
dim subkeys, rtn, EnumKey
dim subkeys1, subkeys2 

rtn = Reg.EnumKey(hkey, key, subkeys)
if rtn = 0 then
    EnumKey = subkeys
else
    err.raise vbObjectError + rtn, "RegistryProvider: ", _
        "Error returned attempting to enumerate keys under " _
        & key & ": " & rtn
    wscript.quit 1
end if   

wscript.echo "key count=", ubound(EnumKey)
' HKEY_CLASSES_ROOT\TypeLib\{0A055C02-BABE-4480-BB7B-A8EC723CE9C0}\1.0\0\win32

dim i, j, pos, vallib, valpath
' loop over the GUID values for all registered TypeLibs
for i = 0 to ubound(EnumKey)
   'wscript.echo EnumKey(i)
   
   rtn = Reg.EnumKey(hkey, key & "\" & EnumKey(i), subkeys1)
   'wscript.echo typename(subkeys1), " count=", ubound(subkeys1) 
   
   ' a few typelibs are empty GUID's with no subkey collections, skip them 
   if not typename(subkeys1) = "Null" then
      
      ' loop over the versions
      for j= 0 to ubound(subkeys1)
            
         ' sanity check: version must be a decimal value containing a decimal point! 
         pos = instr(subkeys1(j),".")
         if pos < 1 then exit for
         '  /* end sanity check  */
         
         ' subkeys2(0) is the language code
         rtn = Reg.EnumKey(hkey, key & "\" & EnumKey(i) & "\" & subkeys1(j), subkeys2)
         if ubound(subkeys2)> 2 then wscript.echo "lang_count=", ubound(subkeys2), " ", EnumKey(i)
         ' valib is the library name
         Reg.GetStringValue hkey, _
            key & "\" & EnumKey(i) & "\" & subkeys1(j)  , Null, vallib
         
         Reg.GetStringValue hkey, _
            key & "\" & EnumKey(i) & "\" & subkeys1(j) & "\" & subkeys2(0)& "\" &  "win32" , Null, valpath
         
         if not typename(valpath) = "Null" then
'            WScript.Echo "[" & vallib &  "] ", EnumKey(i)," ",  valpath 
			printFileInfo machine,valpath
            counter = counter + 1
         end if
      next
   else
      WScript.echo "null ", EnumKey(i)
   end if         
next
WScript.Echo counter

Sub printFileInfo(strComputer,fileName)

'strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
fileName = ReplaceTest(fileName,"\\","\\")
Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_Datafile Where name = '" & fileName & "'")

For Each objFile in colFiles
name = objFile.FileName
ext = objFile.Extension
size = objFile.FileSize
version = objFile.Version
WScript.Echo """" & fileName & """," & size & ",""" & version & """"
name = ""
size = ""
version = ""
Next

End Sub

Function ReplaceTest(strInput,patrn,replStr)
  Dim regEx, str1               ' Create variables.
'  str1 = "The quick brown fox jumped over the lazy dog."
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.Global = True
  regEx.IgnoreCase = True            ' Make case insensitive.
  ReplaceTest = regEx.Replace(strInput, replStr)   ' Make replacement.
End Function
