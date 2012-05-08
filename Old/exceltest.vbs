
'On Error Resume next

Set objShell = CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
objExcel.Visible = True

inputFile = "servers.txt.test"
set objTS = objFS.OpenTextFile(inputFile,1)

Workbook = 1
counter = 1 

Set objWorksheet = objExcel.Workbooks.Add

OuName = objTS.readline
do until objTS.AtEndOfStream



'Once the first 3 premade sheets are done, add new ones
If counter > 3 Then

   set objWorksheet = objExcel.Sheets.Add( , objExcel.WorkSheets(objExcel.WorkSheets.Count))

'Set objWorksheet = objExcel.Workbooks.Add


   'set objWorksheet = objExcel.WorkSheets.Add

 
 End If 

 objExcel.worksheets(counter).Activate
   
 objExcel.worksheets(counter).Name = OuName



counter = counter +1
OuName = objTS.readline
Loop


msgBox "Finished"


