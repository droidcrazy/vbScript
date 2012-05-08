'These define the style of the border
Const xlNone = -4142
Const xlContinuous = 1
Const xlDash = -4115
Const xlDashDot = 4
Const xlDashDotDot = 5
Const xlDot = -4118
Const xlDouble = -4119
Const xlSlantDashDot = 13

'These define the weight of the border
Const xlHairLine = 1
Const xlMedium = -4138
Const xlThick = 4
Const xlThin = 2

'Thise is handy to make borders have the default color index
Const xlAutomatic = -4105

'These define the placement of border pieces
Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12

domain = inputbox ("Please enter domain name:","Group Review","houston")
group = inputbox ("Please enter group name:","Group Review","domain admins")

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
'objExcel.Workbooks.Add
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
objExcel.worksheets(2).delete
objExcel.worksheets(2).delete
objExcel.Worksheets(1).name = "Admin Review"

Select Case Month(Now)
	Case 1,2,3
	quarter = "Q1"
	Case 4,5,6
	quarter = "Q2"
	Case 7,8,9
	quarter = "Q3"
	Case 10,11,12
	quarter = "Q4"
	Case Else
	quarter = "Unknown"
End Select

objExcel.Range("a1").Value = quarter & " Review of """ & domain & "\" & group & """ Members"
objExcel.Range("b1").Value = "Today's date:"
objExcel.Range("c1").Value = date
objExcel.Cells(3, 1).value = "Name"
objExcel.Cells(3, 2).value = "Usage Activity"
objExcel.Cells(3, 3).value = "Account Type"
objExcel.Cells(3, 4).value = "Status"
With objExcel.Range("a1","d3").Font
	.Name = "Calibri"
	.Size = 11
	.Bold = True
End With

Set grp = GetObject("WinNT://" & domain & "/" & group)
rowvar = 4
For Each member In grp.members
objExcel.Cells(RowVar, 1).value = member.fullname & " (" & getFullName(member.ADSpath) & ")"
On Error Resume Next
objExcel.Cells(RowVar, 2).value = member.lastlogin
If Err.Number <> 0 Then 
Select Case Err.Number
Case -2147463155
objExcel.Cells(RowVar, 2).value = "No login date."
Case Else
objExcel.Cells(RowVar, 2).value = Err.Number & ":" & Err.Description
End Select
Err.Clear
End If
On Error Goto 0
Select Case member.accountdisabled
Case -1
objExcel.Cells(RowVar, 4).value = "Disabled"
objExcel.range((objexcel.cells(RowVar, 1)), (objexcel.cells(rowvar, 4))).Interior.Color = 49407
Case 0
objExcel.Cells(RowVar, 4).value = "Active"
Case Else
objExcel.Cells(RowVar, 4).value = "Unknown Account Status"
End Select
rowvar = rowvar + 1
Next

With objExcel.Range("A3:D" & rowvar - 1)
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End With


objExcel.Cells.Select
objExcel.Cells.EntireColumn.AutoFit

rowvar = rowvar + 1

objExcel.Cells(RowVar, 1).value = "List Reviewed By:"
objExcel.Cells(RowVar + 1, 1).value = "Scott Schultz - VP of Infrastructure"
objExcel.Cells(RowVar + 3, 1).value = "Jeff Richards - CIO"
objExcel.Range("A"&rowvar&":D"&rowvar+3).Font.Bold = True

With objExcel.Range("b"&rowvar+1&":D"&rowvar+1).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

With objExcel.Range("b"&rowvar+3&":D"&rowvar+3).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

Set WshShell = wscript.createobject("wscript.shell")
strTime = Right(100 + Month(Now), 2) & "-" & Right (100 + Day(Now), 2) & "-" & Year(Now) & "." & Right(100 + hour(now), 2) & Right( 100 + Minute(now), 2)
OutputFile = wshshell.currentdirectory & domain & "-" & group & " " & quarter & " review." & strTime & ".xlsx"
objWorkbook.SaveAs OutputFile
objExcel.quit
wscript.echo "Administrator group enumeration is done. Output file is " & OutputFile

Function getFullName(username) 
   arrayU = Split(username,"/") 
   arraylen = UBound(arrayU) 
   getFullName = arrayU(arraylen - 1) & "/" & arrayU(arraylen) 
End Function 'getFullName