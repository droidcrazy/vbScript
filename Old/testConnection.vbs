Const adOpenStatic = 3
Const adLockOptimistic = 3
Const ForWriting = 2
Const ForReading = 1
Const ForAppending = 8

sSQL = "Select PUBLIC_VIEWS.vUpdate.UpdateId, PUBLIC_VIEWS.vUpdate.CreationDate, " & _
	"  PUBLIC_VIEWS.vUpdate.InstallationRebootBehavior, " & _
	"  PUBLIC_VIEWS.vUpdate.IsDeclined, PUBLIC_VIEWS.vUpdate.DefaultTitle " & _
	"From PUBLIC_VIEWS.vUpdate " & _
	"Where PUBLIC_VIEWS.vUpdate.IsDeclined = 'false' " & _
	"Order By SUSDB.PUBLIC_VIEWS.vUpdate.CreationDate"
	
	
set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")

objConnection.Open "Provider=SQLOLEDB;Data Source=houmwinetop01;Trusted_Connection=Yes;Initial Catalog=SUSDB;"
objRecordSet.Open sSQL, objConnection, adOpenStatic, adLockOptimistic
objRecordset.MoveFirst
While Not objRecordset.EOF
	strOutput = """"
	strOutput = strOutput&objRecordset.Fields("UpdateId").value&""","""
	strOutput = strOutput&objRecordset.Fields("DefaultTitle").value&""","""
	strOutput = strOutput&objRecordset.Fields("CreationDate").value&""","""
	strOutput = strOutput&objRecordset.Fields("InstallationRebootBehavior").value&""""
	WScript.Echo strOutput
	objRecordset.MoveNext
Wend
objRecordset.Close
objConnection.Close