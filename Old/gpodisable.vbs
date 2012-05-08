	' This code can enable or disable the user or computer settings of a GPO.
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set Arg = Wscript.Arguments
	
	' ------ SCRIPT CONFIGURATION ------
	strGPO    = arg.item(0)
	strDomain = "Houston.personix.local" ' e.g. "rallencorp.com"
	boolUserEnable = False
	boolCompEnable = False
	' ------ END CONFIGURATION --------
	'If strGPC = "" Then
	'	wscript.echo "No GPO Specified"
	'Else
	
	set objGPM = CreateObject("GPMgmt.GPM")
	set objGPMConstants = objGPM.GetConstants( )

	' Initialize the Domain object
	set objGPMDomain = objGPM.GetDomain(strDomain, "", objGPMConstants.UseAnyDC)

	' Find the specified GPO
	set objGPMSearchCriteria = objGPM.CreateSearchCriteria
	objGPMSearchCriteria.Add objGPMConstants.SearchPropertyGPODisplayName, _
	                         objGPMConstants.SearchOpEquals, cstr(strGPO)
	set objGPOList = objGPMDomain.SearchGPOs(objGPMSearchCriteria)
	if objGPOList.Count = 0 then
	   WScript.Echo "Did not find GPO: " & strGPO
	   WScript.Echo "Exiting."
	   WScript.Quit
	elseif objGPOList.Count > 1 then
	   WScript.Echo "Found more than one matching GPO. Count: " & _
	                objGPOList.Count
	   WScript.Echo "Exiting."
	   WScript.Quit
	else
	   WScript.Echo "Found GPO: " & objGPOList.Item(1).DisplayName
	end if

	' You can comment out either of these if you don't want to set one:

	objGPOList.Item(1).SetUserEnabled boolUserEnable
	WScript.Echo "User settings: " & boolUserEnable

	objGPOList.Item(1).SetComputerEnabled boolCompEnable
	WScript.Echo "Computer settings: " & boolCompEnable
'	End If
