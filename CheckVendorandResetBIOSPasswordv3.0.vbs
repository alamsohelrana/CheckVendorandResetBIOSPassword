
Function ClearLenovoBiosPassword(currPassword)
	Dim colItems
	strRequest = "POP," + currPassword + "," + "," + "ascii,us" + ";"

	strComputer = "LOCALHOST" ' Change as needed.

	Set objWMIService = GetObject("WinMgmts:" _
	&"{ImpersonationLevel=Impersonate}!\\" & strComputer & "\root\wmi")

	Set colItems = objWMIService.ExecQuery("Select * from Lenovo_SetBiosPassword")

	strReturn = "error"

	For Each objItem in colItems
		ObjItem.SetBiosPassword strRequest, strReturn
	Next

	WScript.Echo " ClearLenovoBiosPassword: "+ strReturn
End Function

Function ClearHPBiosPassword(currPassword)
	strComputer = "."

    Set objSWbemLocator = CreateObject _
      ("WbemScripting.SWbemLocator")
    Set objWMIService = objSWbemLocator.ConnectServer _
      (strComputer, "root\HP\InstrumentedBIOS")
    Set objShare = objWMIService.Get _
      ("HPBIOS_BIOSPassword.InstanceName='ACPI\PNP0C14\1_0'")

	IF objShare.IsSet = 1 Then
		strName     = "Setup Password"
		strPassword = "<utf-16/>" & currPassword ' This must be the OLD Password that is set
		strAttributeValue = "<utf-16/>"

		Set objShare = objWMIService.Get _
		  ("HPBIOS_BIOSSettingInterface.InstanceName='ACPI\PNP0C14\1_0'")
	  
		'Wscript.Echo objShare.InstanceName + " ======= " '+ objShare.vALUE

		Set objInParam = _
		objShare.Methods_("SetBIOSSetting").InParameters.SpawnInstance_()

		objInParam.Properties_.Item("Name")     = strName
		objInParam.Properties_.Item("Password") = strPassword
		objInParam.Properties_.Item("Value")    = strAttributeValue

		Set objOutParams = objWMIService.ExecMethod _
		("HPBIOS_BIOSSettingInterface.InstanceName='ACPI\PNP0C14\1_0'", _
		"SetBIOSSetting", objInParam)

		strResultMsg = "Result: " & objOutParams.Return
		Dim strReturn
		Select Case objOutParams.Return
			Case 0 strReturn = "Success"
			Case 1 strReturn = "Not Supported"
			Case 2 strReturn = "Unspecified Error"
			Case 3 strReturn = "Timeout"
			Case 4 strReturn = "Failed"
			Case 5 strReturn = "Invalid Parameter"
			Case 6 strReturn = "Access Denied"
			Case Else strReturn = "..."
		End Select
		WScript.Echo
		WScript.Echo "ClearHPBiosSetupPassword returned: " & strReturn 
		Else
		WScript.Echo "ClearHPBiosSetupPassword No Password is Set"
	End If
End Function

Function ClearDELLBiosPassword(currPassword)
    Wscript.Echo "It's a Dell System"
	Dim strNameSpace 
	Dim strComputerName 
	Dim strClassName 
	Dim objInstance 
	Dim strAttributeName
	Dim strAttributeValue
	Dim strAuthorizationToken
	Dim returnValue 
	Dim objWMIService
	Dim ColSystem
	Dim oInParams

	strNameSpace = "root/dcim/sysman" 
	strComputerName = "."
	strClassName = "DCIM_BIOSService" 
	strAttributeName = "AdminPwd"
	strAttributeValue = "" 'New Password
	strAuthorizationToken = currPassword
	returnValue = 0 
	'*** Retrieve the instance of DCIM_BIOSService class

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate," &_ 
	"AuthenticationLevel=pktprivacy}\\" & strComputerName & "\" &_ 
	strNameSpace) 
	Set ColSystem=objWMIService.execquery ("Select * from " & strClassName) 
	For each objInstance in ColSystem  
		Set oInParams= objInstance.Methods_("SetBIOSAttributes").InParameters.SpawnInstance_ 
		oInParams.AttributeName = strAttributeName
		oInParams.AttributeValue = strAttributeValue
		oInParams.AuthorizationToken = strAuthorizationToken 
		Set returnValue = objInstance.ExecMethod_("SetBIOSAttributes", oInParams) 
	Next  
	'*** If any errors occurred, let the user know

	If Err.Number <> 0 Then
		WScript.Echo "Clearing admin password failed." 
	Else
        WScript.Echo "Clearing admin password Succeded."
	End If
End Function

strClassName = "Win32_ComputerSystemProduct" 

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate," &_ 
"AuthenticationLevel=pktprivacy}\\" & "." & "\" &_ 
strNameSpace) 
Set ColSystem=objWMIService.execquery ("Select Vendor,Name,Version from " & strClassName) 

For each objInstance in ColSystem
	SystemVendor = objInstance.Vendor
Next

'Wscript.Echo CStr(SystemVendor)
'If SystemVendor match

Set reDELL = New RegExp
With reDELL
    .Pattern    = "\bdell\b"
    .IgnoreCase = True
    .Global     = False
End With

Set reLENOVO = New RegExp
With reLENOVO
    .Pattern    = "\bLENOVO\b"
    .IgnoreCase = True
    .Global     = False
End With

Set reHP = New RegExp
With reHP
    .Pattern    = "\bHP\b"
    .IgnoreCase = True
    .Global     = False
End With

If reDELL.Test( CStr(SystemVendor) ) Then
    Wscript.Echo "It's a Dell System"
    On Error Resume Next
    ClearDELLBiosPassword("1234")
    ClearDELLBiosPassword("12sfsd4")
    ClearDELLBiosPassword("1GFgfs34df234")
ElseIf reLENOVO.Test( CStr(SystemVendor) ) Then
    Wscript.Echo "It's a Lenovo System"
	On Error Resume Next
	ClearLenovoBiosPassword("123456")
	ClearLenovoBiosPassword("123456")
	ClearLenovoBiosPassword("123456")
	ClearLenovoBiosPassword("123456")
	ClearLenovoBiosPassword("123456")
ElseIf reHP.Test( CStr(SystemVendor) ) Then
    Wscript.Echo "It's an HP System"
	On Error Resume Next
	ClearHPBiosPassword("hp@1358")
	ClearHPBiosPassword("Sohi2016")
	ClearHPBiosPassword("ath3ns2004")
	ClearHPBiosPassword("LImuId_123")
	ClearHPBiosPassword("dsfver345")
	Else
    Wscript.Echo "System vendor name : " & SystemVendor & " ; which is unknown !, exiting the script"
    WScript.Quit 
End If

