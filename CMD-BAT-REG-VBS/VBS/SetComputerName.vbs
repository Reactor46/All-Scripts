<job id="SetComputerName">
<script language="VBScript" src="ZTIUtility.vbs"/>
<script language="VBScript">

' // ***************************************************************************
' // 
' // SetComputerName.wsf
' //
' // File:      SetComputerName.wsf
' // Version:   1.4
' // Date:	04/04/2012
' // Owner:	David Coulter
' // 
' // Purpose:	Determine and Set the OSDComputerName property
' // 
' // Deps:	ZTIUtility.vbs
' //
' // Change:    1.4 Modified Asset check to Trim (we found some asset tags with spaces, not blanks)
' //            1.3 Changed IsValidName check from 'VM' to 'V' per new VDI naming standard
' // 
' // ***************************************************************************

Option Explicit

Dim iRetVal

On Error Resume Next
iRetVal = ZTIProcess
ProcessResults iRetVal
On Error Goto 0

Function ZTIProcess()
	oLogging.CreateEntry "SetComputerName: Begin ZTIProcess", LogTypeInfo

	Dim strMfg, strModel, strCSName
	Dim strAssetTag
	Dim strSerialNumber
	Dim objWMIService
	Dim colCompSystem, colSysEnclosure, colBIOS
	Dim objItem
	Dim blnValidCSName, blnValidAssetTag, blnValidSerialNumber
	Dim strNewDeviceName : strNewDeviceName = Null

	oLogging.CreateEntry "SetComputerName: Start - Collect Info", LogTypeInfo
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colCompSystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each objItem in colCompSystem
		strMfg = objItem.Manufacturer
		oLogging.CreateEntry "SetComputerName: Mfg = " & strMfg, LogTypeInfo
		strModel = objItem.Model
		oLogging.CreateEntry "SetComputerName: Model = " & strModel, LogTypeInfo
		strCSName = objItem.Name
		oLogging.CreateEntry "SetComputerName: Name = " & strCSName, LogTypeInfo
	Next

	Set colSysEnclosure = objWMIService.ExecQuery("SELECT * FROM Win32_SystemEnclosure")
	For Each objItem in colSysEnclosure
		strAssetTag = objItem.SMBIOSAssetTag
		oLogging.CreateEntry "SetComputerName: AssetTag = " & strAssetTag, LogTypeInfo
	Next

	Set colBIOS = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
	For Each objItem in colBIOS
		strSerialNumber = objItem.SerialNumber
		oLogging.CreateEntry "SetComputerName: SerialNumber = " & strSerialNumber, LogTypeInfo
	Next
	oLogging.CreateEntry "SetComputerName: Finish - Collect Info", LogTypeInfo
	
	oLogging.CreateEntry "SetComputerName: Start - Naming Logic", LogTypeInfo
	If IsNull(strCSName) = True Then
		oLogging.CreateEntry "SetComputerName: Name is Null", LogTypeWarning
	ElseIf StrComp(Left(strCSName, 6), "MININT", vbTextCompare) = 0 OR StrComp(Left(strCSName, 8), "MINWINPC", vbTextCompare) = 0 Then
		oLogging.CreateEntry "SetComputerName: Name is Defaulted", LogTypeWarning
	ElseIf NOT (StrComp(Left(strCSName, 1), "D", vbTextCompare) = 0 AND Len(strCSName) = 7) AND NOT (strComp(Left(strCSName, 1), "V", vbTextCompare) = 0) Then
		oLogging.CreateEntry "SetComputerName: Name is not Valid", LogTypeWarning
	Else
		blnValidCSName = True
		oLogging.CreateEntry "SetComputerName: Name is Valid", LogTypeInfo
	End If	
	If IsNull(strAssetTag) = True OR Len(Trim(strAssetTag)) = 0 Then
		oLogging.CreateEntry "SetComputerName: AssetTag is Null", LogTypeWarning
	ElseIf (Trim(strAssetTag) = "No Asset Tag" OR LCase(Trim(strAssetTag)) = "to be filled by o.e.m.") Then
		oLogging.CreateEntry "SetComputerName: AssetTag is Incorrectly Set", LogTypeWarning
	Else
		blnValidAssetTag = True
		oLogging.CreateEntry "SetComputerName: AssetTag is Valid", LogTypeInfo
	End If
	If IsNull(strSerialNumber) = True Then
		oLogging.CreateEntry "SetComputerName: SerialNumber is Null", LogTypeWarning
	ElseIf LCase(Trim(strSerialNumber)) = "to be filled by o.e.m." Then
		oLogging.CreateEntry "SetComputerName: SerialNumber is Incorrectly Set", LogTypeWarning
	Else
		blnValidSerialNumber = True
		oLogging.CreateEntry "SetComputerName: SerialNumber is Valid", LogTypeInfo
	End If
	
	If (strModel = "Virtual Machine" OR strModel = "VMware Virtual Platform" OR strModel = "VirtualBox") Then
		oLogging.CreateEntry "SetComputerName: Device is a Virtual Machine... using Name if Valid", LogTypeInfo
		If blnValidCSName Then
			strNewDeviceName = strCSName
		End If
	Else
		If blnValidCSName Then
			strNewDeviceName = strCSName
		ElseIf blnValidAssetTag Then
			strNewDeviceName = strAssetTag
		End If
	End If
	
	oLogging.CreateEntry "SetComputerName: Finish - Naming Logic", LogTypeInfo

	If IsNull(strNewDeviceName) Then
		oLogging.CreateEntry "SetComputerName: strNewDeviceName is blank... prompting", LogTypeWarning
		strNewDeviceName = InputBox("Please enter a name for this Device.", "Computer Name", , 30,30)
	End If

	oLogging.CreateEntry "SetComputerName: OSDComputerName Set to " & strNewDeviceName, LogTypeInfo
	oEnvironment.Item("OSDComputerName") = strNewDeviceName
	ZTIProcess = Success

	oLogging.CreateEntry "SetComputerName: End ZTIProcess", LogTypeInfo
End Function

</script>
</job>