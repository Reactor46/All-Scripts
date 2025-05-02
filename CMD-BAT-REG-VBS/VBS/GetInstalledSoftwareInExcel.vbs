
'#########################################################################
'  Script name:		GetInstalledSoftwareInExcel.vbs
'  Created on:		08/16/2011
'  Author:			Dennis Hemken
'  Purpose:			Get Installed Software from computers in an 
'					Microsoft Excel Document.
'					List Software Installed on Multiple Computers,
'					Display That Data in Excel and mark the
'					defined forbidden sofware
'#########################################################################

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 

Dim strComputer

Dim strDisplayName
Dim strQuietDisplayName 
Dim strInstallDate 
Dim strVersionMajor 
Dim strVersionMinor 
Dim strEstimatedSize

Dim objExcel
Dim objWorkbook
Dim objSheet
Dim lngRow
Dim lngRow2
Dim WshShell
Dim BtnCode

Dim intReturn
Dim intVMajorValue
Dim intVMinorValue
Dim intSizeValue
Dim strDisplayNameValue
Dim strQuietDisplayNameValue
Dim strDateValue
Dim objReg
Dim strKey
Dim strSubkey
Dim arrSubkeys

Dim arrForbiddenApps()
Dim arrComputer()

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open ("C:\Concepts\pc_installed_software.xls")
	Set objSheet=objExcel.Workbooks.Item(1)
	
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	objExcel.Visible = True
	lngRow2 = 2
	
	'strComputer = "." 
	strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 

	strDisplayName = "DisplayName" 
	strQuietDisplayName = "QuietDisplayName" 
	strInstallDate = "InstallDate" 
	strVersionMajor = "VersionMajor" 
	strVersionMinor = "VersionMinor" 
	strEstimatedSize = "EstimatedSize" 
 
 	BtnCode = WshShell.Popup("Start Searching for installed Applications?", 7, "WindowsOnlineActivation:", 4 + 32)

	If BtnCode = 6 Then
		'WScript.Echo "Get Installed Applications"
		Call fct_GetForbiddenApplications(objExcel, arrForbiddenApps)
		
		Call fct_GetComputer(objExcel, arrComputer)
		
		For each strComputer In arrComputer
			On Error Resume Next
			
			wscript.echo strComputer

			Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			If Err.Number <> 0 Then Call ErrorCheck(Err.Number, Err.Description)
			
			objReg.EnumKey HKLM, strKey, arrSubkeys 

			objExcel.Workbooks.Item(1).Sheets.Add 
		    objSheet.ActiveSheet.Name=strComputer

			objExcel.Cells(1, 1).Value = strComputer
			objExcel.Cells(1, 1).Font.Bold = True
			objExcel.Cells(1, 2).Value = strDisplayName
			objExcel.Cells(1, 2).Font.Bold = True
			objExcel.Cells(1, 3).Value = strQuietDisplayName
			objExcel.Cells(1, 3).Font.Bold = True
			objExcel.Cells(1, 4).Value = strInstallDate
			objExcel.Cells(1, 4).Font.Bold = True
			objExcel.Cells(1, 5).Value = "Version"
			objExcel.Cells(1, 5).Font.Bold = True
			objExcel.Cells(1, 6).Value = strEstimatedSize
			objExcel.Cells(1, 6).Font.Bold = True
			For Each strSubkey In arrSubkeys 
				intReturn = objReg.GetStringValue(HKLM, strKey & strSubkey, strDisplayName, strDisplayNameValue) 
				If intReturn <> 0 Then 
					objReg.GetStringValue HKLM, strKey & strSubkey, strQuietDisplayName, strQuietDisplayNameValue 
				End If 
				If Trim(strDisplayNameValue) <> "" Then
					
					objReg.GetStringValue HKLM, strKey & strSubkey, strInstallDate, strDateValue 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strVersionMajor, intVMajorValue 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strVersionMinor, intVMinorValue 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEstimatedSize, intSizeValue 
					
					if fct_IsAppForbidden(strDisplayNameValue, arrForbiddenApps) = True Then
						objExcel.Cells(lngRow2, 2).Font.Bold = True
						objExcel.Cells(lngRow2, 2).Interior.ColorIndex = 3
					End If
					
					objExcel.Cells(lngRow2, 2).Value = strDisplayNameValue
					objExcel.Cells(lngRow2, 3).Value = strQuietDisplayNameValue
					Call fct_SetDateValue (strDateValue)
					objExcel.Cells(lngRow2, 4).Value = strDateValue
					If intVMajorValue <> "" Then
						objExcel.Cells(lngRow2, 5).Value = "V." & intVMajorValue & "." & intVMinorValue
						If intSizeValue <> "" Then
							objExcel.Cells(lngRow2, 6).Value = intSizeValue / 1024 & " MB"
						End if						
					End if
					lngRow2 = lngRow2 +1
				End If
			Next
			Set objRange = objExcel.Range("B1") 
			objRange.Activate 
			Set objRange = objExcel.ActiveCell.EntireColumn 
			objRange.Autofit() 
			Set objRange = Nothing
			Set objReg = Nothing
			
		Next
		'Wscript.Echo "Searching for installed Applications finished"
	Else
		'Wscript.Echo "Searching for installed Applications not started"
	End If

objExcel.Save
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing

Public Function fct_SetDateValue (strDateValue)
Dim strYear
Dim strMonth
Dim strDay
	If strDateValue <> "" Then 
		strYear =  Left(strDateValue, 4) 
		strMonth = Mid(strDateValue, 5, 2) 
		strDay = Right(strDateValue, 2) 
	'some Registry entries have improper date format 
		On Error Resume Next  
		strDateValue = DateSerial(strYear, strMonth, strDay) 
	End If 
End Function

Public Function fct_GetForbiddenApplications(objExcel, arrForbiddenApps)
Dim nRow

	objExcel.Sheets(2).select
	nRow = 2
	Do Until objExcel.Cells(nRow,1).Value = ""
		Redim Preserve arrForbiddenApps(nRow-2)
		arrForbiddenApps(nRow-2) = objExcel.Cells(nRow,1).Value
		nRow = nRow +1
	Loop
	
End Function

Public Function fct_GetComputer(objExcel, arrComputer)
Dim nRow

	objExcel.Sheets(1).select
	nRow = 2
	Redim arrComputer(0)
	Do Until objExcel.Cells(nRow,1).Value = ""
		Redim Preserve arrComputer(nRow-2)
		arrComputer(nRow-2) = objExcel.Cells(nRow,1).Value
		nRow = nRow +1
	Loop
	' local
	If nRow <= 3 and arrComputer(0) = "" Then
		Redim Preserve arrComputer(0)
		arrComputer(0) = "."
	End If
	
End Function

Public Function fct_IsAppForbidden(strDisplayNameValue, arrForbiddenApps) 
Dim x
Dim nPos

	For x = 0 to UBound(arrForbiddenApps)
		nPos = InStr(1,UCase(strDisplayNameValue) ,UCase(arrForbiddenApps(x)))
		If nPos > 0 Then
			fct_IsAppForbidden = True
			Exit For
		End if
	Next
	
End Function

Sub ErrorCheck(sErrorCode, sErrorDescription) 

	Select Case sErrorCode 
	Case 462 
		MsgBox "Target computer is not found!" & vbCrLf & VbCrLf _ 
		& "Check that target computer is online and" & vbCrLf _ 
		& "it's local firewall is disabled.",64,"Computer not found" 
		
		Err.Clear 
	Case Else 
		MsgBox "Error occurred." & vbCrLf & VbCrLf _ 
		& "Error code is:" & sErrorCode & vbCrLf _ 
		& "Error description is: " & sErrorDescription,64,"Mystical error occurred" 

		Err.Clear 
	End Select 

End Sub 

