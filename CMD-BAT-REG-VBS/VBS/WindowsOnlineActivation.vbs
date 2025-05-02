'#########################################################################
'  Script name:		WindowsOnlineActivation.vbs
'  Created on:		07/15/2011
'  Author:			Dennis Hemken
'  Purpose:			Read an Excelfile with IP-adresses or comuternames
'					and activate the Windows Operation System
'#########################################################################

Dim strComputer
Dim objExcel
Dim objWorkbook
Dim WshShell
Dim lngRow
Dim BtnCode
Dim strProductID
Dim strActivationRequired
Dim strServerName

    strComputer = "."
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open ("C:\Concepts\Activate_Windows.xls")
    Set WshShell = WScript.CreateObject("WScript.Shell")
    
    objExcel.Visible = True
	lngRow = 2
	
	BtnCode = WshShell.Popup("Start Windows Activation?", 7, "WindowsOnlineActivation:", 4 + 32)

	If BtnCode = 6 Then
		Wscript.Echo "Windows activation begins!"
		
		Do Until objExcel.Cells(lngRow,1).Value = ""
			strComputer = objExcel.Cells(lngRow, 1).Value
			Call fct_WindowsOnlineActivation(strComputer, strProductID, strActivationRequired, strServerName)
			objExcel.Cells(lngRow, 2).Value = strServerName
			objExcel.Cells(lngRow, 3).Value = strProductID
			objExcel.Cells(lngRow, 4).Value = strActivationRequired
			lngRow = lngRow + 1
		Loop
		Wscript.Echo "Windows activation finished"
	Else
		Wscript.Echo "Windows activation not started"
	End If

objExcel.Save
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing




Public Function fct_WindowsOnlineActivation(strComputer, strProductID, strActivationRequired, strServerName)

Dim objWMIService
Dim colWindowsProducts

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colWindowsProducts = objWMIService.ExecQuery ("Select * from Win32_WindowsProductActivation")

	For Each objWindowsProduct in colWindowsProducts
		ObjWindowsProduct.ActivateOnline()
		strProductID = ObjWindowsProduct.ProductID
		strActivationRequired = ObjWindowsProduct.ActivationRequired
		strServerName = ObjWindowsProduct.ServerName
	Next

End Function