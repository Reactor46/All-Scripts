'#########################################################################
'  Script name:		GetAllODBCDriver_XLS.vbs
'  Created on:		07/08/2011
'  Author:			Dennis Hemken
'  Purpose:			Returns the all installed ODBC Driver
'					in a new Excel Document
'#########################################################################

Const HKEY_LOCAL_MACHINE = &H80000002

Dim strComputer

    strComputer = "."
    fct_GetAllODBCDriver_XLS(strComputer)

Public Function fct_GetAllODBCDriver_XLS(strComputer)

Dim objRegistry
Dim strRegPath
Dim strAODBCDriverNames
Dim strAValueTypes
Dim strODBCDriverName
Dim strValue
Dim objExcel
Dim objRange
Dim lngRow

	Set objExcel = CreateObject("Excel.Application") 
 
	objExcel.Visible = True 
	objExcel.Workbooks.Add 
	lngRow = 1
	objExcel.Cells(lngRow, 1).Value = "Driver Name" 
	objExcel.Cells(lngRow, 2).Value = "Value" 
	
	objExcel.Cells(lngRow, 1).Font.Bold = True
	objExcel.Cells(lngRow, 2).Font.Bold = True

	strRegPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
	
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	
	objRegistry.EnumValues HKEY_LOCAL_MACHINE, strRegPath, strAODBCDriverNames, strAValueTypes

	For i = 0 to UBound(strAODBCDriverNames)
		lngRow = lngRow + 1
		
		strODBCDriverName = strAODBCDriverNames(i)
		objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strRegPath, strODBCDriverName, strValue
		objExcel.Cells(lngRow, 1).Value = strODBCDriverName
		objExcel.Cells(lngRow, 2).Value = strValue
	Next
	
	Set objRange = objExcel.Range("A1") 
	objRange.Activate 
	Set objRange = objExcel.ActiveCell.EntireColumn 
	objRange.Autofit() 
	Set objRange = objExcel.Range("B1") 
	objRange.Activate 
	Set objRange = objExcel.ActiveCell.EntireColumn 
	objRange.Autofit() 
	
End Function