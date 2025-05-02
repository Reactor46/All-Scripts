
'#########################################################################
'  Script name:		XLS_SaveAsSYLK.vbs
'  Created on:		07/07/2011
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft Excel Document
'					and then saves the file in SYLK format.
'#########################################################################

Dim AppExcel 
Dim OpenWorkbook
Const xlSYLK = 2

Set AppExcel = CreateObject("Excel.Application")

AppExcel.Visible = True

Set OpenWorkbook = AppExcel.Workbooks.Open("C:\Concepts\Temp.xls")
	
OpenWorkbook.SaveAs "C:\Concepts\SYLK\Faktura", xlSYLK

OpenWorkbook.Close
Set OpenWorkbook = Nothing

AppExcel.Quit
Set AppExcel = Nothing