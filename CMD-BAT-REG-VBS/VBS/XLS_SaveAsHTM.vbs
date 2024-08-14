
'#########################################################################
'  Script name:	XLS_SaveAsHTM.vbs
'  Created on:	4/29/2010
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft Excel Document
'								and then saves the file in HTML format.
'#########################################################################

Dim AppExcel 
Dim OpenWorkbook
Const xlsSaveAsHTML = 44

Set AppExcel = CreateObject("Excel.Application")

AppExcel.Visible = True

Set OpenWorkbook = AppExcel.Workbooks.Open("C:\Concepts\Faktura.xls")
	
OpenWorkbook.SaveAs "C:\Concepts\HTML\Faktura", xlsSaveAsHTML

OpenWorkbook.Close
Set OpenWorkbook = Nothing

AppExcel.Quit
Set AppExcel = Nothing