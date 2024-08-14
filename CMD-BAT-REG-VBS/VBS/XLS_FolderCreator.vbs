'###################################################################################
'  Script name:		XLS_FolderCreator.vbs
'  Created on:		05.07.2011
'  Author:			Dennis Hemken
'  Purpose:			Opens an existing Microsoft Excel Document with
'					one parent-folder (in cell A2)
'					and a lot of child-folder names (begin in B2) for a batch.
'					This VBS creates automatically folders.
'###################################################################################


Dim objFSO
Dim objExcel
Dim objWorkbook
Dim WshShell
Dim BtnCode
Dim lngRow
Dim strFolderTarget
Dim strFolderTargetToCreate

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open ("C:\Concepts\FolderCreator.xls")

Set WshShell = WScript.CreateObject("WScript.Shell")

objExcel.Visible = True

lngRow = 2

strFolderTarget = objExcel.Cells(lngRow, 1).Value

BtnCode = WshShell.Popup("Create Folder ?", 7, "XLS_FolderCreator:", 4 + 32)

If BtnCode = 6 Then
	Wscript.Echo "Folder creation begins: " & strFolderTarget
	Do Until objExcel.Cells(lngRow,2).Value = ""
		strFolderTargetToCreate = strFolderTarget & "\" & objExcel.Cells(lngRow, 2).Value
		If objFSO.FolderExists(strFolderTarget) Then
			objFSO.CreateFolder strFolderTargetToCreate
		End If
		lngRow = lngRow + 1
	Loop
	Wscript.Echo "Folder creation finished: " & strFolderTarget
Else
	Wscript.Echo "Folders are not created: " & strFolderTarget
End If

objExcel.Quit
Set objFSO = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

