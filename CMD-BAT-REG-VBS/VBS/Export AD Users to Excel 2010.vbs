' Bind to RootDSE - this object is used to 
' get the default configuration naming context
set objRootDSE = getobject("LDAP://RootDSE")

set WshShell = CreateObject ("WScript.Shell")
desktop = WshShell.SpecialFolders ("Desktop")

' File name to export to
 strExportFile = desktop & "\ACU_AD-Directory.xlsx" 
'Delete existing file
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strExportFile) Then
Set aFile = fso.GetFile(strExportFile)
aFile.Delete
End If
strRoot = objRootDSE.Get("DefaultNamingContext")
' Filter for user accounts - could be modified to search for specific users,
' such as those with mailboxes, users in a certain department etc.
strfilter = "(&(objectCategory=Person)(objectClass=User))"
strAttributes = "displayName,ipPhone,telephoneNumber,mobile,homePhone,facsimileTelephoneNumber,mail,streetAddress,l,st,postalCode,co,c,title,department,company"
strScope = "subtree"

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")

cn.open "Provider=ADsDSOObject;"
cmd.ActiveConnection = cn
cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _
		   strAttributes & ";" & strScope
set rs = cmd.execute

' Use Excel COM automation to open Excel and create an excel workbook
set objExcel = CreateObject("Excel.Application")
set objWB = objExcel.Workbooks.Add
set objSheet = objWB.Worksheets(1)

' Copy Field names to header row of worksheet
For i = 0 To rs.Fields.Count - 1
	objSheet.Cells(1, i + 1).Value = rs.Fields(i).Name
	objSheet.Cells(1, i + 1).Font.Bold = True
	objSheet.Cells(1, i + 1).Interior.ColorIndex = 17	
Next

' Copy data to the spreadsheet
objSheet.Range("A2").CopyFromRecordset(rs)
' Sort by displayName
Set objRange = objSheet.UsedRange
Set objRange2 = objExcel.Range("A1")
objRange.Sort objRange2, 1, , , , , , 1
'Freeze Panes
objExcel.Windows(1).SplitColumn = 1
objExcel.Windows(1).SplitRow = 1
objExcel.Windows(1).FreezePanes = True
'Autofit
objRange.EntireColumn.Autofit() 
'AutoFilter
objRange.AutoFilter
' Save the workbook
objWB.SaveAs(strExportFile)
' Clean up
rs.close
cn.close
set objSheet = Nothing
set objWB =  Nothing
objExcel.Quit()
set objExcel = Nothing