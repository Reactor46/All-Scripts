'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 14/03/2009
' CreateDistributionGroups_FromExcel.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' This Script Will Get an Excel Or CSV File and Create Distribution Groups
' According to the File.
' You can decide the Type and destination OU for each Distribution Group in the File
' The Excel File (or CSV) need to be in the Following Order :
' Group Name, Destination OU, Type
' You can Either write to the file with or without headers - Change the IntRow Accordingly
' The Types that Are allowed are : Local, Global, Universal
' Destination OU can be any Format you like (but make sure it is a full path)
Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const LOG_FILE = "C:\New_Distribution_Groups.txt"

Sub CreateDistGroup (Name, LDAPOU, strType)
' This Sub Will Create a Distribution Group of any Type in the Destination OU
' Input  : Group Name, Destination OU LDAP (AdsPath), Group Type (Local, Global, Universal)
' Output : Creation of The Group in The Desired OU, Report to a Log File.
	Set objOU = GetObject(LDAPOU)
	Set objGroup = objOU.Create("Group", "cn=" & Name)
	objGroup.Put "sAMAccountName", Name
	objGroup.Put "groupType", strType
	objGroup.SetInfo
	If Err.Number = 0 Then
		objFile.WriteLine "Group " & Name & " was Created Succefully in " & LDAPOU
	Else
		objFile.WriteLine "Error Creating Group " & Name & vbNewLine _
			& vbTab & "Error Description : " & Err.Description
	End If
End Sub

Function OULDAPPath(strOU)
' This Function will Set the OU Path to the LDAP Form so the Script can work
' Input  : OU String in the Following Forms :
'			"LDAP://ou=Users,dc=contoso,dc=com"
'			"ou=Users,dc=contoso,dc=com"
'			"contoso.com/Users"
' Output : LDAP OU Path
Dim newOuPath
Dim i
Dim arrPath, arrDC
If InStr(UCase(strOU), "OU=") Then
	If InStr(StrOU, "LDAP://") Then
		OULDAPPath = strOU
	Else
		OULDAPPath = "LDAP://" & strOU
	End If
Else 
	If InStr(strOU, "/") Then
		arrPath = Split(strOU, "/")
		newOuPath = ""
		For i=UBound(arrPath) To 1 Step -1
			newOuPath = newOuPath & "OU=" & arrPath(i) & ","
		Next
		
		arrDC = Split(arrPath(0),".")
		For i=0 To UBound(arrDC)
			newOuPath = newOuPath & "DC=" & arrDC(i) & ","
		Next
		' Remove the Ending ','
		newOuPath = Left(newOuPath,Len(newOuPath)-1)
		OULDAPPath = "LDAP://" & newOuPath
	End If
End If
End Function

Function GetType(strType)
' This Function will determine the Group Type
' Input  : String Type
' Output : Group Type
Select Case strType
	Case "Local" : GetType = ADS_GROUP_TYPE_LOCAL_GROUP
	Case "Global" : GetType= ADS_GROUP_TYPE_GLOBAL_GROUP
	Case "Universal" : GetType = ADS_GROUP_TYPE_UNIVERSAL_GROUP
	Case Else GetType = ADS_GROUP_TYPE_UNIVERSAL_GROUP
End Select
End Function

Set objDialog = CreateObject("UserAccounts.CommonDialog")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objExcel = Createobject("Excel.Application")

' Locate Distribution Groups File
' The File Needs to be in the Folowing Order :
' Group Name, Destination OU (AdsPath - No LDAP:// in the Beginning)
' UsersGroup, "OU=Groups,OU=HelpDesk,DC=MyDomain,DC=Com"
objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
   FileLoc = objDialog.FileName
End If

' Output Location And Name

If objFso.FileExists (LOG_FILE) THEN
	Set objFile = objFso.GetFile (LOG_FILE)
	objFile.Delete
End If 
Set objFile = objFso.CreateTextFile (LOG_FILE, True)

objFile.WriteLine "Log Started : " & Now

' Get List Of Distribution Groups
objexcel.Workbooks.Open(FileLoc)

' Set IntRow to 1 if no Headers in the File
' Set IntRow to 2 if Headers Exist in the File
IntRow = 2

Do Until objExcel.cells(introw,1).value=""
	' Getting the Values from the Excel File
	strGroup = objExcel.cells(introw,1).value
	strOU = objExcel.cells(introw,2).value
	strType = objExcel.cells(introw,3).value
	' Creating the Group - Sepecify Name, Dest OU and Type
	CreateDistGroup strGroup,OULDAPPath(strOU), GetType(strType)
	IntRow = IntRow + 1
Loop

' Clean Up and End Log File
objFile.WriteLine "Log Ended : " & Now
objFile.Close
objExcel.Quit