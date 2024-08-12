'################################################################
'##    DATE: 03-14-2013     #####     AUTHOR: Keith Shelton    ##
'################################################################
'#   This script was written to clean up orphaned home folders. #
'# It does this by polling Active Directory and generating a    #
'# list of folders that should exist based on the Home Folder   #
'# property. It then compares that list to the folders that     #
'# actually do exist.                                           #
'#   It will generate a text file that lists all of the         #
'# exceptions that were found.                                  #
'#   This script also has the ability to move all orphans to    #
'# another location for review.                                 #
'################################################################


Option Explicit

Dim objRootDSE, strDNSDomain, adoCommand, adoConnection
Dim strBase, strFilter, strAttributes, strQuery, adoRecordset
Dim strHomeDir, outFile, strMsg, arrOutFile, intIndex, outPath
Dim objFSO, strUsersFolder, objUsersFolder, objFolder, objFile
Dim objList, strDir, strTarget, blnTarget, WshShell

'################################################################
'# The following variables MUST be populated before this script #
'# is run:                                                      #
'# outFile = Location to write the exception text file to       #
'# strUsersFolder = This is the location that your users' home  #
'#                  folders are located.                        #
'#                                                              #
'# strTarget can be left blank if you do not intend to move the #
'#           orphaned folders. This will be the path to the     #
'#           folder that you will want to move the orphans to   #
'################################################################
strTarget = "Move_to_Path"
outFile="PATH_TO\REPORT.txt"
strUsersFolder = "PATH_TO_HOME_DIRS_FOLDER"

Set WshShell = WScript.CreateObject("Wscript.Shell")

' Create dictionary object to track home folders.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare

' Add all subfolders of share to dictionary object.
Set objFSO = CreateObject("Scripting.FileSystemObject")

'See if strTarget has been entered and if so, valid
If objFSO.FolderExists(strTarget) Then
	blnTarget = True
Else
	blnTarget = False
	strMsg = "The folder, " & strTarget & ", does not exist or could not be found." & vbCrLf & "The folders will not be moved."
	WshShell.Popup strMsg, 5, "Information"
End If

'Verify exception report location is valid
arrOutFile = Split(outFile,"\")
intIndex = UBound(arrOutFile)
outPath = Left(outFile, Len(outFile)-Len(arrOutFile(intIndex)))
If objFSO.FolderExists(outPath) = False Then
	strMsg = "The path (" & outPath & ") does not exist or could not be found." & vbCrLf & "The script will now exit."
	WshShell.Popup strMsg, 15, "Error"
	WScript.Quit
End If

'Verify that the folder containing users' home directories is valid.
If objFSO.FolderExists(strUsersFolder) = False Then
	strMsg = "The path (" & strUsersFolder & ") does not exist or could not be found." & vbCrLf & "The script will now exit."
	WshShell.Popup strMsg, 15, "Error"
	WScript.Quit
End If

Set objUsersFolder = objFSO.GetFolder(strUsersFolder)
For Each objFolder In objUsersFolder.SubFolders
	' Key value is the full folder path,
	' item value is the folder name.
	objList(objFolder.Path) = objFolder.Name
Next

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Search for all users with home directories.
strFilter = "(&(objectCategory=person)(objectClass=user)" & "(homeDirectory=*))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "homeDirectory"

' Construct the LDAP query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Run the query.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False
Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
	' Retrieve values.
	strHomeDir = adoRecordset.Fields("homeDirectory").Value
	If (objList.Exists(strHomeDir) = True) Then
		' This folder is assigned as home directory to a user.
		' Flag as used by blanking out the folder name.
		objList(strHomeDir) = ""
	End If
	adoRecordset.MoveNext
Loop

' Enumerate home directories in the share.
For Each strDir In objList.Keys
	' Check if assigned to a user.
	If (objList(strDir) <> "") Then
		' This folder is not assigned.
		'Uncomment the following line if you want results echoed to the screen
		'WScript.Echo strDir
		'Write line to output file
		If (objFSO.FileExists(outFile)) Then
			Set objFile = objFSO.OpenTextFile(outFile,8,True)
			objFile.WriteLine(strDir)
			objFile.Close
		Else
			Set objFile = objFSO.CreateTextFile(outFile,True)
			objFile.Write strDir & vbCrLf
			objFile.Close
		End If
		
		' Move the folder.
		If blnTarget = True Then
			objFSO.CopyFolder strDir, strTarget & "\" & objList(strDir)
			objFSO.DeleteFolder strDir, TRUE
		End If
		
	End If
Next

' Clean up.
adoRecordset.Close
adoConnection.Close