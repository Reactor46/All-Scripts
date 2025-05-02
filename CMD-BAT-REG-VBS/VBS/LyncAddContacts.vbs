'LyncAddContacts.vbs
'Version 1.1, 01/08/2011
'Jeff Guillet, jeff@expta.com
'The EXPTA {blog} - http://www.expta.com

'Make sure we're using CScript
If Instr(lcase(Wscript.FullName), "wscript") Then ShowSyntax

'Validate script parameters
If (Wscript.Arguments.Count = 0) OR (Wscript.Arguments.Count > 3) Then ShowSyntax
If Lcase(Wscript.Arguments(0)) = "/backup" Then
	modeBackup = True
	If Wscript.Arguments.Count = 1 Then ShowSyntax
	strTarget = Wscript.Arguments(1)
	If Wscript.Arguments.Count = 3 Then strSQLServer = Wscript.Arguments(2)
ElseIf Lcase(Wscript.Arguments(0)) = "/import" Then
	modeImport = True
	If Wscript.Arguments.Count = 1 Then ShowSyntax
	strTarget = Wscript.Arguments(1)
	If Wscript.Arguments.Count = 3 Then strSQLServer = Wscript.Arguments(2)
Else
	modeExport = True
	strSIPUser = Wscript.Arguments(0)
	If Right(strSIPUser, 1) = "?" OR Instr(strSIPUser, "@") = 0 Then ShowSyntax
	If Wscript.Arguments.Count = 2 Then strSQLServer = Wscript.Arguments(1)
End If

'Get CN for root domain
Set objRoot = getobject("LDAP://RootDSE")
domainName = objRoot.get("defaultNamingContext")
Set WshShell = WScript.CreateObject ("WScript.Shell")

'Check for dbimpexp.exe
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FileExists("dbimpexp.exe") Then
	Wscript.Echo "The DBIMPEXP utility was not found in the current directory."
	Wscript.Echo "Copy DBIMPEXP.EXE from the \SUPPORT folder of the Lync Server 2010"
	Wscript.Echo "installation media to the same folder where this tool resides."
	Wscript.Quit
End If

If modeBackup Then
	'Backup all user information in the Lync database
	If Len(strSQLServer) Then
		'Export all user data using DBIMPEXP
		'Lync Server Enterprise Edition. requires SQL server parameter
		r = WshShell.Run("dbimpexp.exe /hrxmlfile:" & strTarget & " /sqlserver:" & _
		   strSQLServer & " /restype:user", 1, True)
	Else
		'Lync Server Standard Edition
		r = WshShell.Run("dbimpexp.exe /hrxmlfile:" & strTarget & " /restype:user", 1, True)
	End If
	Wscript.Echo "Backup complete."
	Wscript.Quit
End If

If modeExport Then
	'Export contact groups and contacts from the specified user

	'Search the default domain for the source user and validate that it is SIP enabled
	If CheckForSIPUser(strSIPUser) = False Then
		Wscript.Echo "SIP user '" & strSIPUser & "' not found."
		Wscript.Quit
	End If

	'Export the target user's contact groups and contacts using dbimpexp
	Wscript.Echo "Exporting contact groups and contacts from SIP user: " & strSIPUser
	Wscript.Echo
	If Len(strSQLServer) Then
		'Export the contact groups and contacts using DBIMPEXP
		'Lync Server Enterprise Edition. requires SQL server parameter
		r = WshShell.Run("dbimpexp.exe /hrxmlfile:TempExport.xml /sqlserver:" & _
		   strSQLServer & " /restype:user /user:" & strSIPUser, 1, True)
	Else
		'Lync Server Standard Edition
		r = WshShell.Run("dbimpexp.exe /hrxmlfile:TempExport.xml /restype:user /user:" _
		   & strSIPUser, 1, True)
	End If

	'Truncate the exported unicode XML file to include only ContactGroups and Contacts
	If fso.FileExists("TempExport.xml") Then
		Set origFile = fso.OpenTextFile("TempExport.xml", 1, False, True)
		On Error Resume Next
		Set newFile = fso.OpenTextFile("Export.xml", 2, True, True)
		If Err Then
			Wscript.Echo "Unable to create Export.XML file. Make sure you have rights to write to this"
			Wscript.Echo "folder and that you are running from an elevated Command Prompt."
			Wscript.Quit
		End If
		On Error Goto 0
		Do Until origFile.AtEndOfStream
			rLine = origFile.ReadLine
			endOfFile = Instr(rLine, "</Contacts>")
			If endOfFile Then
				Exit Do
			Else
				newFile.WriteLine(rLine)
			End If
		Loop
		If endOfFile = 0 Then
			Wscript.Echo "This user has no contacts to export!"
			Wscript.Quit
		End If
		newFile.WriteLine Left(rLine, endOfFile + 10) & "</HomedResource></HomedResources>"
		origFile.Close
		newFile.Close
		'fso.DeleteFile "TempExport.xml"
	Else
		Wscript.Echo "ERROR! TempExport.xml file not found"
		Wscript.Echo "Check that you have permission to write to this folder"
		Wscript.Echo "and that the parameters you entered above are correct."
		Wscript.Quit
	End If

	'Done processing
	Wscript.Echo "Completed processing.  Use the /import switch to import the"
	Wscript.Echo "contact groups and contacts to target users."
	Wscript.Quit
Else
	'Import exported contact groups and contacts to target user(s)
	If NOT fso.FileExists("Export.xml") Then ShowSyntax
	If Instr(Ucase(strTarget), "DC=") Then
		'Locate the SIP users in the target OU
		Set cn = createobject("ADODB.Connection")
		Set cmd = createobject("ADODB.Command")
		Set rs = createobject("ADODB.Recordset")
		cn.Open "Provider=ADsDSOObject;"
		cmd.Activeconnection = cn
		cmd.CommandText="SELECT msRTCSIP-PrimaryUserAddress FROM 'LDAP://" _
		   & strTarget & "' WHERE msRTCSIP-PrimaryUserAddress = '*'"
		cmd.Properties("Page Size") = 2000
		On Error Resume Next
		Set rs = cmd.Execute
		If Err Then
			Wscript.Echo "Error locating Distinguished Name '" & strTarget & "'"
			Wscript.Quit
		End If
		On Error Goto 0
		Wscript.Echo "Importing contacts to all SIP users in the following OU:"
		Wscript.Echo vbTab & Chr(34) & strTarget & Chr(34)
		Wscript.Echo
		'Loop through each SIP user, importing contact groups and contacts to each one
		Do Until rs.EOF
			sIPUser = Mid(rs.Fields("msRTCSIP-PrimaryUserAddress").Value, 5)
			Wscript.Echo "Importing to SIP user: " & sIPUser & "..."
			ImportContacts sIPUser, strSQLServer
			rs.MoveNext
		Loop
		If sIPUser = "" Then
			Wscript.Echo "No SIP users were found in the OU: " & strTarget
			Wscript.Quit
		End If
	Else
		'Import contact groups and contacts to target SIP user
		If CheckForSIPUser(strTarget) = False Then 
			Wscript.Echo "SIP user '" & strTarget & "' not found."
			Wscript.Quit
		End If
		Wscript.Echo "Importing contacts to SIP user: " & strTarget
		Wscript.Echo
		ImportContacts strTarget, strSQLServer
	End If
	Wscript.Echo
	Wscript.Echo "Completed importing contact groups and contacts."
	Wscript.Echo "Users must sign out and back into Lync Server to see the updates."
	Wscript.Quit
End If

Sub ImportContacts(strTarget, strSQLServer)
	'Write a new Import.xml file using the strTarget SIP address
	Set origFile = fso.OpenTextFile("Export.xml", 1, False, True)
	Set newFile = fso.OpenTextFile("Import.xml", 2, True, True)
	Do Until origFile.AtEndOfStream
		rLine = origFile.ReadLine
		editLoc = Instr(rLine, "UserAtHost=")
		If editLoc Then 
			'Replace the SIP address in Export.xml with strTarget
			endQuote = Instr(editLoc + 12, rLine, Chr(34))
			rLine = Left(rLine, editLoc + 11) & strTarget & Mid(rLine, endQuote)
		End If
		self = Instr(Lcase(rLine), "<contact buddy=" & Chr(34) & Lcase(strTarget))
		If self > 0 Then
			'Skip the target as a contact
			firstPart = Left(rLine, self - 1)
			rLine = Mid(rLine, self + 1)
			Do
				endContacts = Instr(rLine, "</Contacts>")
				If endContacts Then
 					rLine = firstPart & Mid(rLine, endContacts)
					Exit Do
				End If
				endContact = Instr(rLine, "</Contact>")
				If endContact Then
					rLine = firstPart & Mid(rLine, endContact + 10)
				Else
					rLine = origFile.ReadLine
				End If
			Loop Until endContact
		End If
		newFile.WriteLine(rLine)
	Loop
	origFile.Close
	newFile.Close
	
	'Import the contact groups and contacts using dbimpexp
	Set WshShell = WScript.CreateObject ("WScript.Shell")
	If Len(strSQLServer) Then
		'Lync Server Enterprise Edition
		r = WshShell.Run("dbimpexp.exe /import /hrxmlfile:Import.xml /sqlserver:" _
		   & strSQLServer & " /restype:user", 1, True)
	Else
		'Lync Server Standard Edition
		r = WshShell.Run("dbimpexp.exe /import /hrxmlfile:Import.xml /restype:user", 1, True)
	End If
End Sub

Sub ShowSyntax
	'Show usage, notes and examples
	Msg = "This tool exports Lync contact groups and contacts from a specified SIP user and "
	Msg = Msg & "imports them to a single SIP user, or all SIP users in a target OU." _
	   & vbCRLF & vbCRLF
	Msg = Msg & "Backup usage:" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs /backup filename.xml [SQL server host name]" _
	   & vbCRLF & vbCRLF
	Msg = Msg & "Export usage:" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs SIPAddress [SQL server host name]" _
	   & vbCRLF & vbCRLF
	Msg = Msg & "Import usage:" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs /import SIPAddress | distinguished"
	Msg = Msg & "  name of OU [SQL server host name]" _
	   & vbCRLF & vbCRLF
	Msg = Msg & "Note: The SQL server host name parameter is necessary for Lync Enterprise Edition." _
	   & vbCRLF & vbCRLF
	Msg = Msg & "This tool requires the DBIMPEXP.exe tool located on the Lync Server 2010 installation media "
	Msg = Msg & "in the \SUPPORT folder. Copy DBIMPEXP.exe to the folder where this tool resides." _
	  & vbCRLF & vbCRLF
	Msg = Msg & "Usage Examples:" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs source@expta.com" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs source@expta.com sql.expta.com" _
	  & vbCRLF & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs /import jeff@expta.com" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs /import jeff@expta.com sql.expta.com" & vbCRLF
	Msg = Msg & "  CScript LyncAddContacts.vbs /import CN=Users,DC=expta,DC=com"
	MsgBox Msg, 0, "LyncAddContacts.vbs"
	Wscript.Quit
End Sub

Function CheckForSIPUser(strSIPUser)
	'Search Active Directory for matching SIP user
	Set cn = createobject("ADODB.Connection")
	Set cmd = createobject("ADODB.Command")
	Set rs = createobject("ADODB.Recordset")
	cn.Open "Provider=ADsDSOObject;"
	cmd.Activeconnection = cn
	cmd.CommandText="SELECT ADsPath FROM 'LDAP://" & domainName _
	   & "' WHERE msRTCSIP-PrimaryUserAddress = 'sip:" & strSIPUser & "'"
	Set rs = cmd.Execute
	If rs.EOF Then
		CheckForSIPUser = False
	Else
		CheckForSIPUser = True
	End If
End Function
