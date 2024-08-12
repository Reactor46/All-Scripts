On Error Resume Next

Set objFS = CreateObject("Scripting.FileSystemObject") 'setup of files
Set objNewFileGPO = objFS.CreateTextFile("GPO.TXT")
Set objNewFileOUs = objFS.CreateTextFile("OUs.TXT")
set objNewFileSites = objFS.CreateTextFile("Sites.TXT")

Const ADS_SCOPE_SUBTREE = 2				' Default AD connections

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE


Set rootDSE = GetObject("LDAP://rootDSE")
DSEroot=rootDSE.Get("DefaultNamingContext")

arrA = split(dseroot,",")				' read default naming context for further use
b = ubound(arrA) 
j = len(arrA(b)) + 8
j0 = arrA(b) & ";0][LDAP"
j1 = arrA(b) & ";1][LDAP"
j2 = arrA(b) & ";2][LDAP"
j3 = arrA(b) & ";3][LDAP"

' Enumerating Group policies in domain from system\Policies container

objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & DSEroot & "' WHERE " & "objectCategory='groupPolicyContainer'"

Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	strPath = objRecordSet.Fields("ADsPath").Value
	Set objUser = GetObject(strPath)

	objnewFileGPO.Write objUser.CN & "," & objUser.DISPLAYNAME & ","

	if objuser.Flags = 1 Then		' Identification of GPO Status
			objNewFileGPO.Write "User Configuration Settings Disabled"
	else if objuser.Flags = 3 Then
			objNewFileGPO.Write "All Settings Disabled"
	else if objuser.Flags = 2 Then
			objNewFileGPO.Write "Computer Configuration Settings Disabled"
	else if objuser.Flags = 0 Then
			objNewFileGPO.Write "Enabled"
	End If
	END IF
	END IF
	END IF

	objNewFileGPO.Writeline
	objRecordSet.MoveNext
Loop

'--------------------------------------------------------------
' information retrieval of GPOs on Organizational Units in domain
'--------------------------------------------------------------

objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & DSEroot & "' WHERE " & "objectCategory='organizationalunit'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	strPath = objRecordSet.Fields("ADsPath").Value
	Set objOU = GetObject(strPath)
	objnewFileOUs.WriteLine objOU.name & "," & objOU.gplink
	objRecordSet.MoveNext
Loop

'---------------------------------------------------
'	information of GPOs on Domain
'---------------------------------------------------
	Set objOU = GetObject("LDAP://" & DSEROOT)
	objnewFileOUs.WriteLine objOU.name & "," & objOU.gplink

'---------------------------------------------------
'	GPO information of Sites
'---------------------------------------------------

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.Filter = Array("site")
 
For Each objSite In objSitesContainer
	objNewFileSites.Writeline objsite.cn & "," & objSite.Gplink

Next


objNewFileGPO.close
objNewFileOUs.close
objNewFileSites.close


wscript.echo "Parasing completed ... !!"


'---------------------------------------------------
'	Data santization
'---------------------------------------------------

Set objNewFileGPO = objFS.OpenTextFile("GPO.TXT",1)
set objNewFileW = objFS.CreateTextFile("result.txt",2)
set objNewFileG = objFS.CreateTextFile("resultG.txt",2)

Do until objNewFileGPO.AtEndofStream
	strlineC = split(objNewFileGPO.readline,",")
i=0
objNewFileW.Write strlineC(1) & "," & strlineC(2)
objNewFileg.Write strlineC(1) & "," & strlineC(2)
Set objNewFileOUs = objFS.OpenTextFile("OUs.TXT",1)

Do until objNewFileOUs.AtEndOfStream
	strline = objNewFileOUs.readline
	arrone = split(strline,",")

	for each arrtwo in arrone 
		if left(arrtwo,5) = "[LDAP" or left(arrtwo,j)= j0 or left(arrtwo,j)= j1 or left(arrtwo,j)= j2 or  left(arrtwo,j)= j3 Then
			arrthree = split(arrtwo,"=")
			for each arrfour in arrthree
				if left(arrfour,1) = "{" Then
					'objNewFileW.Write "," & arrfour
						if strlineC(0) = arrfour Then
							objNewFileg.Write "," & arrone(0) 'arrfour
							i=i+1
						End If
				End If
			Next
		End If
	next
loop
	objNewFileOUs.close

Set objNewFileSites = objFS.OpenTextFile("Sites.TXT",1)
Do until objNewFileSites.AtEndOfStream
	strline = objNewFileSites.readline
	arrone = split(strline,",")
	'objNewFileW.Write arrone(0)
	for each arrtwo in arrone 
		if left(arrtwo,5) = "[LDAP" or left(arrtwo,j)= j0 or left(arrtwo,j)= j1 or left(arrtwo,j)= j2 or  left(arrtwo,j)= j3 Then
			arrthree = split(arrtwo,"=")
			for each arrfour in arrthree
				if left(arrfour,1) = "{" Then
					'objNewFileW.Write "," & arrfour
						if strlineC(0) = arrfour Then
							objNewFileg.Write "," & arrone(0) 'arrfour
							i=i+1
						End If
				End If
			Next
		End If
	next
loop
	objNewFileSites.close

	objNewFileW.Write "," & "Linked in " & i & " places directly"
	objNewFileW.Writeline
	objNewFileg.WriteLine
loop


objNewFileGPO.close
objNewFilew.close
objNewFileg.close
