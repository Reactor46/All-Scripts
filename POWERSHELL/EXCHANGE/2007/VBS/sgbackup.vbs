set objServer = CreateObject("CDOEXM.ExchangeServer")
set objmbd = createobject("CDOEXM.MailboxStoreDB")
Set objStorageGroup = CreateObject("CDOEXM.StorageGroup")
strServerName = "servername"
objServer.DataSource.Open strServerName
For Each StorageGroup In objServer.StorageGroups
	if instr(StorageGroup,"Recovery Storage Group") = 0 then
		objStorageGroup.datasource.open StorageGroup
		for each Mailstore in objStorageGroup.MailboxStoreDBs
			objmbd.datasource.open "LDAP://" & mailstore
			Wscript.echo "Mail Store " & objmbd.Name
			Wscript.echo "Last Backed Up : " & objmbd.LastFullBackupTime
			wscript.echo
		next
		for each Pubstore in objStorageGroup.MailboxStoreDBs
			objmbd.datasource.open "LDAP://" & Pubstore
			Wscript.echo "Public Folder Store " & objmbd.Name
			Wscript.echo "Last Backed Up : " & objmbd.LastFullBackupTime
			wscript.echo
		next 
	end if
Next
