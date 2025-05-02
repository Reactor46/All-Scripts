set iExs = createobject("CDOEXM.ExchangeServer")
set isg = createobject("CDOEXM.StorageGroup")
set ipsDB = createobject("CDOEXM.PublicStoreDB")
set imbxDB = createobject("CDOEXM.MailboxStoreDB")
set ift = createobject("CDOEXM.FolderTree")
set WshShell = createobject("Wscript.shell")
wscript.echo Chr(13) + Chr(13) + Chr(13) + "================================="

Dim server_nm 
server_nm = WshShell.ExpandEnvironmentStrings("%computername%")
iExs.datasource.Open server_nm
  
Wscript.echo "ExchangeVersion = " + iExs.ExchangeVersion
Wscript.echo "DirectoryServer = " + iExs.DirectoryServer
Wscript.echo 

For Each storegroup In iExs.StorageGroups
    if instr(storegroup,"CN=Recovery Storage Group,") = 0 then
   	 isg.DataSource.Open storegroup
   	 For Each pubstore In isg.PublicStoreDBs
		  ipsDB.DataSource.Open pubstore
    	  	  wscript.echo ipsDB.name + Chr(13)
		  Wscript.echo "Lastbackup on : " & ipsDB.LastFullBackupTime
		  wscript.echo 
  	  Next 'public store
  	  For Each mbx In isg.MailboxStoreDBs
  	    	  imbxDB.DataSource.Open mbx
  	   	  wscript.echo imbxDB.name + Chr(13)
		  wscript.echo "LastBackup on : " & imbxdb.LastFullBackupTime
		  wscript.echo 
 	   Next 'mailbox
    end if
 Next 'storage group


  Set iExs = Nothing
  Set isg = Nothing
  Set ipsDB = Nothing
  Set imbxDB = Nothing
  Set ift = Nothing

