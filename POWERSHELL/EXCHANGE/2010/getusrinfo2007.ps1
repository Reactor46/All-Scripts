param([String] $samaccountname) 
$MailboxStats = get-mailboxstatistics $samaccountname
$soServer = get-ExchangeServer $MailboxStats.ServerName
$moMailStore = get-mailboxdatabase $MailboxStats.Database
$ffEdbFileFilter = "name='" + $moMailStore.edbfilepath.ToString().Replace("\","\\") + "'"
$mbEdbSize = get-wmiobject CIM_Datafile -filter $ffEdbFileFilter -ComputerName $soServer.Name
$divval = $mbEdbSize.fileSize/100
$pcStore =  ($MailboxStats.TotalItemSize.Value/$divval)/100
"User DisplayName : " + $MailboxStats.DisplayName
"Mail Server : " + $MailboxStats.ServerName
"Exchange Version : " + $soServer.AdminDisplayVersion
"Mailbox Store : " + $MailboxStats.DatabaseName
"Storage Group : " + $MailboxStats.StorageGroupName
"MailStore Size : " + "{0:#.00}" -f ($mbEdbSize.fileSize/1GB) + " GB"
"Mailbox Size : " + $MailboxStats.TotalItemSize.Value.ToMB() + " MB"
"Percentage of Store Used by User : " + "{0:P1}" -f $pcStore
