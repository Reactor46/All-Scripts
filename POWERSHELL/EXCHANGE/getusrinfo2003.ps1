param([String] $samaccountname) 
$root = [ADSI]'LDAP://RootDSE' 
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(samaccountname=" + $samaccountname + ")))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$srSearchResult = $dfsearcher.FindOne()
if ($srSearchResult -ne $null){
	$uoUserobject = $srSearchResult.GetDirectoryEntry()
	$msStore = [ADSI]("LDAP://" +  $uoUserobject.homemdb)
	$soServer = [ADSI]("LDAP://" + $msStore.msExchOwningServer)
	$sgStorageGroup = $msStore.psbase.Parent
	$ffEdbFileFilter = "name='" + $msStore.msExchEDBFile.ToString().replace("\","\\") + "'"
	$ffStmFileFilter = "name='" + $msStore.msExchSLVFile.ToString().replace("\","\\") + "'"
	$mbEdbSize =get-wmiobject CIM_Datafile -filter $ffEdbFileFilter -ComputerName $soServer.Name
	$mbStmSize =get-wmiobject CIM_Datafile -filter $ffStmFileFilter -ComputerName $soServer.Name
	[int64]$csCombinedSize = [double]$mbEdbSize.FileSize + [int64]$mbStmSize.FileSize
	$msFilter = "LegacyDN='" + $uoUserobject.legacyExchangeDN + "'"
	$mbsize = get-wmiobject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -filter $msFilter -ComputerName $soServer.Name
	$divval = ($csCombinedSize/1024)/100
	$pcStore =  ($mbsize.size/$divval)/100
	"User DisplayName : " + $uoUserobject.displayName
	"Mail Server : " + $soServer.Name
	"Exchange Version : " + $soServer.SerialNumber
	"Mailbox Store : " + $msStore.Name
	"Storage Group : " + $sgStorageGroup.Name
	"MailStore Size : " + "{0:#.00}" -f ($csCombinedSize/1GB) + " GB"
	"Mailbox Size : " +  "{0:#.00}" -f ($mbsize.Size/1KB) + " MB"
	"Percentage of Store Used by User : " + "{0:P1}" -f $pcStore

}

