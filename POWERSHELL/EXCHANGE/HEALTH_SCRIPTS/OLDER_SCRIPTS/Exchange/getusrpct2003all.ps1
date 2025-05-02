$snServerNames = @{ }
$mbSizes = @{ }
$mbStores =  @{ }
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(objectCategory=msExchExchangeServer)'
$searcher1 = $searcher.FindAll()
#Get Server Objects
foreach($server in $searcher1){ 
	$soServerObject = $server.getDirectoryEntry()
	$snServerNames.add([String]$soServerObject.distinguishedName,$soServerObject)
	#Get Mailbox Sizes
	$qrQueryresults = get-wmiobject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $soServerObject.Name 
	foreach ($mbMailbox in $qrQueryresults){
		$mbSizes.add([String]$mbMailbox.MailboxGUID,$mbMailbox)	
	}
}
$searcher.Filter = '(objectCategory=msExchPrivateMDB)'
$searcher2 = $searcher.FindAll()
foreach ($mailstore in $searcher2){ 
	$moMailStoreObject = $mailstore.getDirectoryEntry()
	$soServer = $snServerNames[[String]$moMailStoreObject.msExchOwningServer]
	$ffEdbFileFilter = "name='" + $moMailStoreObject.msExchEDBFile.ToString().replace("\","\\") + "'"
	$ffStmFileFilter = "name='" + $moMailStoreObject.msExchSLVFile.ToString().replace("\","\\") + "'"
	$mbEdbSize =get-wmiobject CIM_Datafile -filter $ffEdbFileFilter -ComputerName $soServer.Name
	$mbStmSize =get-wmiobject CIM_Datafile -filter $ffStmFileFilter -ComputerName $soServer.Name
	[int64]$csCombinedSize = [double]$mbEdbSize.FileSize + [int64]$mbStmSize.FileSize
	$mbStores.add([String]$moMailStoreObject.distinguishedName,[int64]$csCombinedSize)
}

$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(msExchHomeServerName=*)))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$searcher2 = $dfsearcher.FindAll()
foreach ($uaUsers in $searcher2){ 
     	$uaUserAccount = New-Object System.DirectoryServices.directoryentry 
        $uaUserAccount = $uaUsers.GetDirectoryEntry() 
	$gaGuidArray =  $uaUserAccount.msExchMailboxGuid.value
	$adGuid =  "{" + $gaGuidArray[3].ToString("X2") + $gaGuidArray[2].ToString("X2") + $gaGuidArray[1].ToString("X2") + $gaGuidArray[0].ToString("X2") + "-" + 
	$gaGuidArray[5].ToString("X2") + $gaGuidArray[4].ToString("X2") + "-"  + $gaGuidArray[7].ToString("X2") + $gaGuidArray[6].ToString("X2") + "-" +
	$gaGuidArray[8].ToString("X2") + $gaGuidArray[9].ToString("X2") + "-" + $gaGuidArray[10].ToString("X2") + $gaGuidArray[11].ToString("X2") + 
	$gaGuidArray[12].ToString("X2") + $gaGuidArray[13].ToString("X2") + $gaGuidArray[14].ToString("X2") + $gaGuidArray[15].ToString("X2") + "}"
        $mbsize = [double]$mbSizes[$adGuid].Size
	$divval = ($mbStores[$uaUserAccount.HomeMDB][0]/1024)/100
	$pcStore =  ($mbsize/$divval)/100
	$uaUserAccount.Name.ToString() + "," + "{0:#.00}" -f ($mbsize/1KB) + "," + "{0:P1}" -f $pcStore 
	}


