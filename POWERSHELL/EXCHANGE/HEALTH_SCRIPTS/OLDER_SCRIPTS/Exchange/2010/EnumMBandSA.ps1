$root = [ADSI]'LDAP://RootDSE' 
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(&(&(&(mailnickname=*)(objectCategory=person)(objectClass=user))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$dfsearcher.PropertiesToLoad.Add("msExchMailboxSecurityDescriptor")
$srSearchResult = $dfsearcher.FindAll()
foreach ($emResult in $srSearchResult) {
	$uoUserobject = New-Object System.DirectoryServices.directoryentry
	$uoUserobject = $emResult.GetDirectoryEntry()
	$emProps = $emResult.Properties
	[byte[]]$DaclByte = $emProps["msexchmailboxsecuritydescriptor"][0]
	$adDACL = new-object System.DirectoryServices.ActiveDirectorySecurity
	$adDACL.SetSecurityDescriptorBinaryForm($DaclByte)
	$mbRightsacls =$adDACL.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])
	"Mailbox - " +  $uoUserobject.DisplayName
	foreach ($ace in $mbRightsacls){
		if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){		
			
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
			        $aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname
				}
				"	ACE UserName : " + $AceName
				""
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::CreateChild){ 
					"		Full Mailbox Access"}
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::WriteOwner -ne 0){ 
					"		Take Ownership"}
			        If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::WriteDacl){ 
					"		Modify User Attributes"}
			        If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::ListChildren){ 
					"		Is mailbox primary owner of this object"}
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::Delete){ 
					"		Delete mailbox storage"}
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::ReadControl){ 
					"		Read permissions"}

		}
	}
	$Sendasacls = $uoUserobject.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])|? {$_.ObjectType -eq 'ab721a54-1e2f-11d0-9819-00aa0040529b'}
	$Recieveasacls = $uoUserobject.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])|? {$_.ObjectType -eq 'ab721a56-1e2f-11d0-9819-00aa0040529b'}
	if ($Sendasacls -ne $null){
		foreach ($ace in $Sendasacls)
		{
			if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
				$aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname
				}
								""
				"	ACE UserName : " + $AceName
				"		Send As Rights"
			}
	
		}
	}
	if ($Recieveasacls -ne $null){
		foreach ($ace in $Recieveasacls)
		{
			if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
				$aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname
				}
								""
				"	ACE UserName : " + $AceName
				"		Recieve As Rights"
			}
		}
	}
}


