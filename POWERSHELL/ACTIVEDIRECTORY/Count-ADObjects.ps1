#Get Domain List
$objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$DomainList = @($objForest.Domains | Select-Object Name,DomainMode)
$fct_lvl_mode_Forest = $objForest.ForestMode

$array = @() 
 
#Act on each domain
foreach($Domain in $DomainList){
	$Domain_name = $Domain.Name
	$fct_lvl_mode_Domain = $Domain.DomainMode
	
	Write-Host "Checking $Domain_name" -fore red
	$ADsPath = [ADSI]"LDAP://$Domain_name"
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
	$objSearcher.Pagesize = 100000
	$objSearcher.SearchScope = "Subtree"

	#User
	$objSearcher.Filter = "(&(objectCategory=person)(objectClass=user))"
 	$colResults = $objSearcher.FindAll()
	$cnt_user = $colResults.count

	#Contact
	$objSearcher.Filter = "(objectClass=contact)"
 	$colResults = $objSearcher.FindAll()
	$cnt_contact = $colResults.count

	#Security Group
	$objSearcher.Filter = "(groupType:1.2.840.113556.1.4.803:=2147483648)"
 	$colResults = $objSearcher.FindAll()
	$cnt_group = $colResults.count

	#Distribution Group
	$objSearcher.Filter = "(&(objectCategory=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))"
 	$colResults = $objSearcher.FindAll()
	$cnt_dl = $colResults.count

	#Computer
	$objSearcher.Filter = "(&(objectCategory=computer)(!(operatingSystem=*server*)))"
 	$colResults = $objSearcher.FindAll()
	$cnt_computer = $colResults.count

	#Server
	$objSearcher.Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(userAccountControl:1.2.840.113556.1.4.803:=8192)))"
 	$colResults = $objSearcher.FindAll()
	$cnt_server = $colResults.count

	#DC
	$objSearcher.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))"
 	$colResults = $objSearcher.FindAll()
	$cnt_dc = $colResults.count

	#OU
	$objSearcher.Filter = "(objectCategory=organizationalUnit)"
 	$colResults = $objSearcher.FindAll()
	$cnt_ou = $colResults.count

	#GPO
	$objSearcher.Filter = "(objectCategory=groupPolicyContainer)"
 	$colResults = $objSearcher.FindAll()
	$cnt_gpo = $colResults.count

	$Properties = @{domain=$Domain_name;domain_mode=$fct_lvl_mode_Domain;forest_mode=$fct_lvl_mode_Forest;user=$cnt_user;contact=$cnt_contact;group=$cnt_group;distributionlist=$cnt_dl;workstation=$cnt_computer;server=$cnt_server;domaincontroller=$cnt_dc;ou=$cnt_ou;gpo=$cnt_gpo}
	$Newobject = New-Object PSObject -Property $Properties
	$array +=$newobject

} 
$array | ConvertTo-Csv -NoTypeInformation -Delimiter "," | Foreach-Object {$_ -replace '"', ''} | Out-File "C:\LazyWinAdmin\Firewall\ad_info.csv" -Encoding ASCII