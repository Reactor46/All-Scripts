Param([string]$siteURL)

Function CreatePulseIntranetTermGroup() {

	if($siteURL -eq "") {
        throw "No Parameter 'Site URL' found. Please run .\TermStoreManagement 'Site URL'"
    }#end if

    $web = Get-SPWeb -Identity $siteURL

    $session = New-Object Microsoft.SharePoint.Taxonomy.TaxonomySession($siteURL)

    if ($session.TermStores.Count -lt 1) {
        throw "No term stores found. The Taxonomy Service is offline or missing"
    }#end if

    Write-Host `n
    Write-Host "Creating Pulse Intranet Managed Metadata Group"
    Write-Host "________________________________"

	[Microsoft.SharePoint.Taxonomy.TermStore]$store = $session.TermStores["Managed Metadata Service - Internal"]
	
	#create the group
	$store.CreateGroup("Pulse Intranet", [System.Guid]::NewGuid())
	$store.CommitAll()

	#create the termset
	$termGroup = $store.Groups["Pulse Intranet"]
	$termSet = $termGroup.CreateTermSet("Tags")
	$store.CommitAll()

    Write-Host `n
    Write-Host "Process Complete"
    Write-Host `n

}

CreatePulseIntranetTermGroup