# Load the SharePoint PowerShell module (if not already loaded)
#if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
#    Add-PSSnapin Microsoft.SharePoint.PowerShell
#}

$url = 'https://prod-pulse.kscpulse.com/'
$termStoreName = 'Managed Metadata Service - Internal'
$termGlobalGroupName = 'Old Pulse Navigation'
$termSetName = 'Global Navigation'

function Get-NavStore() {
    $web = Get-SPWeb $url
    $site = $web.Site
    $taxSession = Get-SPTaxonomySession -Site $site
    return $taxSession.TermStores[$termStoreName]
}

$termStore = Get-NavStore

function Get-NavTerms() {
    #$termStore = Get-NavStore
    $termGroup = $termStore.Groups[$termGlobalGroupName]
    $termSet = $termGroup.TermSets[$termSetName]
    return $termSet.Terms
}



$NavTerms = Get-NavTerms
#$newBaseUrl = $null

    ForEach($Nav in $NavTerms){
        
        #$Url = $Nav.GetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl")
        $Linkurl = $Nav.LocalCustomProperties._Sys_Nav_SimpleLinkUrl
        #$Nav
        if ($Linkurl -ne $null){
        if ($Linkurl.StartsWith("https://pulse.kscpulse.com")){
        $newUrl = $Linkurl.Replace("https://pulse.kscpulse.com","")
        # Update the term's URL to be relative to the new base URL
        Write-Host $Linkurl - $newUrl
        
        #$Nav.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", $newUrl) 
        }
        # Commit the changes
        #$termStore.CommitAll()
        
        }
    }