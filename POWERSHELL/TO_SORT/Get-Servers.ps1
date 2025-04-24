Function Get-Servers {
    [cmdletbinding(DefaultParameterSetName='All')]
    Param (
        [parameter(ParameterSetName='WebSupport')]
        [switch]$WebSupport,
        [parameter(ParameterSetName='Sitecore')]
        [switch]$SiteCore
    )
    Write-Verbose "Parameter Set: $($PSCmdlet.ParameterSetName)"
    Switch ($PSCmdlet.ParameterSetName) {
        'All' {
            $ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"
        }
        'WebSupport' {
            $ldapFilter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*)(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com))"
        }
        'SiteCore' {
            $ldapFilter = "(&(objectCategory=computer)(OperatingSystem=Windows*Server*)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com))"
        }
      }
    $searcher = [adsisearcher]""
    $Searcher.Filter = $ldapFilter
    $Searcher.pagesize = 10
    $searcher.sizelimit = 5000
    $searcher.PropertiesToLoad.Add("name") | Out-Null
    $Searcher.sort.propertyname='name'
    $searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach {
        $_.Properties.name
    }
}
