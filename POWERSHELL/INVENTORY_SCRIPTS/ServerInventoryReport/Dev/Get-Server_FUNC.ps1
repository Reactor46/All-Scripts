Function Get-Server {
    [cmdletbinding(DefaultParameterSetName='FNBM')]
    Param (
    [Parameter(ParameterSetName='FNBM')]
    [switch]$FNBM,
    [Parameter(ParameterSetName='PHX')]
    [switch]$PHX,
    [Parameter(ParameterSetName='TST')]
    [switch]$TST,
    [Parameter(ParameterSetName='BIZ')]
    [switch]$BIZ    
       ) 
    Write-Verbose "Parameter Set: $($PSCmdlet.ParameterSetName)"
    Switch ($PSCmdlet.ParameterSetName) {
    'FNBM' {
        $searchroot = 'LDAP://DC=fnbm,DC=corp'
            }
    'PHX'{
        $searchroot = 'LDAP://DC=phx,DC=fnbm,DC=corp'
            }
    'TST' {
        $searchroot = 'LDAP://DC=creditoneapp,DC=tst'
            }
    'BIZ' {
        $searchroot = 'LDAP://DC=creditoneapp,DC=biz'
            }
    }
     
    $ldapFilter = '(&(objectCategory=computer)(OperatingSystem=Windows*Server*))'

    $Searcher = [adsisearcher]""
    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = $ldapFilter
    $Searcher.pagesize = 100
    $Searcher.sizelimit = 50000
    $Searcher.PropertiesToLoad.Add("name") | Out-Null
    $Searcher.sort.propertyname='name'
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.name}
}