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
     
    $Searcher = [adsisearcher]""
    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = '(&(objectCategory=computer)(OperatingSystem=Windows*Server*))'
    $Searcher.pagesize = 10000
    $Searcher.sizelimit = 50000
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.dnshostname}
}