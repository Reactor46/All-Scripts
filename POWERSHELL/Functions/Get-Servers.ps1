    Param (
    [parameter(Mandatory=$False)]
    [ValidateSet('CORP','SVC','RES','PROD')]
    [string]$Domain = 'CORP'
)

Function Get-Servers {
    $Searcher = [adsisearcher]""
    If ($Domain = "SVC"){$searchroot = [ADSI]"LDAP://DC=svc,DC=prod,DC=vegas,DC=com"}
    elseif ($Domain = "PROD"){$searchroot = [ADSI]"LDAP://DC=prod,DC=vegas,DC=com" }
    elseif ($Domain = "RES"){$searchroot = [ADSI]"LDAP://DC=res,DC=vegas,DC=com"}
    else {$searchroot = [ADSI]"LDAP://DC=corp,DC=vegas,DC=com"}
    $Searcher.SearchRoot = $searchroot
    $Searcher.Filter = '(&(objectCategory=computer)(OperatingSystem=Windows*Server*))'
    $Searcher.pagesize = 10000
    $Searcher.sizelimit = 50000
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.dnshostname}
}