Function Get-Server {
Param (
    [parameter()]
    [ValidateSet('Contoso','PHX','TST','BIZ')]
    [string]$Domain = 'Contoso'
)
    Switch($Domain){
    {$_ -eq 'PHX'} {$Searchroot = 'LDAP://DC=phx,DC=contoso,DC=com'}
    {$_ -eq 'BIZ'} {$Searchroot = 'LDAP://DC=CREDITONEAPP,DC=BIZ'}
    {$_ -eq 'TST'} {$Searchroot = 'LDAP://DC=CREDITONEAPP,DC=TST'}
    {$_ -eq 'Contoso'} {$Searchroot = 'LDAP://DC=contoso,DC=com'}
                   }
    $Searcher = [ADSISEARCHER]""
    $Searcher.SearchRoot = $SearchRoot
    $Searcher.Filter = "(&(&(objectCategory=computer)(OperatingSystem=Windows*Server*)(!(useraccountcontrol:1.2.840.113556.1.4.803:=2))))"
    $Searcher.pagesize = 10
    $Searcher.sizelimit = 5000
    $searcher.PropertiesToLoad.Add("name") | Out-Null
    $Searcher.Sort.PropertyName='name'
    $Searcher.Sort.Direction = 'Ascending'
    $Searcher.FindAll() | ForEach-Object {$_.Properties.name}
}