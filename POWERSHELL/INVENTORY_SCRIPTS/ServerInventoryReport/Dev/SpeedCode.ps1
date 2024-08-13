
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50
51
52
53
54
55
56
57
58
59
60
61
62
63
64
65
66
67
68
69
70
71
72
73
74
75
76
77
78
79
80
81
82
83
84
85
86
87
88
89
90
91
92
93
94
95
96
97
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()
 
Set-Location -Path 'C:\demo'
 
Add-Type -AssemblyName System.DirectoryServices.Protocols
Import-Module -Name .\S.DS.P.psd1
Add-PSSnapin -Name 'Quest.ActiveRoles.ADManagement'
 
$searcher = [adsisearcher]'(&(objectclass=user)(objectcategory=person))'
$searcher.SearchRoot = 'LDAP://DC=domain,DC=com'
$searcher.PageSize = 1000
$searcher.PropertiesToLoad.AddRange(('samaccountname'))
 
function Get-QueryResult
{
    [CmdletBinding()]
    
    Param
    (
        [Parameter(Mandatory=$true)]
        [int]$Id
    )
 
    switch ($id)
    {
        1    { ( Get-ADUser -Filter 'objectClass -eq "user" -and objectCategory -eq "person"' -SearchBase 'DC=domain,DC=com' -Properties SamAccountName).SamAccountName }
        2    { ( Get-ADUser -LDAPFilter '(&(objectclass=user)(objectcategory=person))' -SearchBase 'DC=domain,DC=com' -Properties SamAccountName).SamAccountName }
        3    { ( Get-ADObject -Filter 'objectCategory -eq "person" -and objectClass -eq "user"' -SearchBase 'DC=domain,DC=com' -Properties SamAccountName).SamAccountName }
        4    { ( Get-ADObject -LDAPFilter '(&(objectclass=user)(objectcategory=person))' -SearchBase 'DC=domain,DC=com' -Properties SamAccountName).SamAccountName }
        5    { ( Get-ADObject -LDAPFilter 'sAMAccountType=805306368' -SearchBase 'DC=domain,DC=com' -Properties SamAccountName).SamAccountName }
        6    { ( Get-QADUser -SearchRoot 'DC=domain,DC=com' -DontUseDefaultIncludedProperties -IncludedProperties SamAccountName -SizeLimit 0).SamAccountName }
        7    { ( $searcher.FindAll() ) }
        8    { (Find-LdapObject -SearchFilter:'(&(objectclass=user)(objectcategory=person))' -SearchBase:'DC=domain,DC=com' -LdapServer:'' -PageSize 1000 -PropertiesToLoad:@('sAMAccountName')) }
        9    { (Find-LdapObject -SearchFilter:'sAMAccountType=805306368' -SearchBase:'DC=domain,DC=com' -LdapServer:'' -PageSize 1000 -PropertiesToLoad:@('sAMAccountName')) }
        10   { (Find-LdapObject -SearchFilter:'(&(objectclass=user)(objectcategory=person))' -SearchBase:'DC=domain,DC=com' -LdapServer:'' -PageSize 1000) }
        11   { (Find-LdapObject -SearchFilter:'sAMAccountType=805306368' -SearchBase:'DC=domain,DC=com' -LdapServer:'' -PageSize 1000) }
        12   { (dsquery user -o samid 'DC=domain,DC=com' -limit 0) }
        13   { (dsquery * -filter '(&(objectclass=user)(objectcategory=person))' -attr samAccountName -attrsonly -limit 0) }
        14   { (dsquery * -filter 'sAMAccountType=805306368' -attr samAccountName -attrsonly -limit 0) }
        15   { ([regex]::match((.\AdFind.exe -b 'DC=domain,DC=com' -f '(&(objectclass=user)(objectcategory=person))' -c),'\d{5}').value) 2> $null }
        16   { ([regex]::match((.\AdFind.exe -b 'DC=domain,DC=com' -f 'sAMAccountType=805306368' -c),'\d{5}').value) 2> $null }
        17   { ([regex]::match((.\AdFind.exe -b 'DC=domain,DC=com' -sc adobjcnt:user -c),'\d{5}').value) 2> $null }
    }
}
 
# Check
 
for ($i = 1; $i -le 17; $i++)
{ 
    if ($i -ge 15)
    {
        $count = Get-QueryResult -Id $i
    }
    else
    {
        $count = (Get-QueryResult -Id $i | Measure-Object).Count
    }
 
    [PSCustomObject]@{
        Query = $i
        Count = $count
    }
}
 
# Measure
 
for ($i = 1; $i -le 17; $i++)
{
    New-Variable -Name "query$i" -Value $('{0:N2}' -f (Measure-Command -Expression { Get-QueryResult -ID $i }).TotalSeconds)
}
 
[PSObject]@{
    'Get-ADUser -Filter objectClass and objectCategory'                = $query1
    'Get-ADUser -LDAPFilter objectclass objectcategory'                = $query2
    'Get-ADObject -Filter objectClass and objectCategory'              = $query3
    'Get-ADObject -LDAPFilter objectclass objectcategory'              = $query4
    'Get-ADObject -LDAPFilter sAMAccountType=805306368'                = $query5
    'Quest'                                                            = $query6
    '[adsisearcher]'                                                   = $query7
    'Find-LdapObject objectClass and objectCategory PropertiesToLoad'  = $query8
    'Find-LdapObject sAMAccountType=805306368 PropertiesToLoad'        = $query9
    'Find-LdapObject objectClass and objectCategory'                   = $query10
    'Find-LdapObject sAMAccountType=805306368'                         = $query11
    'dsquery user -o samid'                                            = $query12
    'dsquery objectClass and objectCategory'                           = $query13
    'dsquery sAMAccountType=805306368'                                 = $query14
    'adfind objectClass and objectCategory'                            = $query15
    'adfind sAMAccountType=805306368'                                  = $query16
    'adfind -sc adobjcnt:user'                                         = $query17
}.GetEnumerator()  | Sort-Object -Property Value | Select-Object -Property @{
    Name       = 'Query'
    Expression = {$_.Name}
}, @{
    Name       = 'TotalSeconds'
    Expression = {[double]$_.Value}
} | Sort-Object -Property TotalSeconds | Format-Table -AutoSize