$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot

# Search with a specific search base
$search.SearchRoot = [ADSI]"LDAP://LASDC02/DC=fnbm,DC=corp"

# Search with a filter, Disabled Users
$filter = "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"
$search.Filter = $filter
# Invoke the search
$search.FindOne()
$search.FindAll()
# Invoke the search with additional parameters
$search.FindOne() | Select -ExpandProperty Properties
<#
Function Convert-LastLogonTimeStamp {
 
Param([int64]$LastOn=0)
 
[datetime]$utc="1/1/1601"
if ($LastOn -eq 0) {
    $utc
} else {
    [datetime]$utc="1/1/1601"
    $i=$LastOn/864000000000
 
    [datetime]$utcdate = $utc.AddDays($i)
    #adjust for time zone
    $offset = Get-WmiObject -class Win32_TimeZone
    $utcdate.AddMinutes($offset.bias)
}
} #end function




<# Take the results and make it more Human Readable
$r = $search.findone().Properties.GetEnumerator() | foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
$r | Select Name,Title,Department,DistinguishedName,WhenChanged,LastLogonTimeStamp

$r = $search.findone().Properties.GetEnumerator() | foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
$r | Select Name,Title,Department,DistinguishedName,WhenChanged,@{Name="LastLogon";Expression={Convert-LastLogonTimeStamp $_.lastLogonTimeStamp}} |
Out-GridView #>