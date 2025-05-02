# Create instance of an object
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot

# Search with a specific search base
$search.SearchRoot = [ADSI]"LDAP://<SERVER>/OU=Employees,DC=fnbm,DC=corp"

# Search with a filter, Disabled Users
$filter = "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"
$search.Filter = $filter
# Invoke the search
$search.FindOne()
$search.FindAll()
# Invoke the search with additional parameters
$search.FindOne() | Select -ExpandProperty Properties

# Take the results and make it more Human Readable
$r = $search.findone().Properties.GetEnumerator() | foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
$r | Select Name,Title,Department,DistinguishedName,WhenChanged,LastLogonTimeStamp

# Create a function to make the Last Logon TimeStamp more Human Readable

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

# Results of using the function combined with the $search and filter
$r = $search.findone().Properties.GetEnumerator() | foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
$r | Select Name,Title,Department,DistinguishedName,WhenChanged,@{Name="LastLogon";Expression={Convert-LastLogonTimeStamp $_.lastLogonTimeStamp}}

# Combine those results to a GUI readable 
$all | Select Path | Out-GridView

# Process each user, gather properties, create an object from each user

$disabled = Foreach ($user in $all) {
$user.Properties.GetEnumerator() | 
foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
}

# First Results
$disabled | Select Name,Title,Department,DistinguishedName,WhenChanged,@{Name="LastLogon";Expression={Convert-LastLogonTimeStamp $_.lastLogonTimeStamp}}

# Results as a nice Format-Table
$disabled | sort Department | Format-Table -GroupBy Department -Property Name,
Title,@{Name="LastLogon";Expression={Convert-LastLogonTimeStamp $_.lastLogonTimeStamp}},Distinguishedname

# Another Filter
[adsisearcher]$Searcher = $filter
$searcher.SearchRoot = [ADSI]"LDAP://<SERVER>/DC=fnbm,DC=corp"


###### Find Expired Accounts and calculate tick value for AccoutnExpires attribute

$today = Get-Date
[datetime]$utc = "1/1/1601"
$ticks = ($today - $utc).ticks
$searcher.filter = "(&(objectCategory=person)(objectClass=user)(!accountexpires=0)(accountexpires<=$ticks))"

####### Find user accounts that have not loggedon in 120 days
$days = 120
$cutoff = (Get-Date).AddDays(-120)
$ticks = ($cutoff - $utc).ticks
 
$searcher.filter = "(&(objectCategory=person)(objectClass=user)(lastlogontimestamp<=$ticks))"
$all = $searcher.FindAll()


###### Work with the results generated

$inactive = Foreach ($user in $all) {
$user.Properties.GetEnumerator() | 
foreach -begin {$hash=@{}} -process {
$hash.add($_.key,$($_.Value))
} -end {[pscustomobject]$Hash}
}
 
$inactive | Select Name,Title,Department,DistinguishedName,WhenChanged,
@{Name="LastLogon";Expression={Convert-LastLogonTimeStamp $_.lastLogonTimeStamp}} |
Out-Gridview -title "Last Logon"