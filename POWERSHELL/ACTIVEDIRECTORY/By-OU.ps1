$groups = Get-ADGroup -filter * -SearchBase "OU=Las_Vegas,DC=contoso,DC=com"
ForEach ($g in $groups) 
{
$path = "C:\LazyWinAdmin\Logs\Contoso.corp\" + $g.Name + ".csv"
Get-ADGroup -Identity $g.Name -Properties * | select name,description | Out-File $path -Append

$results = Get-ADGroupMember -Identity $g.Name -Recursive | Get-ADUser -Properties displayname, name 

ForEach ($r in $results){
New-Object PSObject -Property @{       

    DisplayName = $r.displayname | Out-File $path -Append
  }
}   
}