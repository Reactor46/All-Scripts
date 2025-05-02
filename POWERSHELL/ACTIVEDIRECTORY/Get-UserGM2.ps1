$groups = Get-ADGroup -filter * -SearchBase "OU=Las_Vegas,DC=contoso,DC=com"
$output = ForEach ($g in $groups) 
 {
 $results = Get-ADGroupMember -Identity $g.name -Recursive | Get-ADUser -Properties displayname, objectclass, name 

 ForEach ($r in $results){
 New-Object PSObject -Property @{
        GroupName = $g.Name
        Username = $r.name
        ObjectClass = $r.objectclass
        DisplayName = $r.displayname
     }
    }
 } 
 $output | Export-Csv -Path c:\LazyWinAdmin\AD-groups-users.csv -NoTypeInformation﻿