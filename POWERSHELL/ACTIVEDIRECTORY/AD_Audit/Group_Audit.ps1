$groups = Get-Content -Path .\All_GroupMembership.txt
    ForEach ($group in $groups){Get-ADNestedGroupMembers "$group" |
     Export-CSV .\Groups\$group.NedstedMembers.csv -NoTypeInformation     
     }
