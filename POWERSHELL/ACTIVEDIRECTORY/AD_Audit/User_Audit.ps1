$users = Get-Content -Path .\High_Priv_Users.txt
    ForEach ($user in $users){.\Get-UserGM.ps1 -User $user | Select-Object Group |
     
     Out-File -Append -FilePath .\Users\$user.GroupMembership.txt
     }
