$AdminUsers = Get-ADGroupMember -Identity "Administrators" -Recursive
$DAUsers = Get-ADGroupMember -Identity "Domain Admins" -Recursive
$EAUsers = Get-ADGroupMember -Identity "Enterprise Admins" -Recursive
$BAUsers = Get-ADGroupMember -Identity "Backup Operators" -Recursive

Write-Host "There are" $AdminUsers.count "User(s)/Group(s) in the Administrators Group"
Write-Host "There are" $DAUsers.count "User(s)/Group(s) in the Domain Admins Group"
Write-Host "There are" $EAUsers.count "User(s)/Group(s) in the Enterprise Admins Group"
Write-Host "There are" $BAUsers.count "User(s)/Group(s) in the Backup Operators Group"