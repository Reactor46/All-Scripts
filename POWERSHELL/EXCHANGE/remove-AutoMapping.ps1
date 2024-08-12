$FixAutoMapping = Get-MailboxPermission sharedmailbox |where {$_AccessRights -eq "FullAccess" -and $_IsInherited -eq $false}

$FixAutoMapping | Remove-MailboxPermission

$FixAutoMapping | ForEach {Add-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights:FullAccess -AutoMapping $false} 