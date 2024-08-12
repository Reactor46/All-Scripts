$Result=@()
$allMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object -Property Displayname,PrimarySMTPAddress
$totalMailboxes = $allMailboxes.Count
$i = 1 
$allMailboxes | ForEach-Object {
$mailbox = $_
Write-Progress -activity "Processing $($_.Displayname)" -status "$i out of $totalMailboxes completed"
$folderPerms = Get-MailboxFolderPermission -Identity "$($_.PrimarySMTPAddress):\Calendar"
$folderPerms | ForEach-Object {
$Result += New-Object PSObject -property @{ 
MailboxName = $mailbox.DisplayName
User = $_.User
Permissions = $_.AccessRights
}}
$i++
}
$Result | Select MailboxName, User, Permissions |
Export-CSV "D:\CalendarPermissions.CSV" -NoTypeInformation -Encoding UTF8 -Append