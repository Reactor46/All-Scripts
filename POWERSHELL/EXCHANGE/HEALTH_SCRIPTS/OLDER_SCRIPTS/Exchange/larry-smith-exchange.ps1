Get-Mailbox -Identity 'Larry Smith' | Search-Mailbox -SearchDumpsterOnly -TargetMailbox 'Discovery Search Mailbox' -TargetFolder 'LarrySmith-RecoverableItems'

Get-DistributionGroupMember -Identity 'NOC' |  Get-Mailbox | Select-Object * | ft -AutoSize
Get-Mailbox -Identity 'Larry Smith' | Select-Object * | Select Name,RecoverableItemsQuota,RecoverableItemsWarningQuota
Get-DistributionGroupMember -Identity 'NOC@creditone.com' | Where-Object RecipientType -eq UserMailBox | Get-Mailbox | Select-Object * | Select Name,RecoverableItemsQuota,RecoverableItemsWarningQuota

$Users = Get-DistributionGroupMember -Identity 'NOC' | Where-Object RecipientType -eq UserMailBox
ForEach($user in $Users){
Get-Mailbox | Search-Mailbox Search-Mailbox -SearchDumpsterOnly -TargetMailbox 'Discovery Search Mailbox' -TargetFolder $user'-RecoverableItems'