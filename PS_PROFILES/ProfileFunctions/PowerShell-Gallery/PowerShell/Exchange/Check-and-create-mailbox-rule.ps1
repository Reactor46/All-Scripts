#
#Check inbox rule
#
Get-InboxRule -Mailbox user@example.com | Where-Object {$_.Name -match "YOUR RULE NAME"} | Format-List Name, Enabled
#
#Ccreate inbox rule
#
New-InboxRule -Mailbox user@example.com -Name "YOUR RULE NAME" -Enabled $true -ForwardTo "user2@example.com" -BypassTransportRuleProcessing