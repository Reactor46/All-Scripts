# api: multitool
# version: 0.1
# title: Mailbox stats
# description: Get-MailboxFolderStatistics
# type: inline
# category: exchange
# depends: funcs_base
# param: username
# hidden: 0
# key: m4|mboxstats?
# status: beta
# config: -
#
# Show folders and sizes of mailbox


Param($username = (Read-Host "Username"));

Import-ExchangeSession
Get-Mailbox -Identity $username | Get-MailboxFolderStatistics | FT -Prop FolderPath, FolderType, FolderSize | Out-String
#Get-InboxRule -Identity $username