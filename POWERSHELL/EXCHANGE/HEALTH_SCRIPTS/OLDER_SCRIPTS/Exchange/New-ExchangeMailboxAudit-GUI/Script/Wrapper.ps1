## Exchange MailBox Report Wrapper
CD C:\Scripts\Exchange\New-ExchangeMailboxAudit-GUI\Script
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
. .\New-ExchangeMailboxAudit.ps1
    New-ExchangeMailboxAuditReport -LoadConfiguration
    