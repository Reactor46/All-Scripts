## Exchange MailBox Report Wrapper
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
CD C:\Scripts\ReportingScripts\
.\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html -SendMail $true -MailFrom USONSVREX01@uson.local -MailTo "itsupport@usonv.com","Wailani.Aquino@optum.com" -MailServer mail.optummso.com
    