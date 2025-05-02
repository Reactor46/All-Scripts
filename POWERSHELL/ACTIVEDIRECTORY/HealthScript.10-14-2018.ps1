GoGo-PSExch
$mb1 = get-healthreport -server vvcexcmb01 | where {$_.alertvalue -ne “healthy”} | ft -auto | Out-String
$mb2 = get-healthreport -server vvcexcmb02 | where {$_.alertvalue -ne “healthy”} | ft -auto | Out-String
$wa1 = get-healthreport -server vvcexcwa01 | where {$_.alertvalue -ne “healthy”} | ft -auto | Out-String
$wa2 = get-healthreport -server vvcexcwa02 | where {$_.alertvalue -ne “healthy”} | ft -auto | Out-String
$idxmb1 = Get-MailboxDatabaseCopyStatus mb1 | where {$_.ContentIndexState -ne "Healthy"} | ft -auto name,status,contentindexstate | Out-String
$idxmb2 = Get-MailboxDatabaseCopyStatus mb2 | where {$_.ContentIndexState -ne "Healthy"} | ft -auto name,status,contentindexstate | Out-String
$idxarc = Get-MailboxDatabaseCopyStatus ExchangeArchives | where {$_.ContentIndexState -ne "Healthy"} | ft -auto name,status,contentindexstate | Out-String
$idxmb4 = Get-MailboxDatabaseCopyStatus mb4 | where {$_.ContentIndexState -ne "Healthy"} | ft -auto name,status,contentindexstate | Out-String
$idxmb5 = Get-MailboxDatabaseCopyStatus mb5 | where {$_.ContentIndexState -ne "Healthy"} | ft -auto name,status,contentindexstate | Out-String
$body = $mb1 + $mb2 + $wa1 + $wa2 + $idxmb1 + $idxmb2 + $idxarc + $idxmb4 + $idxmb5
Send-MailMessage -From "Exchange2013@vegas.com" -To "john.battista@vegas.com" -SMTPServer intmail.vegas.com -Subject "Exchange 2013 Health Report" -Body $body

Get-PSSession | Remove-PSSession
