$intSent = $intRec = 0

$email = Read-Host "Enter the e-mail address of the person you wish to have the email count for:  "
$startDate = Read-Host "Enter the start date in the format mm/dd/yyyy: "
$endDate = Read-Host "Enter the end date in the format mm/dd/yyyy: "

Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $startDate -End $endDate -Sender $email -EventId RECEIVE | ? {$_.Source -eq "STOREDRIVER"} | foreach {$intSent++}

Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $startDate -End $endDate -Recipients $email -EventId DELIVER | foreach {$intRec++}


Write-Host "E-mails sent:    ", $intSent

Write-Host "E-mails received:", $intRe


