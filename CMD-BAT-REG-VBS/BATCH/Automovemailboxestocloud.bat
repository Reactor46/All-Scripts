powershell .\Automovemailboxestocloud.ps1 -RemoteHostName "hybridFQDN" -TargetDeliveryDomain "Company.mail.onmicrosoft.com" -Resourcescaretedpast "-3" -OnpremuserId "domain\serviceaccount" -OnlineuserId "serviceaccount@labtest.com" -smtpserver smtpserver -from DonotReply@labtest.com -erroremail "Reports@labtest.com" -reportemail "teamaddress@labtest.com"