<#
CopyLeft SVET :)
.Synopsis
   Removes a Web App Pool and IIS a Wesbsite for Powershell 5
.DESCRIPTION
   Removes an Application Pool and Website.
.EXAMPLE
   Another example of how to use this cmdlet

#>
#Needed for PowerShell 5
Install-Module -Name WebAdministration
#Needed for PowerShell 5
Import-Module WebAdministration
#Find-Module -Name *IIS*
#Find-Module -Name *WebAdministration*
#Needed for PowerShell 5
Install-Module -Name IISAdministration
#Needed for PowerShell 5
Import-Module IISAdministration
$wap = Read-Host -Prompt 'Input the Web App Pool Name'
Write-Host "`n$line"
Write-Host "Next actions will be performed on App Pools and IIS websites that contain $wap in their name !!!" -ForegroundColor Yellow
Write-Host "`n$line"

Get-Website -Name $wap
Get-WebAppPoolState -Name $wap

$yesno = Read-Host -Prompt 'Do you wish to proceed with stopping the website? y/n:'

if ($yesno -like 'y'){

Remove-WebAppPool -Name $wap -verbose
#Start-WebAppPool -Name *$wap* -verbose
Write-Host "`n$line"
Write-Host "App Pool was sucessfully removed!" -ForegroundColor Green
Write-Host "`n$line"

Get-Website -Name $wap
Remove-Website $wap -Verbose
Write-Host "`n$line"
Write-Host "IIS Website was sucessfully removed !!!" -ForegroundColor RED
Write-Host "`n$line"

Write-Host "The Window will close in 5 seconds"
Start-Sleep -Seconds 5

}

else {
    Write-Host Negative User Choise
}


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUC4fPI9ROgjFPy2w/AVLdD7pc
# xlCgggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
# AQsFADAZMRcwFQYDVQQDDA5HbyBMaXZlIFVLIEx0ZDAeFw0yMjA3MTYxMTU3MzZa
# Fw0yNzA3MTUyMTAwMDBaMBkxFzAVBgNVBAMMDkdvIExpdmUgVUsgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jBbQhei1OZxT2X07woTaSXAAIhc
# nY1a8OwQdzKLpWKCskGbDmGBH0bdqLq1m+y8jWKnf1tEnzaFVxUKrcwq57hy8og2
# xK9Ciq2+X43ZpA4acLUnvK0auUNig2vwWXlnPr4vh4g6+E6nMvvyhPA5ND+Uvhm2
# wSi7Z+qlmYQqm7wI7IEoaGy/8zyPB1vnWOL2UUYL8oTSklGdA6/wPGcV0S41vfGR
# 1YClE/AaAMB2+GZgkGgw9gQvKwsNRFIJSKu3OHICfrPRJgObdMnKr6F2jFl2LP+E
# Rf/Qy7BkTMwLIuxm2Tik+yPh3ZLGWGr2TtQoipNcdIg/7jmGuepTssD5/QIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFDBp69/PtsKtexPQHwwNhwtX73JqMA0GCSqGSIb3DQEBCwUAA4IBAQBQ4+Qm
# FQd/u421ZKpspyhsYw/Hjnn27uBwJ1420lpQSApuI6BnsmlZlpNVAl5ETMtbzezo
# x4xwkQ2H4/+FWg4/Mic0adMwAgqGEeOpWYZxpUQFnzCIlyH6QmXvDLuR5FwVUFKt
# fdbUj5waOOY6UJ3MSGWEykAX3R3jF/g9AgZmfvTeK9ko72QU595uIHRRXubq3akk
# 8WGiZcvgZ6WwzbHDw3iM3Q16OKOGKC/UQgJ1eiawavJUF7DQawB1VonfnbT9caZB
# 8H88Qf8u3g2FQlStoZ1lYfIUj37PZfbaBhVVZyMS7FxmSmfLXvQLtV0pT7nSRK63
# Ki30MG53JtJQdqYmMYIBzjCCAcoCAQEwLTAZMRcwFQYDVQQDDA5HbyBMaXZlIFVL
# IEx0ZAIQJLzUFQ0UMqZIQD0LynYJzzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUKMjaNcxnkyHK
# ufogu3/wRWT9Kf4wDQYJKoZIhvcNAQEBBQAEggEAKCQcpnIUrLFCyY6eCdeCi/KT
# O6ZlNz23qwW1R1aYHkLorfjWrmB9v2Y+1rKXZwIX3GerR6iNMr6aFEpZgKmK5026
# Chm+mlNbHjbITYhAu4XTvBI5Z/0Dd+YdXAIXHDyvCUNhGB3pEIoos4ulnmEySJe5
# LPvxGjyZmQdEoMbnr7bXRFnSoS3cmWu8jepk1XPTWyZpAZOZ3WbVkQqsk0uJq6vK
# WkPs0pKPSrI7+399Khdsh4j8i+xqJMyQHpp2+bkoHcnEL6lydzTP6+j1hnYxxRml
# oOW689otUdZrg6z/IiB1jMa/MfhrrQGNXZDthvle+I0HFFupox01sFNrcgN5OQ==
# SIG # End signature block
