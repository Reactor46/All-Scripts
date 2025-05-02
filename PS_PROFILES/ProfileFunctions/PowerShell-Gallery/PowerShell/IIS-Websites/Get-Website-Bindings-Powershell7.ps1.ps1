<#
CopyLeft SVET :)
.Synopsis
   Gets Website Bindings Powershell 7
.DESCRIPTION
   Long description
.EXAMPLE
   PS C:\Users\Administrator\Desktop> & '.\get web binding.ps1'
Input the website  Name: WEBSITE


Next actions will be performed on IIS websites that contain WEBSITE in their name !!!


Total bindings are


Count    : 1
Average  :
Sum      :
Maximum  :
Minimum  :
Property :

protocol           : http
bindingInformation : 192.168.2.171:80:WEBSITEURL
sslFlags           : 0

SSL Bindings are
Count    : 0
Average  :
Sum      :
Maximum  :
Minimum  :
Property :

NO SSL bindings are
protocol           : http
bindingInformation : 192.168.2.171:80:WEBSITEURL
sslFlags           : 0

Count    : 1
Average  :
Sum      :
Maximum  :
Minimum  :
Property :
.EXAMPLE
   Another example of how to use this cmdlet
#>
#Needed for PowerShell 7
Get-PSRepository PSGallery
#Needed for PowerShell 7
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#Needed for PowerShell 7
Find-Module -Name PowerShellGet | Install-Module
#Needed for PowerShell 7
Install-Module -Name PowerShellGet
#Needed for PowerShell 5
Install-Module -Name WebAdministration
#Needed for PowerShell 5
Import-Module WebAdministration
#Find-Module -Name *IIS*
#Find-Module -Name *WebAdministration*

#Needed for PowerShell 7
Install-Module -Name xWebAdministration
#Needed for PowerShell 7
Import-Module xWebAdministration
#Needed for PowerShell 7
Install-Module -Name WebAdministrationDsc
#Needed for PowerShell 7
Import-Module WebAdministrationDsc
#Needed for PowerShell 5 and 7
Install-Module -Name IISAdministration
#Needed for PowerShell 5 and 7
Import-Module IISAdministration

$site = Read-Host -Prompt 'Input the website  Name'
Write-Host "`n$line"
Write-Host "Next actions will be performed on IIS websites that contain $site in their name !!!" -ForegroundColor Yellow
Write-Host "`n$line"

Write-Host "Total bindings are"
Get-WebBinding -Name $site | measure
Get-WebBinding -Name $site

Write-Host "SSL Bindings are" -ForegroundColor Green
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 1
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 1 | measure

Write-host "NO SSL bindings are" -ForegroundColor Red
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 0
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 0 | measure


Start-Sleep -Seconds 5


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYK3V62lkaUwiIKDx9pr0U4rB
# 53igggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUvJKPDJJbQIz+
# Rslrh6LZOJN1F1AwDQYJKoZIhvcNAQEBBQAEggEAe2ExvWJU+4os5lmC5U1fMu3n
# 5lBoVQN9f39ZWd7mPzkz1brmVPNfdV2/9sK36FKF2Y4BzH5s71wuykWK9YniKdO8
# MtBADXn4/P9OSoQ7cOSDg4cBDVcGDyvGHVq/pWghbJnK7erVUuurP28KvbN5MosJ
# VFgOw1AXAryx7cYCfRiN1gD9mRasb1DB0k12zBVUn+mxU93Bv4OO7DRbGIMhTbMn
# h+Fd0qx7OaW70KBccKITHlfD5cc3yzavTNQ8L1jyAKp2qJVwxhCRG/vxUgqNEflu
# +scagV2YN7XyK2XSeg7V/6qu9X+0BIsYYF8MbZL/6AsqFMqsKjAfSIRB0dcWQg==
# SIG # End signature block
