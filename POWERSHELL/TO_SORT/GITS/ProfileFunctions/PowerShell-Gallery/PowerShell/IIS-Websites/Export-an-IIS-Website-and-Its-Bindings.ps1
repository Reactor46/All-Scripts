<#
CopyLeft SVET :)
.Synopsis
   Export an IIS Website and Its Bindings
.DESCRIPTION
   Export an IIS Website and Its Bindings
.EXAMPLE
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

Write-Host "Total bindings are "
Get-WebBinding -Name $site | measure
Get-WebBinding -Name $site

Write-Host "SSL Bindings are"
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 1
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 1 | measure

Write-host "NO SSL bindings are"
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 0
Get-WebBinding -name $site | Where-Object -Property sslFlags -eq 0 | measure


Start-Sleep -Seconds 5


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhfaG2B/sQ9aWtujtXvUn13bH
# YOCgggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUDt3Qn1/yObs7
# dmY4MOVkuvx/ZxowDQYJKoZIhvcNAQEBBQAEggEAXUTKULt/W4rxf0uK2G3UDNUD
# nDSsmSdkWkDVin43sqP+Lu/QoaSiqVTAHCeOFUDoUStVTYtLtl8qQJqony1Df1fh
# 73QYzwAHCSiNeJVm+5WqGmxbEnjWpoTeisbRebBoVwf0HeeVJNIFd5QPqf6OOQ0W
# vyWe1NZFFe2j33h3MeZU2oJGtMHPH9JVhC2dTzI5Gu5o0rNMDhVIa4jwX7uqXsRz
# tXkrnzK7oPgI1N+hJ0/LsbpWxLXqa6qLiPL5Mxe3EAhKMGG8tXYvkA7utjyjRCS7
# pTeREgkP6bpQmGjfrBcJhecru2bfhE3CMjP9oFFeATh7Q67atr6ZWfW+uR+OuA==
# SIG # End signature block
