<#
CopyLeft SVET :)
.Synopsis
   Stops a Web App Pool and IIS a Wesbsite Powershell 7
.DESCRIPTION
   Stops a Web App Pool and IIS a Wesbsite. If the Applocation Pool and/or the Website are stopped it will not work.
.EXAMPLE
PS C:\Users\Administrator\Desktop> .'.\stop.ps1'
Input the Web App Pool Name: WEBSITE


Next actions will be performed on App Pools and IIS websites that contain WEBSITE in their name !!!



Name             ID   State      Physical Path                  Bindings
----             --   -----      -------------                  --------
WEBSITE          37   Started    D:\Websites\WEBSITE            http 192.168.2.171:80:WEBSITEURL

Value : Started

Do you wish to proceed with stopping the website? y/n:: y
VERBOSE: Performing the operation "Set Item" on target "Item: ConfirmPreference Value: High".
VERBOSE: Performing the operation "Set Item" on target "Item: ConfirmPreference Value: High".


App Pools that contain WEBSITE were sucessfully stopped!


VERBOSE: Performing the operation "Set Item" on target "Item: ConfirmPreference Value: High".
VERBOSE: Performing the operation "Set Item" on target "Item: ConfirmPreference Value: High".


IIS Websites that contain WEBSITE in their name were STOPPED !!!


The Window will close in 5 seconds


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

$wap = Read-Host -Prompt 'Input the Web App Pool Name'
Write-Host "`n$line"
Write-Host "Next actions will be performed on App Pools and IIS websites that contain $wap in their name !!!" -ForegroundColor Yellow
Write-Host "`n$line"

Get-Website -Name *$wap*
Get-WebAppPoolState -Name *$wap*

$yesno = Read-Host -Prompt 'Do you wish to proceed with stopping the website? y/n:'

if ($yesno -like 'y'){

Stop-WebAppPool -Name *$wap* -verbose
#Start-WebAppPool -Name *$wap* -verbose
Write-Host "`n$line"
Write-Host "App Pools that contain $wap were sucessfully stopped!" -ForegroundColor Green
Write-Host "`n$line"


Stop-Website *$wap* -Verbose
Write-Host "`n$line"
Write-Host "IIS Websites that contain $wap in their name were STOPPED !!!" -ForegroundColor RED
Write-Host "`n$line"

Write-Host "The Window will close in 5 seconds"
Start-Sleep -Seconds 5

}

else

{
    Write-Host Negative User Input!
}


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuL+jcpyagYEL68yWqVKVqnCQ
# tM6gggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUajPoel/cV8Qi
# +f/e/ufDw3gz2+EwDQYJKoZIhvcNAQEBBQAEggEAMIZqgLvDrvtogK4BQjhzqJKM
# 2hEkSFz169VZGithHmD31zrajC/uVPvvn0lGju43Zd5wR4Jd+ryeXH3uAS/8Vu29
# dkxeIcWut8a3VNEOGon9E9F3ikVLC+BVO0/e/J6AGKfKklaPDmirtO5fZw7ARrNH
# GX3hTXBzR1Cjt4CkW1CCDXaYfek0BfpqNqeehTqPy30lYNIg9KMz3ZafW0GKbR+1
# KtD+K9MaCJ0WvBCB+i4kmTMc1XjA+is+nSWDGN+8P1JqVL31LB+uq9mmmrb1W7eA
# iVnWQh1usFBC8JS3LGsD5Co96dmQedZhWCxMBHc2swR0fOHRwjZkP3cceRMDRw==
# SIG # End signature block
