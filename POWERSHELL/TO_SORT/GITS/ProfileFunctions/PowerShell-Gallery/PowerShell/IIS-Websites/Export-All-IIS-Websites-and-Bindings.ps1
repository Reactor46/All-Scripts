<#
CopyLeft SVET :)
.Synopsis
   Export All IIS Websites and Bindings
.DESCRIPTION
   Export All IIS Websites and Bindings
.EXAMPLE
    PS C:\Users\Administrator> Get-Website | Select-Object -ExpandProperty Bindings | ft

    Collection
    ----------
    {http 192.168.2.171:80:site1, https 192.168.2.171:443:site1
    {http 192.168.2.171:80:site1, https 192.168.2.171:443:site2
    {http 192.168.2.171:80:site3, https 192.168.2.171:site3

.EXAMPLE
    PS C:\Users\Administrator> $ws = Get-Website
    PS C:\Users\Administrator> $ws.PhysicalPath
    %SystemDrive%\inetpub\wwwroot
    D:\Websites\site1
    D:\Websites\site2
    D:\Websites\site3

.EXAMPLE
    PS C:\Users\Administrator> $ws.PhysicalPath | ConvertTo-Html | Out-File C:\Users\$env:UserName\Desktop\path.html
    PS C:\Users\Administrator> $ws.Bindings.Collection

    protocol bindingInformation			sslFlags
    -------- ------------------			--------
    http     *:80:                      0
    http     89.215.227.163:80:         0
    http     192.168.2.171:80:site1     0
    https    192.168.2.171:443:site2    1
    http     192.168.2.171:80:site3     0

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

#Method 1
#%windir%\system32\inetsrv\appcmd list site > c:\sites.xls

#Method 2
#Get-WebBinding | ConvertTo-Csv | Out-File "C:\bindings.csv" -Force -Verbose

#Method 3
#Get-Website | ConvertTo-Csv | Out-File "C:\sites.csv" -Force -Verbose

Get-Website | Select-Object -ExpandProperty Bindings | ft
$ws = Get-Website
$ws.PhysicalPath
$ws.PhysicalPath | ConvertTo-Html | Out-File C:\Users\$env:UserName\Desktop\path.html
$ws.Bindings.Collection
$ws.Bindings.Collection | ConvertTo-Html | Out-File C:\Users\$env:UserName\Desktop\bindings.html
$ws.Bindings.Collection.bindingInformation | clip


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrEOQk+4Vts3hcg34WfzbNcpv
# AQKgggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQULyDwETyW+cWb
# BlskhdIR28VoLvEwDQYJKoZIhvcNAQEBBQAEggEATZiHS1GnhmXZr/ptLu21zOAz
# nq7nOdl2we/i3uB5LKM57UUQcGcKnJmREOOH+PX2A07tCBAjDk8dVWOmPN2xa/M2
# iLPIwQWOghsuxk3uKXsC9NwwL5GmDQRv7kg8F/gWinIXN6LyAjEpLlcqmlrCiqj/
# OiOlIm/bnB3lBnE3INlAr87sAQMPlt+y9dO6DKtJN/tf2HnSu4kVT84i3J8e9bq3
# jUNRqqATmXkoLE6mI/d1AYbbLfTclG4LMoKkaPjTS2rq1XUNHGODnpVTJUflU1MX
# 28uQt9sFLXOZ/C8tMEJ4VUgePRevbLxWlNY2diT51YdBfMCFjvJcmer8LwHANg==
# SIG # End signature block
