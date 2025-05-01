<#
CopyLeft SVET :)
.Synopsis
    Perform Hyper-V Planned Replication Failover and Failback via PowerShell
.DESCRIPTION
    Perform Hyper-V Planned Replication Failover and Failback via PowerShell
.EXAMPLE

#>
Get-VM

$vmname = Read-Host -Prompt "Choose a VM for Failover"

$status = Get-VM -ComputerName Server2012 -Name $vmname | Select -ExpandProperty Status
$state = Get-VM -ComputerName Server2012 -Name $vmname | Select -ExpandProperty State

Write-Host Status is $status and State is $state

if ( $status -like "Operating Normally" -and $status -like "Running")

{
    Write-Host Stopping $vmname now... -ForegroundColor Yellow
    Stop-VM -ComputerName Server2012 -Name $vmname
    Write-Host $vmname has been shutted down -ForegroundColor Green
    Start-Sleep -Seconds 5
}

else

{
    Write-Host Please turn it off manually!!!
    Start-Sleep -Seconds 5
}

#Fail Over Steps
Get-VMReplication 

Write-Host Preparing planned failover of the primary VM -ForegroundColor Yellow
Start-VMFailover -Prepare -VMName $vmname -ComputerName Server2012
Write-Host Preparing Completed -ForegroundColor Green

Write-Host Failing over the Replica virtual machine -ForegroundColor Yellow
Start-VMFailover -VMName $vmname -ComputerName DC
Write-Host Failing over Completed -ForegroundColor Green

Get-VMReplication

Write-Host Switching the Replica virtual machine to a primary virtual machine -ForegroundColor Yellow
Set-VMReplication -Reverse -VMName $vmname -ComputerName DC
Write-Host Switching Completed -ForegroundColor Green

Write-Host Starting the virtual machine -ForegroundColor Yellow
Start-VM -VMName $vmname -ComputerName DC
Write-Host $vmname is up and running -ForegroundColor Green

Get-VMReplication 

Start-Sleep -Seconds 5

#Fail Back Steps
Get-VMReplication 

Write-Host Preparing planned failover of the replica VM -ForegroundColor Yellow
Start-VMFailover -Prepare -VMName $vmname -ComputerName DC
Write-Host Preparing Completed -ForegroundColor Green

Write-Host Failing over the Primary virtual machine -ForegroundColor Yellow
Start-VMFailover -VMName $vmname -ComputerName Server2012
Write-Host Failing over Completed -ForegroundColor Green

Get-VMReplication 

Write-Host Switching the primary virtual machine to a replica virtual machine -ForegroundColor Yellow
Set-VMReplication -Reverse -VMName $vmname -ComputerName Server2012
Write-Host Switching Completed -ForegroundColor Green

Write-Host Starting the virtual machine -ForegroundColor Yellow
Start-VM -VMName $vmname -ComputerName Server2012
Write-Host $vmname is up and running -ForegroundColor Green

Get-VMReplication 

Start-Sleep -Seconds 5


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+1pCl1XAbrvTMqqeIiE8rpC5
# whWgggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUZzudde5E1skh
# VlxDxMLQKGCpKb0wDQYJKoZIhvcNAQEBBQAEggEAeQWLP/GqnuIl02sPKbD/Sprm
# TjF2AkvKVVdFFjS4X/d6+z6vcieomJqmQ0UZlmGP/7Ond4LsXGLym3wrw00O7SBv
# XnEQsAR2wbtoPl8jYNhhqvamq0y9YyCt/tCihdRxugnhW7z/hwrrq69ig+s8O48h
# PQ9bljPev1ukSXPbDBAMkwrCQ8C0drYwxpapJ19iG17JBmr6XECpjkicjVmYeDaG
# Ia1VyqJEgjFwoTxr6nl8rP6PEwo0R/b6kJ0GJ+DCnRFMy/OIo0LS82Qot5tR8Oy6
# tUDBHU6C6+aNS+qTdIGWZwAvwPrC6469TFLnuVs/MqD7JJSkyel7yUv2UxsJ5g==
# SIG # End signature block
