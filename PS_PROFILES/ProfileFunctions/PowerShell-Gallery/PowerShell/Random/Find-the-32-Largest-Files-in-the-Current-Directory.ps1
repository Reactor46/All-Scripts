<#
CopyLeft SVET :)
.Synopsis
   Find the 32 Largest Files in the Current Directory
.DESCRIPTION
   Find the 32 Largest Files in the Current Directory - it helps you to calculate the staging quota for DFS Replication
.EXAMPLE

#>
Get-ChildItem .\ -Recurse | Sort-Object Length -Descending | Select-Object -First 32 | Format-Table Name,
    @{n='Length (GB)';e={$_.Length / 1gb -as [int]}},
    @{n='Length (MB)';e={$_.Length / 1mb -as [int]}},
    @{n='Length (KB)';e={$_.Length / 1kb -as [int]}}


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUnoPq/6giuoFId/bfuOGtZ7M9
# LmmgggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUk8XT5yhk8G94
# LEL7Ro02VBgHaMcwDQYJKoZIhvcNAQEBBQAEggEATtUVHrctdau7JFYEwx11k79b
# eNJgA0rjCPxo8KQEEzPSYYlT41xljftHDf6Bn3XR2EzDmcCIO/hud8oZX0HvuatG
# c7xFbclLcj9TlVd6zcXT32IuFa/t0mKlP+B1BBeJkfARUhY33ZNuK6eDXK3VDVzv
# ob+CVbvbEDHrZDyv6DyHNEHRQrxA/O5JGQaDsGmN+zdbX2x7xOAlHaU6WXZ0Jbyr
# aGUcOSuBzilcwLk5OL7nQPHQwdHPeeEtoSiDiR95mqxgyV+e1dGuzyyJssRDoPnN
# dzJXcqgnOdkTaBZp9B4IobRLb+pf/oS9AN5d0yBqIv5YNX7bFd8Bop21tBVPZw==
# SIG # End signature block
