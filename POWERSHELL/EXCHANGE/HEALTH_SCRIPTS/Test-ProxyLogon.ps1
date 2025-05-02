#################################################################################
#
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages.
#
#################################################################################

# Version 21.03.07.1704

# Checks for signs of exploit from CVE-2021-26855, 26858, 26857, and 27065.
#
# Examples
#
# Check the local Exchange server only and save the report:
# .\Test-ProxyLogon.ps1 -OutPath $home\desktop\logs
#
# Check all Exchange servers and save the reports:
# Get-ExchangeServer | .\Test-ProxyLogon.ps1 -OutPath $home\desktop\logs
#
# Check all Exchange servers, but only display the results, don't save them:
# Get-ExchangeServer | .\Test-ProxyLogon.ps1


[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]]
    $ComputerName,

    [Parameter()]
    [string]
    $OutPath = "$PSScriptRoot\Test-ProxyLogonLogs",

    [Parameter()]
    [switch]
    $DisplayOnly
)

process {

    function Test-ExchangeProxyLogon {
        <#
	.SYNOPSIS
		Checks targeted exchange servers for signs of ProxyLogon vulnerability compromise.

	.DESCRIPTION
		Checks targeted exchange servers for signs of ProxyLogon vulnerability compromise.
		Will do so in parallel if more than one server is specified, so long as names aren't provided by pipeline.

		The vulnerabilities are described in CVE-2021-26855, 26858, 26857, and 27065

	.PARAMETER ComputerName
		The list of server names to scan for signs of compromise.
		Do not provide these by pipeline if you want parallel processing.

	.PARAMETER Credential
		Credentials to use for remote connections.

	.EXAMPLE
		PS C:\> Test-ExchangeProxyLogon

		Scans the current computer for signs of ProxyLogon vulnerability compromise.

	.EXAMPLE
		PS C:\> Test-ExchangeProxyLogon -ComputerName (Get-ExchangeServer).Fqdn

		Scans all exchange servers in the organization for ProxyLogon vulnerability compromises
#>
        [CmdletBinding()]
        param (
            [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
            [string[]]
            $ComputerName,

            [pscredential]
            $Credential
        )
        begin {
            #region Remoting Scriptblock
            $scriptBlock = {
                #region Functions
                function Get-ExchangeInstallPath {
                    $p = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction SilentlyContinue).MsiInstallPath
                    if ($null -eq $p) {
                        $p = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v14\Setup -ErrorAction SilentlyContinue).MsiInstallPath
                    }

                    return $p
                }

                function Get-Cve26855 {
                    [CmdletBinding()]
                    param ()

                    Write-Progress -Activity "Checking for CVE-2021-26855 in the HttpProxy logs"

                    $exchangePath = Get-ExchangeInstallPath

                    $files = (Get-ChildItem -Recurse -Path "$exchangePath\Logging\HttpProxy" -Filter '*.log').FullName
                    $count = 0
                    $allResults = @()
                    $sw = New-Object System.Diagnostics.Stopwatch
                    $sw.Start()
                    $files | ForEach-Object {
                        $count++

                        if ($sw.ElapsedMilliseconds -gt 500) {
                            Write-Progress -Activity "Checking for CVE-2021-26855 in the HttpProxy logs" -Status "$count / $($files.Count)" -PercentComplete ($count * 100 / $files.Count)
                            $sw.Restart()
                        }

                        if ((Get-ChildItem $_ -ErrorAction SilentlyContinue | Select-String "ServerInfo~").Count -gt 0) {
                            $fileResults = @(Import-Csv -Path $_ -ErrorAction SilentlyContinue | Where-Object AnchorMailbox -Like 'ServerInfo~*/*' | Select-Object -Property DateTime, RequestId, ClientIPAddress, UrlHost, UrlStem, RoutingHint, UserAgent, AnchorMailbox, HttpStatus)
                            $fileResults | ForEach-Object {
                                $allResults += $_
                            }
                        }
                    }

                    Write-Progress -Activity "Checking for CVE-2021-26855 in the HttpProxy logs" -Completed

                    return $allResults
                }

                function Get-Cve26857 {
                    [CmdletBinding()]
                    param ()

                    Get-WinEvent -FilterHashtable @{
                        LogName      = 'Application'
                        ProviderName = 'MSExchange Unified Messaging'
                        Level        = '2'
                    } -ErrorAction SilentlyContinue | Where-Object Message -Like "*System.InvalidCastException*"
                }

                function Get-Cve26858 {
                    [CmdletBinding()]
                    param ()

                    $exchangePath = Get-ExchangeInstallPath

                    Get-ChildItem -Recurse -Path "$exchangePath\Logging\OABGeneratorLog" | Select-String "Download failed and temporary file" -List | Select-Object -ExpandProperty Path
                }

                function Get-Cve27065 {
                    [CmdletBinding()]
                    param ()

                    $exchangePath = Get-ExchangeInstallPath
                    Get-ChildItem -Recurse -Path "$exchangePath\Logging\ECP\Server\*.log" -ErrorAction SilentlyContinue | Select-String "Set-.+VirtualDirectory" -List | Select-Object -ExpandProperty Path
                }

                function Get-SuspiciousFile {
                    [CmdletBinding()]
                    param ()

                    foreach ($file in Get-ChildItem -Recurse -Path "$env:WINDIR\temp\lsass.*dmp") {
                        [PSCustomObject]@{
                            ComputerName = $env:COMPUTERNAME
                            Type         = 'LsassDump'
                            Path         = $file.FullName
                            Name         = $file.Name
                        }
                    }
                    foreach ($file in Get-ChildItem -Recurse -Path "c:\root\lsass.*dmp" -ErrorAction SilentlyContinue) {
                        [PSCustomObject]@{
                            ComputerName = $env:COMPUTERNAME
                            Type         = 'LsassDump'
                            Path         = $file.FullName
                            Name         = $file.Name
                        }
                    }
                    foreach ($file in Get-ChildItem -Recurse -Path $env:ProgramData -ErrorAction SilentlyContinue | Where-Object Extension -Match ".7z$|.zip$|.rar$") {
                        [PSCustomObject]@{
                            ComputerName = $env:COMPUTERNAME
                            Type         = 'SuspiciousArchive'
                            Path         = $file.FullName
                            Name         = $file.Name
                        }
                    }
                }
                #endregion Functions

                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    Cve26855     = @(Get-Cve26855)
                    Cve26857     = @(Get-Cve26857)
                    Cve26858     = @(Get-Cve26858)
                    Cve27065     = @(Get-Cve27065)
                    Suspicious   = @(Get-SuspiciousFile)
                }
            }
            #endregion Remoting Scriptblock
            $parameters = @{
                ScriptBlock = $scriptBlock
            }
            if ($Credential) { $parameters.Credential = $Credential }
        }
        process {
            if ($null -ne $ComputerName) {
                Invoke-Command @parameters -ComputerName $ComputerName
            } else {
                Invoke-Command @parameters
            }
        }
    }

    function Write-ProxyLogonReport {
        <#
	.SYNOPSIS
		Processes output of Test-ExchangeProxyLogon for reporting on the console screen.

	.DESCRIPTION
		Processes output of Test-ExchangeProxyLogon for reporting on the console screen.

	.PARAMETER InputObject
		The reports provided by Test-ExchangeProxyLogon

	.PARAMETER OutPath
		Path to a FOLDER in which to generate output logfiles.
		This command will only write to the console screen if no path is provided.

	.EXAMPLE
		PS C:\> Test-ExchangeProxyLogon -ComputerName (Get-ExchangeServer).Fqdn | Write-ProxyLogonReport -OutPath C:\logs

		Gather data from all exchange servers in the organization and write a report to C:\logs
#>
        [CmdletBinding()]
        param (
            [parameter(ValueFromPipeline = $true)]
            $InputObject,

            [string]
            $OutPath,

            [switch]
            $DisplayOnly
        )

        begin {
            if ($OutPath -and -not $DisplayOnly) {
                New-Item $OutPath -ItemType Directory -Force | Out-Null
            }
        }

        process {
            foreach ($report in $InputObject) {
                Write-Host "ProxyLogon Status: Exchange Server $($report.ComputerName)"
                if (-not ($report.Cve26855.Count -or $report.Cve26857.Count -or $report.Cve26858.Count -or $report.Cve27065.Count -or $report.Suspicious.Count)) {
                    Write-Host "  Nothing suspicious detected" -ForegroundColor Green
                    Write-Host ""
                    continue
                }

                if ($report.Cve26855.Count -gt 0) {
                    Write-Host "  [CVE-2021-26855] Suspicious activity found in Http Proxy log!" -ForegroundColor Red
                    if (-not $DisplayOnly) {
                        $newFile = Join-Path -Path $OutPath -ChildPath "$($report.ComputerName)-Cve-2021-26855.csv"
                        $report.Cve26855 | Export-Csv -Path $newFile
                        Write-Host "  Report exported to: $newFile"
                    } else {
                        $report.Cve26855 | Format-Table DateTime, AnchorMailbox -AutoSize | Out-Host
                    }
                    Write-Host ""
                }
                if ($report.Cve26857.Count -gt 0) {
                    Write-Host "  [CVE-2021-26857] Suspicious activity found in Eventlog!" -ForegroundColor Red
                    Write-Host "  $(@($report.Cve26857).Count) events found"
                    if (-not $DisplayOnly) {
                        $newFile = Join-Path -Path $OutPath -ChildPath "$($report.ComputerName)-Cve-2021-26857.csv"
                        $report.Cve26857 | Select-Object TimeCreated, MachineName, Message | Export-Csv -Path $newFile
                        Write-Host "  Report exported to: $newFile"
                    }
                    Write-Host ""
                }
                if ($report.Cve26858.Count -gt 0) {
                    Write-Host "  [CVE-2021-26858] Suspicious activity found in OAB generator logs!" -ForegroundColor Red
                    Write-Host "  Please review the following files for 'Download failed and temporary file' entries:"
                    foreach ($entry in $report.Cve26858) {
                        Write-Host "   $entry"
                    }
                    if (-not $DisplayOnly) {
                        $newFile = Join-Path -Path $OutPath -ChildPath "$($report.ComputerName)-Cve-2021-26858.log"
                        $report.Cve26858 | Set-Content -Path $newFile
                        Write-Host "  Report exported to: $newFile"
                    }
                    Write-Host ""
                }
                if ($report.Cve27065.Count -gt 0) {
                    Write-Host "  [CVE-2021-27065] Suspicious activity found in ECP logs!" -ForegroundColor Red
                    Write-Host "  Please review the following files for 'Set-*VirtualDirectory' entries:"
                    foreach ($entry in $report.Cve27065) {
                        Write-Host "   $entry"
                    }
                    if (-not $DisplayOnly) {
                        $newFile = Join-Path -Path $OutPath -ChildPath "$($report.ComputerName)-Cve-2021-27065.log"
                        $report.Cve27065 | Set-Content -Path $newFile
                        Write-Host "  Report exported to: $newFile"
                    }
                    Write-Host ""
                }
                if ($report.Suspicious.Count -gt 0) {
                    Write-Host "  Other suspicious files found: $(@($report.Suspicious).Count)"
                    if (-not $DisplayOnly) {
                        $newFile = Join-Path -Path $OutPath -ChildPath "$($report.ComputerName)-other.csv"
                        $report.Suspicious | Export-Csv -Path $newFile
                        Write-Host "  Report exported to: $newFile"
                    } else {
                        foreach ($entry in $report.Suspicious) {
                            Write-Host "   $($entry.Type) : $($entry.Path)"
                        }
                    }
                }
            }
        }
    }

    if ($DisplayOnly) {
        $ComputerName | Test-ExchangeProxyLogon | Write-ProxyLogonReport -DisplayOnly
    } else {
        $ComputerName | Test-ExchangeProxyLogon | Write-ProxyLogonReport -OutPath $OutPath
    }
}

# SIG # Begin signature block
# MIIjngYJKoZIhvcNAQcCoIIjjzCCI4sCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCKvOYQI+AiC5cK
# HhIsfRLXZPW17RtyR8qx29thj4VfSKCCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVczCCFW8CAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg/QMryWPj
# MaDsuC/3ke2zO+ckjL1nCDtIG0wxUpNkqGQwWgYKKwYBBAGCNwIBDDFMMEqgGoAY
# AEMAUwBTACAARQB4AGMAaABhAG4AZwBloSyAKmh0dHBzOi8vZ2l0aHViLmNvbS9t
# aWNyb3NvZnQvQ1NTLUV4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQB1G5n5EGto
# W8pJ0rAfV5K7G+92n+cbOzCBhdS1Gj+55MNoikSiHJ/m4rNHeo2HBY1Jf71aBI/9
# RHxqHFuJJxZhYiJYTQq7AqggIkrYLHmLEuNT3q4PxR9Z3tRL5nAS37ieVAMkgJyN
# jIO8e6iEzUXI6a/h8lDfK+RtbEfbIzS3G6rwBEKIilSjnq3cVC0gU29hMh7+AfFD
# 7Nfnx6da+ciH8E5fz5WzxAzOV5Yw6E5ina9rNY/GwExcsydYUeO5bvzP6zB4qaaD
# HSzjjMDCn28LnpwPzZHPQcZXutWx6UsavL7dicPHHBKCuB6KkSEw6Kvse9KydCRz
# eemBZXVEh+uooYIS5TCCEuEGCisGAQQBgjcDAwExghLRMIISzQYJKoZIhvcNAQcC
# oIISvjCCEroCAQMxDzANBglghkgBZQMEAgEFADCCAVEGCyqGSIb3DQEJEAEEoIIB
# QASCATwwggE4AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIOK20ty0
# TQiFp/M8XhEOEaS8bhBQ4SqHcKSatFWzjD1kAgZgPOT6zVUYEzIwMjEwMzA4MDcw
# MTQzLjM2N1owBIACAfSggdCkgc0wgcoxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjNCQkQtRTMzOC1FOUExMSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIOPDCCBPEwggPZoAMC
# AQICEzMAAAFPZC519noDWoMAAAAAAU8wDQYJKoZIhvcNAQELBQAwfDELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0
# IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjAxMTEyMTgyNjAyWhcNMjIwMjExMTgy
# NjAyWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
# A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhh
# bGVzIFRTUyBFU046M0JCRC1FMzM4LUU5QTExJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCjFHe1ZPZoKOwb5P3E7/tIHSavithfMf8sJodyJbULIHlrUnaxSeCxNyFKB3pL
# cWOdyQDyJCTRbqRqmC0bSeD1DfT1PIv6/A6HDsZ3Ng7z3QlDg/DElXlfQaSvp32d
# fT9U742O0fvJC7sATEenBaz7fhTXQilwjuHVfU5WqbSxHnTciFWpmAbJc9BPuP+7
# pYXMUpS3awGJZk9cBFfVc9C1rA5cqT4CuIEMSw4HUQsIm4EFbDTMBSPR/hpLSVgo
# I3up1TTOp76o9gGtL+nQcVfVTNE2ffszpHxECA/Fs7XrwcbEFe002RHva0WBPbik
# ZaZeHQEHDi2EZ9MlsjytP2r9AgMBAAGjggEbMIIBFzAdBgNVHQ4EFgQUjo3u1xYG
# EH5Vk781wmTxMV/yoKAwHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUw
# VgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoGCCsGAQUF
# BwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIw
# ADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQsFAAOCAQEAPDHkqxxc
# 5DIOesrRezybooFfl4QxGmNCa6Ru2+vgL27C6wZB0R58kBniWl5AmjLovJlKvJeJ
# JPaeYhU7wVHeXwxwf+kRkQYuGFF2nRkIP8dl2ob6Ad4yb0weD9o6X5hSb6SaQCyD
# /YjoSlD5AgA4KCnsm2Auva7zBm5EIh6fie5LOqM3rnm/OAl2UOnNbffF5sg6vaFy
# 48PB1FMJUZ4gr3T2y8kEXmsE97+2ZjjJUbcE1r+vs+b1v6xZwef1dctBTUWkW1v/
# a/7WqMXtNIjrOHjCwssHhwAfulF7ms4FO1v/PYPOusHG4qbKvMRhxA4MnoYD7h1h
# yScKdxvUrN3luTCCBnEwggRZoAMCAQICCmEJgSoAAAAAAAIwDQYJKoZIhvcNAQEL
# BQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNV
# BAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4X
# DTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIxNDY1NVowfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28
# dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF++18aEssX8XD5WHCdrc+Zitb8BVTJwQx
# H0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRDDNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVH
# gc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSxz5NMksHEpl3RYRNuKMYa+YaAu99h/EbB
# Jx0kZxJyGiGKr0tkiVBisV39dx898Fd1rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL
# /W7lmsqxqPJ6Kgox8NpOBpG2iAg16HgcsOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8
# wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNV
# HQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqFbVUwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwgaAGA1UdIAEB/wSBlTCBkjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsG
# AQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vUEtJL2RvY3MvQ1BTL2Rl
# ZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAFAAbwBsAGkA
# YwB5AF8AUwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH
# 5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUxvs8F4qn++ldtGTCzwsVmyWrf9efweL3H
# qJ4l4/m87WtUVwgrUYJEEvu5U4zM9GASinbMQEBBm9xcF/9c+V4XNZgkVkt070IQ
# yK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1L3mBZdmptWvkx872ynoAb0swRCQiPM/t
# A6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWOM7tiX5rbV0Dp8c6ZZpCM/2pif93FSguR
# JuI57BlKcWOdeyFtw5yjojz6f32WapB4pm3S4Zz5Hfw42JT0xqUKloakvZ4argRC
# g7i1gJsiOCC1JeVk7Pf0v35jWSUPei45V3aicaoGig+JFrphpxHLmtgOR5qAxdDN
# p9DvfYPw4TtxCd9ddJgiCGHasFAeb73x4QDf5zEHpJM692VHeOj4qEir995yfmFr
# b3epgcunCaw5u+zGy9iCtHLNHfS4hQEegPsbiSpUObJb2sgNVZl6h3M7COaYLeqN
# 4DMuEin1wC9UJyH3yKxO2ii4sanblrKnQqLJzxlBTeCG+SqaoxFmMNO7dDJL32N7
# 9ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zP
# WAUu7w2gUDXa7wknHNWzfjUeCLraNtvTX4/edIhJEqGCAs4wggI3AgEBMIH4oYHQ
# pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
# VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1UaGFs
# ZXMgVFNTIEVTTjozQkJELUUzMzgtRTlBMTElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUA6CIM4qrSBzqcjNeHUnde
# KXgqq+iggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkq
# hkiG9w0BAQUFAAIFAOPv9KMwIhgPMjAyMTAzMDgwODU2MDNaGA8yMDIxMDMwOTA4
# NTYwM1owdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA4+/0owIBADAKAgEAAgIR9QIB
# /zAHAgEAAgIRhDAKAgUA4/FGIwIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEE
# AYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GB
# AE+qKV+VRkPJke8f0uSS61UZYH5dgWtK/19zj00Jlly95yiXG8P6CwFUhZzE8vam
# xI7DH+SUCy6LDi78Z6C5JNeOMSxvTCcwLvm5pi1U2k9mkJsaX0lcCnZFauQ9DEHq
# ldeILMuKNuv310LwYu/xVo2QS3Wl6k582BUNoHBK6OroMYIDDTCCAwkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAFPZC519noDWoMAAAAA
# AU8wDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQghE1fWPIpwxH/GlWsTjoGyIvO4hnSrA88Fn6N8Zuh
# vtUwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCAAZyYQ9oJYpMDGciFtGHJ6
# Q8+q+HltMI0QxcbBALU3AjCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAABT2QudfZ6A1qDAAAAAAFPMCIEIMkSD/E62g39wJ/NLa1/Vxq2
# 1N+Bo6pB+znN/j6IukqAMA0GCSqGSIb3DQEBCwUABIIBAFlkS7nHQAhcXzIiLMcW
# 688klK9gmAWaIoJ2FuEq/d6bbfleX+2pajFe+GsoAzp29dAY0/Nbekz0j45fImH9
# TQzlpL5Y0nof2fEApQBgxvNmeFijka0s3V4jxUhRZ+p/4hQYhdqhxSxmVijWWrAR
# nFndYH7QC6g+y6VV8HSezcFdqFT3VQxkVj/joU99qEvPRx1XJ7lQub1Uie3EjnrE
# Kn0eQNMRwzy3T1bn1fQL8nSoMPxJV08wICeUW6iT6uWYH5pd0MN0Vrph7i7lY0v9
# Dqn5nTKNEEZsvTGHd3RP/pXFPN8faPZpHcxeUXTBv0d/PYZLxENQHSVXyh28FWP0
# Mv4=
# SIG # End signature block
