# This script is meant to be launched from a capacity server that produced a
# long operation which was identified by our data mining as the absolute worst
# in terms of overall impact, CPU, disk, etc.

[CmdletBinding()]
param
(
    [Parameter(Position = 0, Mandatory = $false)]
    [ValidateSet("Disk", "Latency", "Impact", "Processor")]
    [string]$Axis
)

# Check whether binaries needed for Watson are accessible

[string]$dataDirectory = Join-Path -Path $env:SystemDrive -ChildPath "Datamining\Logs\LongOperationSummary"
[string[]]$moduleDirectory = Join-Path -Path $env:ExchangeInstallPath -ChildPath "Datacenter\DataMining\Cosmos"
[string[]]$moduleFiles = "Microsoft.Exchange.Data.Common.dll","Microsoft.Exchange.Diagnostics.dll"

Write-Host "Data directory is $dataDirectory"
Write-Host "Module directory is $moduleDirectory"

if ((Test-Path -Path $moduleDirectory -PathType Container) -eq $false)
{
    throw ("Module directory $moduleDirectory is invalid")
}

foreach ($file in $moduleFiles)
{
    [string]$modulePath = Join-Path -Path $moduleDirectory -ChildPath $file

    Write-Host "Finding $file"

    if (Test-Path -Path $modulePath -PathType Leaf)
    {
        Write-Host "    Found $file"
        Import-Module $modulePath
    }
    else
    {
        throw ("Module $file not found")
    }
}

try
{
    [Data.SqlClient.SqlConnection]$connection = New-Object Data.SqlClient.SqlConnection

    Write-Host "Opening connection"
    $connection.ConnectionString = "Server=DatacenterStatusDB.main.exmgmt.local;Integrated Security=true;Connection Timeout=30"
    $connection.Open()

    if ($connection.State -ne [Data.ConnectionState]::Open)
    {
        throw "Failed to open connection to Azure data mart"
    }

    [Data.SqlClient.SqlCommand]$command = $connection.CreateCommand()

    $command.CommandText = `
        "select top 1 * " + `
        "from [edm-exo-mi].[edm-exo-mi].dbo.Store_LongOperationImpact " + `
        "order by [ID] desc, [Impact] desc"

    Write-Host "Starting reader"
    [Data.SqlClient.SqlDataReader]$reader = $command.ExecuteReader()

    try
    {
        # Currently we read only the top one operation from the Impact axis; in
        # the future this may become as many as N operations from other axes.
        # Long operations details are now included with each impact record.

        if ($reader.Read())
        {
            Write-Host "Got operation from database"

            [int]$rank = [int]$reader["Rank"]
            [int]$impact = [int]$reader["Impact"]
            [DateTime]$timeStamp = [DateTime]$reader["TimeStamp"]
            [string]$server = [string]$reader["Server"]
            [string]$databaseGuid = [string]$reader["DatabaseGuid"]
            [string]$mailboxGuid = [string]$reader["MailboxGuid"]
            [string]$clientType = [string]$reader["ClientType"]
            [string]$operationSource = [string]$reader["OperationSource"]
            [string]$operationType = [string]$reader["OperationType"]
            [string]$operationName = [string]$reader["OperationName"]
            [string]$operationDetail = [string]$reader["OperationDetail"]
            [string]$identifier = [string]$reader["Identifier"]
            [string]$correlation = [string]$reader["CorrelationId"]
            [string]$buildNumber = [string]$reader["BuildNumber"]
            [string]$clientProtocol = [string]$reader["ClientProtocol"]
            [string]$clientComponent = [string]$reader["ClientComponent"]
            [string]$clientAction = [string]$reader["ClientAction"]
            [int]$numberOperations = [int]$reader["NumberOperations"]
            [int]$mailboxesAffected = [int]$reader["MailboxesAffected"]
            [long]$chunkElapsedTime = [long]$reader["ChunkElapsedTime"]
            [long]$interactionTime = [long]$reader["InteractionTime"]
            [long]$databaseTime = [long]$reader["DatabaseTime"]
            [long]$directoryTime = [long]$reader["DirectoryTime"]
            [long]$LocksTime = [long]$reader["LocksTime"]
            [long]$plansExecutionTime = [long]$reader["PlansExecutionTime"]
            [long]$pagesPreread = [long]$reader["PagesPreread"]
            [long]$pagesRead = [long]$reader["PagesRead"]
            [long]$pagesDirtied = [long]$reader["PagesDirtied"]
            [long]$cpuKernelTime = [long]$reader["CpuKernelTime"]
            [long]$cpuUserTime = [long]$reader["CpuUserTime"]
            [bool]$isResourceIntense = if ([int]$reader["IsResourceIntense"] -eq 0) { $false } else { $true }
            [int]$mailboxesRank = [int]$reader["MailboxesRank"]
            [int]$latencyRank = [int]$reader["LatencyRank"]
            [int]$diskRank = [int]$reader["DiskRank"]
            [int]$processorRank = [int]$reader["ProcessorRank"]
            [int]$volumeRank = [int]$reader["VolumeRank"]
            [int]$mailboxesScore = [int]$reader["MailboxesScore"]
            [int]$latencyScore = [int]$reader["LatencyScore"]
            [int]$diskScore = [int]$reader["DiskScore"]
            [int]$processorScore = [int]$reader["ProcessorScore"]
            [int]$volumeScore = [int]$reader["VolumeScore"]
            [string]$details = [string]$reader["Details"]

            # Arrange more artificial constructs to appease Watson sensibilities
            [Text.StringBuilder]$stack = New-Object Text.StringBuilder -ArgumentList 2048

            [string]$bin = $identifier.GetHashCode().ToString();
            [string]$exceptionType = if ($isResourceIntense) { "Microsoft.Exchange.Server.Storage.StoreCommonServices.ResourceIntensiveException" } else { "Microsoft.Exchange.Server.Storage.StoreCommonServices.LongOperationException" }

            [void]$stack.Append($exceptionType);
            [void]$stack.AppendFormat(": Batch Rank: {0}, Impact: {1}", $rank, $impact);

            if ($operationDetail.Length -gt 0)
            {
                [void]$stack.AppendLine();
                [void]$stack.AppendFormat("   at Microsoft.Exchange.Server.Storage.StoreCommonServices.{0}()", $operationDetail);
            }

            if ($operationName.Length -gt 0)
            {
                [void]$stack.AppendLine();
                [void]$stack.AppendFormat("   at Microsoft.Exchange.Server.Storage.StoreCommonServices.{0}()", $operationName);
            }

            if ($operationType.Length -gt 0)
            {
                [void]$stack.AppendLine();
                [void]$stack.AppendFormat("   at Microsoft.Exchange.Server.Storage.StoreCommonServices.{0}()", $operationType);
            }

            if ($operationSource.Length -gt 0)
            {
                [void]$stack.AppendLine();
                [void]$stack.AppendFormat("   at Microsoft.Exchange.Server.Storage.StoreCommonServices.{0}()", $operationSource);
            }

            if ($clientType.Length -gt 0)
            {
                [void]$stack.AppendLine();
                [void]$stack.AppendFormat("   at Microsoft.Exchange.Server.Storage.StoreCommonServices.{0}()", $clientType);
            }

            [string]$callStack = $stack.ToString();

        }
        else
        {
            throw "Reader returned no long operation"
        }

        # This script should be have been launched by CA in the context of the same
        # capacity server that generated the long operation, so if the server associated
        # with the operation doesn't match this machine, stop now.

        if ([string]::IsNullOrEmpty($server))
        {
            throw "Aborting for lack of long operation"
        }

        if ($env:ComputerName.ToLower() -ne $server)
        {
            throw ("Script is running on " + ($env:ComputerName) + ", but long operation was from $server, aborting")
        }

        Write-Host "Server" ($env:ComputerName) "matched"
    }
    finally
    {
        if ($reader -ne $null)
        {
            $reader.Close()
        }
    }
}
finally
{
    if ($connection -ne $null)
    {
        $connection.Close()
    }
}

if ([string]::IsNullOrEmpty($details))
{
    [Text.StringBuilder]$builder = New-Object Text.StringBuilder -ArgumentList 40960

    [void]$builder.AppendLine("Impact Information:");
    [void]$builder.AppendFormat("    Number of Operations: {0}", $numberOperations);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Number of Mailboxes affected: {0}", $mailboxesAffected);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Mailboxes Rank: {0}", $mailboxesRank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Latency Rank: {0}", $latencyRank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Disk Rank: {0}", $diskRank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    CPU Rank: {0}", $processorRank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Volume Rank: {0}", $volumeRank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Mailboxes Score: {0}", $mailboxesScore);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Latency Score: {0}", $latencyScore);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Disk Score: {0}", $diskScore);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    CPU Score: {0}", $processorScore);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Volume Score: {0}", $volumeScore);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Overall Rank: {0}", $rank);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Total Impact: {0}", $impact);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("Common Info:");
    [void]$builder.AppendFormat("    Time Stamp: {0}", $timeStamp.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss'.'fffffff"));
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Server: {0}", $server);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Database Guid: {0}", $databaseGuid);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Mailbox Guid: {0}", $mailboxGuid);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Correlation Id: {0}", $correlation);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Build Number: {0}", $buildNumber);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("Client Values:");
    [void]$builder.AppendFormat("    Type: {0}", $clientType);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Protocol: {0}", $clientProtocol);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Component: {0}", $clientComponent);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Action: {0}", $clientAction);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("Operation Values:");
    [void]$builder.AppendFormat("    Source: {0}", $operationSource);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Type: {0}", $operationType);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Name: {0}", $operationName);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Detail: {0}", $operationDetail);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("Latency Values:");
    [void]$builder.AppendFormat("    Chunk elapsed time: {0} ms", $chunkElapsedTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Total Interaction time: {0} ms", $interactionTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("        Database time: {0} ms", $databaseTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("        Directory time: {0} ms", $directoryTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("        Lock wait time: {0} ms", $LocksTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Plan execution time: {0} ms", $plansExecutionTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("Disk I/O Values:");
    [void]$builder.AppendFormat("    Pages Preread: {0}", $pagesPreread);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Pages read: {0}", $pagesRead);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    Pages Dirtied: {0}", $pagesDirtied);
    [void]$builder.AppendLine()
    [void]$builder.AppendLine("CPU Values:");
    [void]$builder.AppendFormat("    Kernel time: {0} ms", $cpuKernelTime);
    [void]$builder.AppendLine()
    [void]$builder.AppendFormat("    User time: {0} ms", $cpuUserTime);
    [void]$builder.AppendLine()

    $details = $builder.ToString()
}
else
{
    # Replace escaped characters with functional tabs and newlines.
    $details = $details.Replace("``t", "`t").Replace("``r``n", "`r`n")
}

# Submit the artificial exception to Watson using reflection, because ExWatson
# is internal and better solutions will take weeks/months to implement/deploy.

[string]$applicationName = "Microsoft.Exchange.Store.Worker";
[string]$assemblyName = "Microsoft.Exchange.Server.Storage.StoreCommonServices";
[string]$eventType = "E12"

Write-Host "Extra details length" ($details.Length)
Write-Host "Accessing Watson"

[Reflection.Assembly]$med = [Reflection.Assembly]::GetAssembly([Microsoft.Exchange.Diagnostics.DiagnosticContext])

if ($med -ne $null)
{
    [Reflection.TypeInfo]$ew = $med.GetTypes() | where { $_.Name -eq "ExWatson" }

    if ($ew -ne $null)
    {
        [Reflection.MethodInfo]$sgwr = $ew.GetMethods() | where { $_.Name -eq "SendGenericWatsonReport" }

        if ($sgwr -ne $null)
        {
            [object[]]$parameters = New-Object object[] 10

            Write-host "Invoking send generic report"

            $parameters[0] = $eventType
            $parameters[1] = $buildNumber
            $parameters[2] = $applicationName
            $parameters[3] = $buildNumber
            $parameters[4] = $assemblyName
            $parameters[5] = $exceptionType
            $parameters[6] = $callStack
            $parameters[7] = $bin
            $parameters[8] = $identifier
            $parameters[9] = $details

            $sgwr.Invoke($null, $parameters)

            Write-Host "Watson invoked"
        }
        else
        {
            Write-Host "GetMethods failed"
        }
    }
    else
    {
        Write-Host "GetTypes failed"
    }
}
else
{
    Write-Host "LoadFile failed"
}

# SIG # Begin signature block
# MIIdxAYJKoZIhvcNAQcCoIIdtTCCHbECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNm7n63fSIlvRZv85/XMZRwsw
# Y9GgghhkMIIEwzCCA6ugAwIBAgITMwAAAJ1CaO4xHNdWvQAAAAAAnTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAy8PvNqh/8yl1
# MrZGvO1190vNqP7QS1rpo+Hg9+f2VOf/LWTsQoG0FDOwsQKDBCyrNu5TVc4+A4Zu
# vqN+7up2ZIr3FtVQsAf1K6TJSBp2JWunjswVBu47UAfP49PDIBLoDt1Y4aXzI+9N
# JbiaTwXjos6zYDKQ+v63NO6YEyfHfOpebr79gqbNghPv1hi9thBtvHMbXwkUZRmk
# ravqvD8DKiFGmBMOg/IuN8G/MPEhdImnlkYFBdnW4P0K9RFzvrABWmH3w2GEunax
# cOAmob9xbZZR8VftrfYCNkfHTFYGnaNNgRqV1rEFt866re8uexyNjOVfmR9+JBKU
# FbA0ELMPlQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGTqT/M8KvKECWB0BhVGDK52
# +fM6MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD9dHEh+Ry/aDJ1YARzBsTGeptnRBO73F/P7wF8dC7nTPNFU
# qtZhOyakS8NA/Zww74n4gvm1AWfHGjN1Ao8NiL3J6wFmmON/PEUdXA2zWFYhgeRe
# CPmATbwNN043ecHiGjWO+SeMYpvl1G4ma0NIUJau9DmTkfaMvNMK+/rNljr3MR8b
# xsSOZxx2iUiatN0ceMmIP5gS9vUpDxTZkxVsMfA5n63j18TOd4MJz+G0I62yqIvt
# Yy7GTx38SF56454wqMngiYcqM2Bjv6xu1GyHTUH7v/l21JBceIt03gmsIhlLNo8z
# Ii26X6D1sGCBEZV1YUyQC9IV2H625rVUyFZk8f4wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBMowggTGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB3jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU9Lrd2PGoAWIJaXw6GaIWPw5ROaowfgYKKwYB
# BAGCNwIBDDFwMG6gRoBEAFMAdQBiAG0AaQB0AFMAdABvAHIAZQBMAG8AbgBnAE8A
# cABlAHIAYQB0AGkAbwBuAFcAYQB0AHMAbwBuAC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBRmyRE
# 2dH/Od1iL/m/+r7viPIRsLcigFNRV8J9pNayMuK1LXPwt/vOY6X9e8RfYtX2209q
# 1NeZDaewhAVVoNWV3WPV/YEEIiG5KgaBM4wZT6u2OXBVPebL3b0bmdJj3T2e8mc7
# KEOoGdxWEukBIPNNb2J48AFaM/olesdPVv8GmVkXyQXDkax7RV/MJJ+6wU9Ci4hd
# NPvCCcCy0M1xqGZe9N72+Wguf/9ht4D+oCZOIfq12w/XAVjr6XKA0cEbnTgbJWBM
# ZfwAeDq5YSx7DLT219rQnDx01g0r932U6TBElqOSoNYhgWMouNd8XGOiaDXgV/TD
# Kk5KMtTUaUSEJfcKoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJ1CaO4xHNdWvQAAAAAAnTAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTYwOTAzMTg0NDA2WjAjBgkqhkiG9w0BCQQxFgQUIFGxHzBiGk8Vq5u6+guT
# pPD/NpIwDQYJKoZIhvcNAQEFBQAEggEAQX90obzyHOHmdHgg4P9gL34DZOiDL4QV
# fWtdgqoBNhePJ9WmgkIdu4vIceR6Gy+lvjAhcbzUgLatzA18nI1L/XsemV5WuXb4
# bKG6gRTbkJWSyXuHJFENvK56RwlakglM01MhvKSNpDnJUsqxcNzTE5RH0K4z2jF4
# 204KgBJ+6JNmt0Qm6YRwpcP2vRY2HJFSzKbE7oX2uXgB6Hy2F7yUwP/RVyc2aZ75
# C5AsuCBs7o6a9Lo07onMAj3CNaNKSX0bcHlu0q3dU/N3wu6rJsqKCkX7L0vEUTPA
# ddslZBFS7vKmfblpmXWFjuyjpyPnJbrTAVWFJYmI8AcVN4hb+94EhQ==
# SIG # End signature block
