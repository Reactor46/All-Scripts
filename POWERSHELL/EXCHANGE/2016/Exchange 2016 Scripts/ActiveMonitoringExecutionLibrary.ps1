# Copyright (c) 2015 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

<#
.SYNOPSIS 
Allows synchronous execution of HealthManager Probes
.DESCRIPTION
Loads the definitions either from the WorkItem or from the XML and allows immediate execution processing
By default, ActiveMonitoring probes are executed using the Service on scheduled iterations in a decoupled manner, making synchronous execution difficult. 
This library provides the tools to execute many probes in synchronously for immediate validation.

* Note that Loading Definitions from WorkItem can only happen once per PS Session, 
as it returns "An item with the same primary key 42 already exists in the table." error which will not let the definition be reloaded for use with this library

.EXAMPLE
Load various probes and invoke the definitions to get results for analysis - Failures will throw errors.

 . .\ActiveMonitoringExecutionLibrary.ps1
$ProbeList = @()
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromWorkItem -XmlFilename 'EdsServiceMonitoring.xml' -ProbeClassName 'Microsoft.Office.Datacenter.ActiveMonitoring.GenericServiceProbe' -ProbeName EDSServiceRunningProbe
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "FfoBackground.Xml" -ProbeName "IsBackgroundJobManagerServiceUp.Probe"
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "MessageTracing.Xml" -ProbeName "MessageTracingRepeatedlyCrashing"
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "Dal.Xml" -ProbeName "FfoRoleDALConnectivityProbe"
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "Dal.Xml" -ProbeName "WstConfigMetadataExpired_Probe"
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "Dns.Xml" -ProbeName "FfoInfraDnsLookupProbe"
$ProbeList += Load-ActiveMonitoringProbeDefinitionFromXml -XmlFilename "BGD_Antispam.Xml" -ProbeName "AntiSpamWatermarksProbe"
$ProbeResults = Invoke-ActiveMonitoringProbeFromDefinitionsInParallel -ProbeDefinitions $ProbeList 
#>
$ErrorActionPreference = 'Stop'

$FlagNonPublic = [System.Reflection.BindingFlags]::NonPublic
$FlagInstance = [System.Reflection.BindingFlags]::Instance
$global:ProbeCache = @{}

function pvt_FindPrivateField($obj, $fieldName)
{
    [System.Type]$typeToCheck = $obj.GetType()
    while (-not ($typeToCheck.FullName -ieq 'Microsoft.Office.Datacenter.WorkerTaskFramework.WorkItem' -or $typeToCheck.Name -eq 'Object'))
    {
        $typeToCheck = $typeToCheck.BaseType
    } 
    
    $field = $null
    if ($typeToCheck.FullName -ieq 'Microsoft.Office.Datacenter.WorkerTaskFramework.WorkItem')
    {
        return $typeToCheck.GetField("$fieldname", $FlagNonPublic -bor $FlagInstance)
    }
    else
    {
        write-error "Field $fieldname is not found on object $($typeToCheck.Fullname)".
    }
}

function pvt_FindPrivateMethod($obj, $methodName, [string[]]$paramList)
{
    [System.Type]$typeToCheck = $obj.GetType()
    while (-not ($typeToCheck.FullName -ieq 'Microsoft.Office.Datacenter.WorkerTaskFramework.WorkItem' -or $typeToCheck.Name -eq 'Object'))
    {
        $typeToCheck = $typeToCheck.BaseType
    } 
    
    $field = $null
    if ($typeToCheck.FullName -ieq 'Microsoft.Office.Datacenter.WorkerTaskFramework.WorkItem')
    {
        $foundMethods = $TypeToCheck.GetMethods($FlagNonPublic -bor $FlagINstance) | ?{$_.Name -like $methodName}
        foreach ($foundMethod in $foundMethods)
        {
            [string[]]$foundMethodParamList = $foundMethod.GetParameters().Name
            if ($foundMethodParamList.Count -eq $paramList.Count)
            {
                $misMatched = $false
                $paramList | ?{-not ($foundMethodParamList -contains $_)} | %{ $misMatched = $true }

                if ($misMatched)
                {
                    continue
                }
                return $foundMethod
            }
       }
    }
    else
    {
        write-error "Method $methodName is not found on object $($typeToCheck.Fullname)".
    }
}

Function Load-ActiveMonitoringProbeDefinitionFromWorkItem($XmlFilename, $ProbeClassName, $ProbeName)
{
    $installPath =(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\V1*\HealthManager | sort -Property PSPath -Descending | select -First 1).MsiInstallPath
    cd $installpath

    if (-not $Global:ProbeCache.ContainsKey($XmlFilename))
    {
        # Load the object, but cache it because HM only lets us load it once per session (it will blow up otherwise)

        # Common/Core DLLs which need to be loaded
        $AssemblyLocalComponents = [System.Reflection.Assembly]::LoadFrom("$installPath\Microsoft.Exchange.Monitoring.ActiveMonitoring.Local.Components.dll") 
        [System.Reflection.Assembly]::LoadFrom("$installPath\Microsoft.Office.Datacenter.WorkerTaskFrameworkInternalProvider.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$installPath\Microsoft.Office.Datacenter.ActiveMonitoringLocal.dll") | Out-Null

        # Settings / Values needed to load definitions from the framework
        [Microsoft.Office.Datacenter.WorkerTaskFramework.Settings]::FileStorageLocation = "$installpath\Monitoring\Config"
        [System.Reflection.Assembly]::LoadFrom("$installPath\Microsoft.Office.Datacenter.Monitoring.ActiveMonitoring.Recovery.dll") | Out-Null
        [Microsoft.Office.Datacenter.Monitoring.ActiveMonitoring.Recovery.RecoveryActionRunner]::IgnoreThrottleProperties = $true
        [System.Reflection.Assembly]::LoadFrom("$installPath\Microsoft.Office365.DataInsights.Uploader.dll") | Out-Null
 
        # Read XML Definition File
        [xml]$WorkItemXml = gc (Join-Path $installPath "Monitoring\Config\$XmlFilename")
        $def = new-object Microsoft.Office.Datacenter.ActiveMonitoring.MaintenanceDefinition
        $def.FromXml($WorkItemXml.Definition.MaintenanceDefinition)
        $AssemblyTarget = [System.Reflection.Assembly]::LoadFrom("$installPath\$($def.AssemblyPath)") 

        # Create accessors for Private fields
        $obj = $AssemblyTarget.CreateInstance($def.TypeName)
        $methodDoWork = $obj.GetType().GetMethod("DoWork", $FlagNonPublic -bor $FlagInstance)
        $methodDoInit = pvt_FindPrivateMethod -obj $obj -methodName "Initialize" -paramList @("definition")
        $fieldBroker = pvt_FindPrivateField -obj $obj -fieldName "broker"

        # Create accessors for the obejcts and functions to run probe with
        [Microsoft.Office.Datacenter.ActiveMonitoring.WorkItemFactory]::DefaultAssemblies["Microsoft.Exchange.Monitoring.ActiveMonitoring.Local.Components.dll"] = $AssemblyLocalComponents
        [Microsoft.Office.Datacenter.ActiveMonitoring.WorkItemFactory]::DefaultAssemblies[$AssemblyLocalComponents.Location] = $AssemblyLocalComponents
        $WorkItemFactory = New-Object Microsoft.Office.Datacenter.ActiveMonitoring.WorkItemFactory
        $type = ("Microsoft.Office.Datacenter.ActiveMonitoring.MaintenanceWorkBroker"+'`'+"1") -as "Type"
        $type = $type.MakeGenericType("Microsoft.Office.Datacenter.ActiveMonitoring.LocalDataAccess" -as "Type")
        $broker = [Activator]::CreateInstance($type, $WorkItemFactory)

        # Use the private accessors to set required types on the object for probe execution
        $fieldBroker.SetValue($obj, $broker) | Out-null
        $methodDoInit.Invoke($obj, $def) | out-null 

        # Errors creating definitions are stored in eventlog, we space out the calls to have their own unique second for start/end, and run the query
        $TimeStart = Get-Date
        Start-Sleep -Seconds 1 | out-null
        $methodDoWork.Invoke($obj, (new-object System.Threading.CancellationToken)) | out-null 
        Start-Sleep -Seconds 1 | out-null
        $TimeEnd = Get-Date
    
        $EventsFound = get-winevent -ea:silentlycontinue -FilterXml @"
        <QueryList>
        <Query Id="1" Path="Microsoft-Windows-AppID/Operational">
            <Select Path="Microsoft-Exchange-ActiveMonitoring/ProbeDefinition">
                                    *[System[(EventID&gt;0) and
                                    TimeCreated[@SystemTime &lt;="$($TimeEnd.ToString("s"))"] and 
                                    TimeCreated[@SystemTime &gt;="$($TimeStart.ToString("s"))"]]] 
                                </Select>
                            </Query>
        </QueryList>
"@ 

        if ($EventsFound)
        {
            $EventsFound | %{
                [xml]$EventData = $_.ToXml()
                $attributesText = $eventdata.Event.UserData.EventXml.ExtensionAttributes
                if ($attributesText -like "<Exception")
                {
		            [xml]$attributes = $attributesText
                    $attributes.Exception
                }
                $attributes.Exception
            } | Select -Unique
        }
    
        $global:ProbeCache[$XmlFilename] = $obj
    }
   
    $objWorkItem = $Global:ProbeCache[$XmlFilename]
    
    # Use the public methods to set the class and return the probe definition
    $ProbeDefDataAccess = $objWorkItem.Broker.GetProbeDefinition($ProbeClassName) | ?{$_.Name -eq $ProbeName}
    if ($ProbeDefDataAccess -ne $null)
    {
        $probeDef = ($ProbeDefDataAccess.InnerQuery | Select-Object -Unique) -as [Microsoft.Office.Datacenter.ActiveMonitoring.ProbeDefinition] 
        if ($probeDef -eq $null)
        {
            $probeDef = $ProbeDefDataAccess -as [Microsoft.Office.Datacenter.ActiveMonitoring.ProbeDefinition] 
        }
    }
    
    if ($ProbeDef -eq $null)
    {
        [int]$DefinitionCount = $objWorkItem.Broker.GetProbeDefinition($ProbeClassName).Name.Count
        Write-Error "Probes Loaded = $DefinitionCount, Specific Probe not found: XmlFilename = $XmlFilename, ProbeClassName = $ProbeClassName, ProbeName = $ProbeName `n`n$DefinitionLoadEventData"
    }
    return $probeDef
}

function Load-ActiveMonitoringProbeDefinitionFromXml($XmlFilename, $ProbeName)
{
    $installPath =(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\V1*\HealthManager | sort -Property PSPath -Descending | select -First 1).MsiInstallPath
    Add-Type -Path "$installPath\Microsoft.Office.Datacenter.ActiveMonitoringLocal.dll"
    
    [xml]$ComponentXml = gc (Join-Path $InstallPath "Monitoring\Config\$XmlFilename")

    # Load XML and Convert WellKnown Classes into their Full Typenames
    $ProbeXml = $ComponentXml.SelectSingleNode("//Probe[@Name ='$ProbeName']")
    $probeDef = new-object Microsoft.Office.Datacenter.ActiveMonitoring.ProbeDefinition
    $ProbeTypePrefix = switch -regex ($ProbeXml.TypeName)
    { 
        'FilteredGenericEventLogProbe|GenericEventLogProbe|GenericProcessCrashDetectionProbe'
            { 
                "Microsoft.Exchange.Monitoring.ActiveMonitoring.Local.Components.Common.Probes."
                break;
            }
        "GenericServiceProbe" 
            { 
                "Microsoft.Office.Datacenter.ActiveMonitoring."
                break;
            }
    }

    # If the definition is present, load it to get the service name
    if ($ComponentXml.Definition.MaintenanceDefinition -ne $null)
    {
        $maintDef = new-object Microsoft.Office.Datacenter.ActiveMonitoring.MaintenanceDefinition
        $maintDef.FromXml($ComponentXml.Definition.MaintenanceDefinition)
        $probeXml.SetAttribute("ServiceName", $null, $maintDef.ServiceName) | Out-Null
        $probeXml.SetAttribute("AssemblyPath", $null, "$installPath\$($maintDef.AssemblyPath)") | Out-Null
    }
    else
    {
        $probeXml.SetAttribute("ServiceName", $null, "") | Out-Null
        $probeXml.SetAttribute("AssemblyPath", $null, "$installPath\Microsoft.Exchange.Monitoring.ActiveMonitoring.Local.Components.dll") | Out-Null
    }
    $probeXml.SetAttribute("Enabled", $null, "True") | Out-Null
    $probeXml.SetAttribute("MaxRetryAttempts", $null, "0") | Out-Null
    $probeXml.SetAttribute("TimeoutSeconds", $null, "0") | Out-Null
    $probeXml.SetAttribute("TypeName", $null, "$ProbeTypePrefix$($probeXml.TypeName)") | Out-Null
    $probeDef.FromXml($ProbeXml)
    
    return $ProbeDef
}

function Invoke-ActiveMonitoringProbeFromDefinition($Probe, [switch]$suppressThrow)
{
    $installPath =(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\V1*\HealthManager | sort -Property PSPath -Descending | select -First 1).MsiInstallPath
    cd $installpath

    $ProbeInfo = "Probe - $($Probe.Name), $($Probe.AssemblyPath), $($Probe.TypeName)"
    Write-Verbose $ProbeInfo
    
    # Setup the Probe Result and Load the Probe object
    [System.Reflection.Assembly]::LoadFrom($Probe.AssemblyPath) | Out-Null
    $probeObj = New-Object $Probe.TypeName
    
    # Set the private members of the probe: Definition, Result and DoWork
    $probeResult = new-object Microsoft.Office.Datacenter.ActiveMonitoring.ProbeResult
    $methodDoWork = $probeObj.GetType().GetMethod("DoWork", $FlagNonPublic -bor $FlagInstance)
    $fieldDefinition = pvt_FindPrivateField -obj $probeObj -fieldName "definition"
    $fieldResult = pvt_FindPrivateField -obj $probeObj -fieldName "result"
    $fieldDefinition.SetValue($probeObj, $Probe) | Out-Null
    $fieldResult.SetValue($probeObj, $probeResult) | Out-Null

    try
    {
        $methodDoWork.Invoke($probeObj, (new-object System.Threading.CancellationToken)) | Out-Null
    }
    catch 
    {
        if ($probeResult.Error)
        {
            $Error = "$($Probe.TypeName): Probe failed: $($probeResult.Error)"
        }
        else
        {
            $error = "$($Probe.TypeName): $($_.Exception.InnerException.Message)"
            if ($probeResult.ExecutionContext)
            {
                $Error += "`nLog: $($probeResult.ExecutionContext)"
            }
        }
        
        if (-not $suppressThrow)
        {
            throw $error
        }
        return $error
    }
    
    # No return string indicates success, as probes indicate success by not failing.
    return
}

function Invoke-ActiveMonitoringProbeFromDefinitionsInParallel($ProbeDefinitions)
{
    $RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, 10)
    $RunspacePool.Open()
    $ScriptBlock = {
        Param (
            $Probe
        )
        $installPath =(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\V1*\HealthManager | sort -Property PSPath -Descending | select -First 1).MsiInstallPath
        . "$InstallPath\ActiveMonitoringExecutionLibrary.ps1"
        cd $installPath
        Invoke-ActiveMonitoringProbeFromDefinition -Probe $Probe -suppressThrow 
    }
    
    # Setup the jobs and start them in parallel
    $Jobs = @()
    $ProbeDefinitions | %{
        $Job = [powershell]::Create().AddScript($ScriptBlock).AddArgument($_)
        $Job.RunspacePool = $RunspacePool
        $Jobs += New-Object PSObject -Property @{
           Pipe = $Job
           Probe = $_
           Result = $Job.BeginInvoke()
        }
    }     

    # Wait for all jobs to complete
    Do {    Start-Sleep -Seconds 1  } 
    While ( $Jobs.Result.IsCompleted -contains $false )

    # Collect the results
    $Results = @()
    ForEach ($Job in $Jobs )
    {   
        $ProbeResults += $Job.Pipe.EndInvoke($Job.Result)
    }

    if ($ProbeResults -ne $null)
    {
        Write-Error $("$($ProbeResults.Count) Probe Errors Reported:`n" + [string]::Join("`n", $ProbeResults))
    }
}


# SIG # Begin signature block
# MIIdyQYJKoZIhvcNAQcCoIIdujCCHbYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyxUMmU5pl5bTefq0357eRQE4
# /P+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBM8wggTLAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB4zAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUuMSbCChVTTQEn/06Zuo42m19ihAwgYIGCisG
# AQQBgjcCAQwxdDByoEqASABBAGMAdABpAHYAZQBNAG8AbgBpAHQAbwByAGkAbgBn
# AEUAeABlAGMAdQB0AGkAbwBuAEwAaQBiAHIAYQByAHkALgBwAHMAMaEkgCJodHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIB
# AI9pMOWC7/HRxedARBsONzwTZ868cmBoickqrx40YoycAxVsJhc6tnOV3kO3HQaF
# edyPoeEOkDBdYG20uZTUcaX2FnoCnAXxj1ozBxnmGsntKx2lAyWRNlE4IfUGHros
# y9rNM/mWLVorPCH9sQU1PVlPChUKKoNQvxRcx17qAdT2fOewnrV7bKxpSHtkzHgh
# KDO9Z4JVmA5lto9qhCfZxQ3y4Dq738sWGUMemKCvirZEMscW2STx4Uud76xqzRCt
# L3CdggCoSZTDzf4Z5IqPPsUUi68ScBIiFdlYz3tb9ICyQV0618UXhqirFkVTrOjA
# kXOz9cZg8BpqXdg4WVlHN+uhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEw
# gY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UE
# AxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmpqbFsKD2tXCAAAAAACa
# MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3
# DQEJBTEPFw0xNjA5MDMxODQ1MDRaMCMGCSqGSIb3DQEJBDEWBBTMZFqYhHJJDU2G
# sITuCwUtksI3QTANBgkqhkiG9w0BAQUFAASCAQBmrhEfVQ88kUBCEH7sxvafMjeL
# ii01HRHxduLqEl8s5LssozoKkMY/+Ig1UgdS6+N8Pi6z59cgvkbG3ImjRKE3TJIP
# uTwOvtQADUV6PG8u2uOPlCbMsf19FQOt7sOHAzMQA7crKeot3I4q5cqSR0qEbwRt
# iP/Zou7zURvykjpDc4Bb8fhLzvu4vC+oAXjsP0iwQQB3Rtf4O3IyCxFEjbKkzHoa
# XEXukw7FF1ssI/WaUtz8DPSAejobjcf8LvtXIOMnwEsYhwzVHsSEkQjzTVW+GKt8
# yGLaO3F/30bl14rODKuRw4yS5pZ4xG1nfi7GVusyGQkOwE+DnIy08Zvw91Ot
# SIG # End signature block
