# .SYNOPSIS
#   This script sets MRS settings in the MRS or Migration Service Config file (including WLM delay configuration in the registry for MRS).
#
# .DESCRIPTION
#   Purpose of the script is to allow configuring MRS or Migration Service configuration settings.
#   This script is typically executed by Ops or product team on-call and is designed to be used only on Exchange Datacenters.
#   MRS runs only on ClientAccessServer. Migration Service runs only on mailbox servers. Hence, this script should only be targeted at either of these servers.
#
# .PARAMETER Delays
#   This is the throttling setting. 2 means WLM delay throttling is enabled. 1 means WLM delay throttling is disabled. 0 means it remains untouched.
#
# .PARAMETER SettingNameValuePairs
#   This is the string containing the name value pairs for MRS or Migration service settings
#   in either MRSConfiguration or AppSettings sections.
#   The expected format is "Name1=Value1_Name2=Value2_Name3=Value3".
#   If the setting is in MRSConfiguration section, it should exist previously. If the setting is in AppSettings section 
#   (of either MRS or Migration Service files) and doesn't exist already, it will be added.
#
# .PARAMETER xPath
#   This is the string specifying the xpath of an Xml attribute to modify in the MRS or Migration service config file.
#   The xpath must be valid i.e. the specified attribute must already exist in order to be modified. The xPath is case-sensitive.
#   This parameter along with xPathValue may be used for settings that are not in MRSConfiguration or AppSettings sections of the configuration file.
#
# .PARAMETER xPathValue
#   This is the string specifying the new value for the attribute specified by the -xPath parameter.
#
# .PARAMETER ConfigToUpdate
#   This is the parameter specifying which configuration file is to be altered. Use 'MRS' to denote MSExchangeMailboxReplication.exe.config, 
#   Migration to denote Microsoft.Exchange.ServiceHost.exe.config and LoadBalance to denote MSExchangeMigrationWorkflow.exe.config.
#
# .PARAMETER RemoveNode
#   This is the parameter specifying that the node selected by the xPath expression is to be removed from the config.
#
# .PARAMETER AddNode
#   This is the parameter specifying that the value specified by xPathValue is going to get added as a child of the node selected by xPath.
#
# .PARAMETER RestartService
#   If this switch is provided then the service that's config is being changed will be restarted.  If the switch isn't provided no restart will occur.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -Delays 0
#   This causes the WLM throttling for MRS setting to remain untouched.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -Delays 1
#   This causes WLM to throttle MRS with delays.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -Delays 2
#   This causes WLM to stop throttling MRS with delays.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -SettingNameValuePairs "MaxActiveMovesPerSourceMDB=5_MaxActiveMovesPerTargetMDB=3_MaxActiveMovesPerSourceServer=50_MaxActiveMovesPerTargetServer=20_MaxTotalRequestsPerMRS=100" -ConfigToUpdate MRS
#   This modifies only the four MRS config settings provided in the SettingNameValuePairs parameter string.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -Delays 2 -SettingNameValuePairs "MaxActiveMovesPerSourceMDB=5_MaxActiveMovesPerTargetMDB=3_MaxActiveMovesPerSourceServer=50_MaxActiveMovesPerTargetServer=20_MaxTotalRequestsPerMRS=100" -MRS
#   This causes WLM to throttle MRS with delays as well as changes the MRS config settings specified.
#
# .EXAMPLE
#   .\ConfigureMRSSettings.ps1 -SettingNameValuePairs "SyncMigrationEnabledMigrationTypes=IMAP,Exchange" -ConfigToUpdate Migration
#   This modifies the SyncMigrationEnabledMigrationTypes setting provided in the Migration config file.
#
# .EXAMPLE
#   .\ConfigureMRSSettings.ps1 -SettingNameValuePairs "K1=V1:V11" -ConfigToUpdate Migration
#   This adds the new row K1 <add key = "K1" value = "V1:V11" /> in the AppSettings section of the Migration config file.
#
# .EXAMPLE
#   .\ConfigureMRSSettings.ps1 -xPath "/configuration/BandConfiguration" -ConfigToUpdate LoadBalance -xPathValue '<band MinSizeMb="100" MaxSizeMb="5000" Enabled="True" Profile="SizeBased" />' -AddNode
#   This adds the new node Band with attributes MinSizeMb=100, MaxSizeMb=5000, Enabled=True and Profile=SizeBased to the LoadBalance config file.
#
# .EXAMPLE
#   .\ConfigureMRSSettings.ps1 -xPath "/configuration/BandConfiguration/band[@MinSizeMb=250]" -ConfigToUpdate LoadBalance -RemoveNode
#   This removes the Band node with MinSizeMb=250 from the LoadBalance config file.
#
# .EXAMPLE
#   .\ConfigureMRSSettings.ps1 -xPath "/configuration/BandConfiguration/band[@MinSizeMb=250]/@MinSizeMb" -ConfigToUpdate LoadBalance -xPathValue "200"
#   This changes Band node with MinSizeMb=250 from the LoadBalance config file to have MinSizeMb=200.
#
# .EXAMPLE
#   ConfigureMrsSettings.ps1 -xPath "//configuration/system.serviceModel/bindings/customBinding/binding/httpsTransport/@maxReceivedMessageSize" -xPathValue "5050505050" -ConfigToUpdate mrs
#   This causes the maxReceivedMessageSize under the specified path to be modified to 5050505050, in the MRS config file.
#
#   Copyright (c) Microsoft Corporation. All rights reserved.
#
#   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#   OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

#
# [PrivilegeLevel(LocalSystem)]
#

[CmdletBinding(DefaultParameterSetName="basic")]
param(
    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [string]$SettingNameValuePairs,

    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [Int32]$Delays,

    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [Parameter(Mandatory=$true, ParameterSetName="xpathadd")]
    [Parameter(Mandatory=$true, ParameterSetName="xpathremove")]
    [string]$xPath,

    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [Parameter(Mandatory=$true, ParameterSetName="xpathadd")]
    [string]$xPathValue,

    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [Parameter(Mandatory=$true, ParameterSetName="xpathadd")]
    [Parameter(Mandatory=$true, ParameterSetName="xpathremove")]
    [ValidateSet("mrs", "migration", "loadbalance")]
    [string]$ConfigToUpdate,

    [Parameter(Mandatory=$true, ParameterSetName="xpathremove")]
    [Switch]$RemoveNode,

    [Parameter(Mandatory=$true, ParameterSetName="xpathadd")]
    [Switch]$AddNode,

    [Parameter(Mandatory=$false, ParameterSetName="basic")]
    [Parameter(Mandatory=$false, ParameterSetName="xpathadd")]
    [Parameter(Mandatory=$false, ParameterSetName="xpathremove")]
    [switch]$RestartService
)

function FileIsMRS($inputFile)
{
    return ($inputFile.ToLower().Equals("mrs"))
}

function FileIsMigration($inputFile)
{
    return ($inputFile.ToLower().Equals("migration"))
}

function HaveSettings()
{
    if (![string]::IsNullOrEmpty($xPath))
    {
        # this takes care of both AdddNode and regular update configuration cases
        if (![string]::IsNullOrEmpty($xPathValue))
        {
            return $true
        }

        if ($RemoveNode)
        {
            return $true
        }
    }
    
    if (![string]::IsNullOrEmpty($SettingNameValuePairs))
    {
        return $true
    }
    
    return $false
}

function VerifyUsage()
{
    [bool]$configChanged = HaveSettings
    if ($Delays -eq 0 -and !$configChanged)
    {
        throw "No change requested. Please run Get-help ConfigureMrsSettings.ps1 -examples to view usage."
    }
    
    if ($Delays -eq 0)
    {
        if ([string]::IsNullOrEmpty($ConfigToUpdate))
        {
            throw "Either Delay or ConfigToUpdate must be specified. Please run Get-help ConfigureMrsSettings.ps1 -examples to view usage."
        }
    }
}

function GetServiceToRestart()
{
    $sName = ""
    switch ($ConfigToUpdate)
    {
        "mrs" { $sName = "MSExchangeMailboxReplication" }
        "migration" { $sName = "MSExchangeServiceHost" }
        "loadbalance" { $sName = "MSExchangeMigrationWorkflow" }
        default { throw "Unknown config: $ConfigToUpdate" }
    }
    
    $service = Get-Service $sName
    if ($service -eq $null)
    {
        throw "Could not find $sName. This script must be targeted only against ClientAccessServer for MRS and Mailbox Server for Migration Service and Load Balance."
    }
    
    return $sName
}

function ConfigureWlmDelay()
{
    if ($Delays -eq 0)
    {
        return
    }

    if ($Delays -eq 2)
    {
        Write-Host "Configuring WLM to ENABLE MRS delays..."
        $valueToSet = 1
    }
    else
    {
        Write-Host "Configuring WLM to DISABLE MRS delays..."
        $valueToSet = 0
    }

    # Configure WLM to switch throttling ON/OFF for MRS.
    $keyHive = "HKLM:\system\CurrentControlSet\services\MSExchange ResourceHealth"
    $keyName = "MRS"
    $mrsValue = Get-ItemProperty -Path $keyHive -Name $keyName

    if ($mrsValue -eq $null)
    {
        throw "Could not find Registry key ${keyHive}:$keyName."
    }

    Write-Host "Current setting for ${keyHive}:$keyName is: $mrsValue"

    Set-ItemProperty -Path $keyHive -Name $keyName -Value $valueToSet
    $mrsValue = Get-ItemProperty -Path $keyHive -Name $keyName
    Write-Host "Setting was set to: $mrsValue"

    Write-Host "WLM Registry configuration successful."
    
    RestartService "MSExchangeMailboxReplication"
}

#Get the install path - must handle both V14 or V15 target servers.
function GetMSIInstallPath()
{
    $DirPath = $null
    # Check higher version first
    $DirPath=(get-itemproperty -Path "HKLM:\software\microsoft\exchangeServer\v15\setup" -ErrorAction SilentlyContinue)
    if ($DirPath -ne $null)
    { 
        Write-Host "Found E15 install path"
    }
    else
    {
        $DirPath=(get-itemproperty -Path "HKLM:\software\microsoft\exchangeServer\v14\setup" -ErrorAction SilentlyContinue)
        if ($DirPath -ne $null)
        { 
            Write-Host "Found E14 install path"
        }
        else
        {
            throw "MSIInstallPath was not found"
        }
    }

    $path = $DirPath.MSIInstallPath.TrimEnd("\")
    Write-Host "Found Dir path = $path"

    return $path
}

function Supports-Removal
{
    param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
    [System.Xml.XmlNode]$Node
    )

    process
    {
        # a null node cannot be removed
        if ($Node -eq $null) 
        {
            return $false
        }

        # top of hierarchy is not removable
        if ($node.ParentNode -eq $null)
        {
            return $false
        }

        # these are the nodes that can be removed
        switch ($node.Name)
        {
            "BandConfiguration" { return $true }
            "WeightConfiguration" { return $true }
            "LoadBalanceConfiguration" { return $true }
        }

        # anything else is removable if the parent is removable
        return Supports-Removal $node.ParentNode
    }
}

function Get-NodePath
{
    param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
    [System.Xml.XmlNode]$Node
    )

    process
    {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Attribute)
        {
            $nodeName = "@{0}" -f $node.Name
        }
        else
        {
            $nodeName = $node.Name
        }

        if ($node.ParentNode -eq $nul)
        {
            return "/$nodeName"
        }

        $parentNode = Get-NodePath $node.ParentNode

        return $("{0}/{1}" -f @($parentNode, $nodeName))
    }
}

function ConfigureSettings
{
    [bool]$continue = HaveSettings
    if (!$continue)
    {
        return
    }
   
    $NameValuePairs = new-object "System.Collections.Generic.Dictionary``2[System.String, System.String]"

    [string[]]$arrayOfValues = $SettingNameValuePairs.Split("_")

    foreach ($value in $arrayOfValues)
    {
        [string[]] $nameValue = $value.Split("=")
        if ($nameValue.Count -eq 2)
        {
            $NameValuePairs.Add($nameValue[0], $nameValue[1])
            Write-Host "Adding element to dictionary: $nameValue[0], $nameValue[1]"
        }
    }

    Write-Host "SettingNameValuePairs has $($SettingNameValuePairs.count) elements."
    Write-Host "$SettingNameValuePairs"

    $ConfigPath = ""

    [string]$ExchangeInstallDir= GetMSIInstallPath
    switch ($ConfigToUpdate)
    {
        "mrs" 
        { 
            Write-Host "Updating MRS Config settings..."
            $ConfigPath = "$ExchangeInstallDir\bin\MSExchangeMailboxReplication.exe.config"
        }

        "migration" 
        { 
            Write-Host "Updating Migration Config settings..."
            $ConfigPath = "$ExchangeInstallDir\bin\Microsoft.Exchange.Servicehost.exe.config"
        }
        
        "loadbalance" 
        { 
            Write-Host "Updating Load Balancing settings..."
            $ConfigPath = "$ExchangeInstallDir\bin\MSExchangeMigrationWorkflow.exe.config"
        }

        default { throw "Unknown config: $ConfigToUpdate" }
    }

    if (-not (Test-Path $ConfigPath -PathType Leaf)) 
    { 
        throw "Could not find the config file for $env:COMPUTERNAME. File Path tried was: $ConfigPath"  
    }
    else
    {
        [xml]$Config = get-content $ConfigPath

        [int]$settings = 0

        $svcName = GetServiceToRestart

        # SettingNameValuePair settings apply to all config modes
        if ($NameValuePairs.count -GT 0)
        {
            foreach ($parameterName in $NameValuePairs.Keys)
            {
                if ($parameterName -ne "Delays")
                {
                    $parameterValue = $($NameValuePairs[$parameterName])
                    $addedNameValueSettingToMrs = $false

                    if (FileIsMRS($ConfigToUpdate))
                    {
                        if ((get-member -inputObject $Config.Configuration.MRSConfiguration -name $parameterName -WarningAction:SilentlyContinue) -ne $null)
                        {
                            # First, look in the MRSConfiguration section (settings used by MRS)
                            $oldValue = $Config.Configuration.MRSConfiguration.$parameterName
                            $Config.Configuration.MRSConfiguration.$parameterName = "$parameterValue"
                            Write-Host "$parameterName will be changed from $oldValue to $parameterValue"
                            $addedNameValueSettingToMrs = $true
                        }
                    }

                    # Add or modify setting in AppSettings section if it wasn't already added to MRSConfiguration section above.
                    if (!$addedNameValueSettingToMrs)
                    {
                        # AppSettings section settings apply to both MRS and Migration Service.
                        if (($Config.configuration.appSettings.selectsinglenode("add[@key='" + $parameterName + "']")) -ne $null)
                        {
                            # Modify existing setting value
                            $oldValue = $Config.Configuration.appSettings.selectsinglenode("add[@key='" + $parameterName + "']").value
                            $Config.Configuration.appSettings.selectsinglenode("add[@key='" + $parameterName + "']").value = "$parameterValue"
                            Write-Host "$parameterName will be changed from $oldValue to $parameterValue"
                        }
                        else
                        {
                            # Setting doesn't exist. Add new element under appSettings.
                            $appS = $Config.configuration.appSettings
                            $oldInnerXml = $appS.InnerXml
                            $newElement = [string]::Format('<add key = "{0}" value = "{1}" />', $parameterName, $parameterValue)
                            $newInnerXml = [string]::Concat($oldInnerXml, $newElement)
                            $appS.InnerXml = $newInnerXml
                            Write-Host "Setting $parameterName not found. Will add new element under appSettings: $newElement"
                        }
                    }

                    $settings++
                }
            }
        }

        # XPath settings apply to all config modes
        if (![string]::IsNullOrEmpty($xPath) -and (![string]::IsNullOrEmpty($xPathValue) -or $RemoveNode))
        {
            Write-host "xPath = $xPath"
            Write-host "xPathValue = $xPathValue"
            
            #Retrieve the node specified by the xPath provided.
            $node = $Config.selectsinglenode("$xPath")
            if ($node -eq $null)
            {
                throw "XPath $xPath not found in $ConfigToUpdate config"
            }

            if ($PSCmdlet.ParameterSetName -eq "basic")
            {
                $oldValue = $node."#text"
                Write-Host "$xPath will be changed from $oldValue to $xPathValue"
                # Modify the value of the Xpath-selected node.
                $node."#text" = "$xPathValue"
                $settings++
            }
            elseif ($PSCmdlet.ParameterSetName -eq "xpathadd")
            {
                Write-Host "Adding node"
                $oldInnerXml = $node.InnerXml
                $newInnerXml = [string]::Concat($oldInnerXml, $xPathValue)
                $node.InnerXml = $newInnerXml
                $settings++
            }
            elseif ($PSCmdlet.ParameterSetName -eq "xpathremove")
            {
                Write-Host "Removing node"

                if (Supports-Removal $node)
                {
                    $node.ParentNode.RemoveChild($node) | out-null
                    $settings++
                }
                else
                {
                    throw "Node $(Get-NodePath $node) is not part of one of the supported hierarchies for removal"
                }
            }
        }

        #Save updated settings back to the file
        $Config.Save($configPath)
        Write-Host "Successfully updated or added $settings Config settings in config for Service: $svcName."
    }

    if ($RestartService)
    {
        RestartService $svcName
    }
}

function RestartService($svcName)
{
    Write-Host "Restarting $svcName to pick up changed configuration..."
    
    [string]$ExchangeInstallDir = GetMSIInstallPath
    $restartScriptPath =  Join-Path $ExchangeInstallDir "DataCenter\RestartService.ps1"
    & "$restartScriptPath" -ServiceName "$svcName"
}

#Main
Write-Host "Num of Parameters passed in = $($PSBoundParameters.count)"
Write-Host "Parameterset name = $($PSCmdlet.ParameterSetName)"
VerifyUsage
ConfigureWlmDelay
ConfigureSettings

# SIG # Begin signature block
# MIIdsAYJKoZIhvcNAQcCoIIdoTCCHZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUySQ7fHTOMdm1I8nlY6ljznbc
# RnigghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLYwggSyAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCByjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUT6uzMhQm9qaN15PUo64f4/jNMSMwagYKKwYB
# BAGCNwIBDDFcMFqgMoAwAEMAbwBuAGYAaQBnAHUAcgBlAE0AUgBTAFMAZQB0AHQA
# aQBuAGcAcwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNo
# YW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAiOOsJVMYIlyBOdZieMcVelGqhLH3uZ5B
# yCWnkobFXDWOgxhQ0wi8HPpxQrFm9dhWIxa4qMvRdP0p6Q1U7Fwfhu5ks+MKtMV7
# jWe3PxuLjMpFETUjIjGARtRw3zbfbOB2bwvX+TyVT96rN3wE4fuveqVkHQQMujr8
# IDeAbFY9rHzh9C5msw7rox4p0hS9VUhiOkENFRVT2cj9ebgPOsQTogOtvZ/OUl2A
# fpfKaDk8AX7QfLwSHt3l+i/zdj+hMQib9JXKTrQi27DtUNE4yXvx3PKYC6Qu4p3l
# l7BgiYyoKd5cddEzbQR8r9I1n4UnXS5c9Vva+ixmpy2cSuU5Cb1XW6GCAigwggIk
# BgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0ECEzMAAACampsWwoPa1cIAAAAAAJowCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ0M1owIwYJ
# KoZIhvcNAQkEMRYEFAqAB1OwPbipd0R+xLzXicVhXZwjMA0GCSqGSIb3DQEBBQUA
# BIIBAIEY0QDNhEgxR/ZAkm+myVwlchbJeswf+gotpqCwVqoGQ8vygG4j/C7cjpsX
# pAh6btQoP2Fb3XFcUZzT24dTAZ8Jcc/Y1eLTXkZ76hS2NEkraNhIsTPVpj6qU3+U
# UHqRb37NADoQytmkiA7aFqBNuzhj12zpRRywWrfvtBvduG/wAKGWeG7OX/39Aoby
# oh3mpOTPW2RZZtTCrc3hVlm+x/sxGaB9igtEs06OJ+5115Cq/o+dtsxr1d+1BBy0
# 7dwEUxKGeVo7hhBBBbOmJPIEaQvjLC9trYZN1sMWALDwfc4/6KcpZztNRMFMHJWh
# 9+zqEWGk0JGyQaWZRtH3p6IQLxk=
# SIG # End signature block
