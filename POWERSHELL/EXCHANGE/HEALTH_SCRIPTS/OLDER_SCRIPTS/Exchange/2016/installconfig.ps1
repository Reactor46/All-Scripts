#*==========================================================================================================
#* Script     :    InstallConfig.ps1
#*==========================================================================================================
#* Company    :    Microsoft Corporation
#*==========================================================================================================
#* Purpose    :    This script will be used to install and uninstall
#*                 post setup configurations of Search Foundation, for Exchange. 
#*
#*==========================================================================================================

# Get input
Param 
(   
    [switch]$help,
    [string]$action,
	[string]$dataFolder,
    [switch]$silent,  
    [switch]$singlenode,
    [string]$systemname,
    [ValidateRange(1024,65535)]
    [int]$baseport
)

$ErrorActionPreference = "stop"
[string]$SCRIPT:DEFAULTSYSTEMNAME = "Fsis"
if (-not $SCRIPT:SYSTEMNAME)
{
    $SCRIPT:SYSTEMNAME = $SCRIPT:DEFAULTSYSTEMNAME
}

#region Installation Modes

[string]$SCRIPT:INSTALL = "i"

[string]$SCRIPT:ATTACH = "a"

[string]$SCRIPT:UNINSTALL = "u"

#endregion

# Set the default value of the base port, if not input as an argument.
if (!$script:baseport)
{
    $script:baseport = 3800
}

#*==========================================================================================================
#* Purpose    :    Print usage details.
#*==========================================================================================================
function PrintUsage([bool]$onError = $false, [string]$errorMessage = [string]::Empty)
{
    if (-not $SCRIPT:SILENT)
    {
        Write-Host ""
        Write-Host "Usage on installation:" -ForegroundColor Cyan
        Write-Host "---------------------" -ForegroundColor Cyan
        Write-Host ""
        Write-Host ".\InstallConfig.ps1 -action {i|a} [-dataFolder path] [-single] [-systemname name] [-baseport port] [-silent]" 
        Write-Host ""
        Write-Host "  -action: Specifies the installation mode."
        Write-Host "    i – Clean installation"
        Write-Host "    a – Attach to an existing data folder"
        Write-Host ""
        Write-Host "  -dataFolder: Data folder path. This must be a path in local machine."
        Write-Host ""
        Write-Host "  -systemname: Name of system to install. Default is 'Fsis'"
        Write-Host ""
        Write-Host "  -single: Installs all components on node"
        Write-Host ""
        Write-Host "  -systemname name: Name of system to install. Default is 'Fsis'"
        Write-Host ""
        Write-Host "  -baseport: The value of the base port"
        Write-Host ""
        Write-Host "Usage on uninstallation:" -ForegroundColor Cyan
        Write-Host "-----------------------" -ForegroundColor Cyan
        Write-Host ""   
        Write-Host ".\InstallConfig.ps1 -action u"
        Write-Host ""
    }

    if ($onError)
    {
        throw $errorMessage
    }
    
    exit
}

# Whether to print help ?
if ($help -or $h)
{
    PrintUsage  
}

#region Validate Environment
#TO DO:: All these validations must be moved to C# in the long run

# Verify running PowerShell version 2.
if (($PSVersionTable).PSVersion.Major -lt [int]2)
{
    throw "This script requires PowerShell 2.0 or higher."  
}

# Verify running with elevated privileges.
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

if ($identity)
{
    $principal = new-object System.Security.Principal.WindowsPrincipal($identity) 
    $administrator = [System.Security.Principal.WindowsBuiltInRole]::Administrator 
    [bool]$isAdmin =  $principal.IsInRole($administrator)

    if (-not $isAdmin)
    {
        throw "You are required to have administrator privileges to execute the configuration script. Aborting ..."
    }
}
else
{
    throw "You are required to have administrator privileges to execute the configuration script. Aborting ..."
}

#endregion

#region Directory and assembly Paths

$SCRIPT:INSTALL_DIRECTORY_PATH = Get-Location

if ($MyInvocation.MyCommand.Path -ne $null)
{
    $SCRIPT:INSTALL_DIRECTORY_PATH = $MyInvocation.MyCommand.Path.SubString(0, $MyInvocation.MyCommand.Path.IndexOf($MyInvocation.MyCommand.Name))
}

[string]$SCRIPT:INSTALLATION_DIRECTORY_PATH = Split-Path $SCRIPT:INSTALL_DIRECTORY_PATH -Parent

[string]$SCRIPT:POSTSETUP_CONFIGURATION_ASSEMBLY_PATH = Join-Path -Path $SCRIPT:INSTALL_DIRECTORY_PATH -ChildPath "\bin\Microsoft.Ceres.Exchange.PostSetup.dll"

[string]$SCRIPT:LOG_DIRECTORY = Join-Path -Path  $SCRIPT:INSTALL_DIRECTORY_PATH -ChildPath "log" 

#endregion

#region Set Data Directory Path

if(-not $script:dataFolder)
{
    $script:dataFolder = Join-Path -Path $SCRIPT:INSTALLATION_DIRECTORY_PATH -ChildPath "\HostController\Data"
}

#endregion

#region Construct log file path.

# Get current time stamp
$SCRIPT:CURRENT_TIME_STAMP = [DateTime]::Now.ToString("yyyyMMddHHmmss")  

switch ($script:action)
{    
    $SCRIPT:INSTALL
    { 
        $SCRIPT:LOG_FILE_NAME = [String]::Format("PostSetup_install_{0}.log", $SCRIPT:CURRENT_TIME_STAMP)
    }
    $SCRIPT:ATTACH
    {
        $SCRIPT:LOG_FILE_NAME = [String]::Format("PostSetup_attach_{0}.log", $SCRIPT:CURRENT_TIME_STAMP)
    }
    $SCRIPT:UNINSTALL
    {
        $SCRIPT:LOG_FILE_NAME = [String]::Format("PostSetup_uninstall_{0}.log", $SCRIPT:CURRENT_TIME_STAMP)
    }
    default 
    {       
        $msg = "Invalid value for action. Available values are '$SCRIPT:INSTALL' to install, '$SCRIPT:ATTACH' to attach or '$SCRIPT:UNINSTALL' to un-install."
        PrintUsage -onError $true -errorMessage $msg
    }
}
 
$SCRIPT:LOG_FILE_PATH = Join-Path  -path $SCRIPT:LOG_DIRECTORY -Childpath $SCRIPT:LOG_FILE_NAME
 
#endregion

# Catch Exceptions
trap [Exception]
{
    if ((-not $SCRIPT:SILENT) -and (-not [string]::IsNullOrEmpty($SCRIPT:LOG_FILE_PATH)) -and (Test-Path -Path $SCRIPT:LOG_FILE_PATH))
    {
        Write-Host "Script execution failed. Please inspect the log file `"$SCRIPT:LOG_FILE_PATH`" for more details." -ForegroundColor Red
    }
    
    throw $_.Exception
}
 
#*==========================================================================================================
#* Purpose    :    Validates installation action, fails on failure. 
#*==========================================================================================================
function ValidateExecuteInstallAction()
{   
    
    # At the very beginning check whether action is properly set, otherwise abort.
    if (-not $script:action)
    {
        PrintUsage -onError $true -errorMessage "You must specify whether you want to install (-action $SCRIPT:INSTALL), attach (-action $SCRIPT:ATTACH) or uninstall (-action $SCRIPT:UNINSTALL)" 
    }
    
    if ($script:args)
    {
        PrintUsage -onError $true -errorMessage "One or more invalid switches passed to the script : $SCRIPT:args"  
    } 
    
    # Check whether invalid action was given.
    switch ($script:action)
    {
        
        $SCRIPT:INSTALL
        { 
            if (IsSearchNodesAvailable)
            {
                throw "Old nodes belonging to the system '$SCRIPT:SYSTEMNAME', already exist in '$script:dataFolder'. Rerun the configuration in attach mode, if you need to reuse them."
            }

            Install
            break
        }

        $SCRIPT:ATTACH
        {
            if (-not (IsSearchNodesAvailable))
            {
                throw "Couldn’t attach the data folder '$script:dataFolder'. Path doesn’t contain old nodes belonging to the system '$SCRIPT:SYSTEMNAME'."
            }

            Install -attachMode $true
            break
        }
        $SCRIPT:UNINSTALL
        {            
            Uninstall       
            break
        }
        default 
        {       
            $msg = "Invalid value for action. Available values are '$SCRIPT:INSTALL' to install, '$SCRIPT:ATTACH' to attach or '$SCRIPT:UNINSTALL' to un-install."
            PrintUsage -onError $true -errorMessage $msg
        }
    }
}

#*================================================================================================
#* Purpose    :    Check whether the data folder contains old nodes from a previous installation.
#                  Currently, the validation is implemented by checking the presence of the NodeProfile.xml.
#                  However, this method doesn't validate the "validity" of the data folder.
#*================================================================================================
function IsSearchNodesAvailable()
{
    if (-not (Test-Path -Path $script:dataFolder))
    {
        return $false
    }
    
    $nodesPath = Join-Path -Path $script:dataFolder -ChildPath ("Nodes\" + $SCRIPT:SYSTEMNAME)
    $nodeNames = @("AdminNode1", "ContentEngineNode1", "IndexNode1", "InteractionEngineNode1")

    foreach ($node in $nodeNames)
    {
        $nodeProfilePath = Join-Path -Path $nodesPath -ChildPath ($node + "\NodeProfile.xml")
        if (Test-Path $nodeProfilePath)
        {
            return $true
        }
    }

    return $false
}

function LoadAssembly
{
    param   (
                $assemblyName = $(Throw "Must specify name of the assembly with the path")
            )
    
    trap [Exception] 
    {       
        throw "Unable load assembly $assemblyName" + $_.Exception
    }
        
    $unwantedOutput = [Reflection.Assembly]::LoadFrom($assemblyName)    
}

$SCRIPT:TRACING_SERVICE_NAME = "SearchExchangeTracing"
$SCRIPT:HOST_CONTROLLER_SERVICE_NAME = "HostControllerService"

#*==========================================================================================================
#* Purpose    :    Performs installation.
#*==========================================================================================================
function Install([bool]$attachMode = $false)
{
    trap [Exception] 
    {       
        Throw "Error occurred while configuring Search Foundation for Exchange." + $_.Exception         
    }
    
    if (-not $SCRIPT:SILENT)
    {
        Write-Host "Configuring Search Foundation for Exchange...."
    }

    # Enable ULS Tracing Service
    Set-Service -Name $SCRIPT:TRACING_SERVICE_NAME -StartupType "Automatic"

    # Enable Host Controller Service.
    Set-Service -Name $SCRIPT:HOST_CONTROLLER_SERVICE_NAME -StartupType "Automatic"
    
    # Do installation.
    LoadAssembly -assemblyName $POSTSETUP_CONFIGURATION_ASSEMBLY_PATH
    [Microsoft.Ceres.Exchange.PostSetup.DeploymentManager]::Install($SCRIPT:INSTALLATION_DIRECTORY_PATH, $script:dataFolder, $script:baseport, $SCRIPT:LOG_FILE_NAME, $SCRIPT:SINGLENODE, $SCRIPT:SYSTEMNAME, $attachMode)
    
    if (-not $SCRIPT:SILENT)
    {
        Write-Host "Successfully configured Search Foundation for Exchange"
    }
}

#*==========================================================================================================
#* Purpose    :    Performs Uninstallation.
#*==========================================================================================================
function Uninstall()
{
    trap [Exception] 
    {       
        Throw "Error occurred while uninstalling Search Foundation for Exchange." + $_.Exception         
    }
    
    if (-not $SCRIPT:SILENT)
    {
        Write-Host "Uninstalling configuration of Search Foundation for Exchange..."
    }
    
    # Do uninstallation
    LoadAssembly -assemblyName $POSTSETUP_CONFIGURATION_ASSEMBLY_PATH
    [Microsoft.Ceres.Exchange.PostSetup.DeploymentManager]::Uninstall($SCRIPT:INSTALLATION_DIRECTORY_PATH, $SCRIPT:LOG_FILE_NAME)
    
    # Enable Host Controller Service.
    Set-Service -Name $SCRIPT:HOST_CONTROLLER_SERVICE_NAME -StartupType "Disabled"
    
    if (-not $SCRIPT:SILENT)
    {
        Write-Host "Uninstallation complete"
    }
}

ValidateExecuteInstallAction
#EndLog

exit 0
# SIG # Begin signature block
# MIIdogYJKoZIhvcNAQcCoIIdkzCCHY8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmybcGoyI99s5G9xuP/niMQHD
# YEugghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBvDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU/WP9Fuh/6oTOqKXNAFxwgVQEaJ0wXAYKKwYB
# BAGCNwIBDDFOMEygJIAiAEkAbgBzAHQAYQBsAGwAQwBvAG4AZgBpAGcALgBwAHMA
# MaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3
# DQEBAQUABIIBADiARpvdh33+Vz0k6yqVCCeqUr2A0A6gBkb3k6w1wF1HDK/e5s3k
# fZe1VVVBTznNKB1hHvaX1+lPJxhr3cZuxPSTRIyiXICSe1NrjvaWsmgeTmRk0ZRI
# QxpS6T1YkM6oq2AFl0Rd5j0aLgVmfBixwNX1BMV/NBzsATgxQ9Yi1Au0cvhm0XMN
# pY/P7VmBWZBUIc/STmnLoBivJe5OmxANQXzkSwl2iGF1yD4Zic6Y+9M9EEB2BK4P
# PVlvzX52yE2oomELxlh0BNdCKPYPgfsrbS8ihhc+zVimcYp72W9ZxbiDweOAnT1m
# LTtQCwpc+CblQ2q+dJe44RY8oaO37MjtIx+hggIoMIICJAYJKoZIhvcNAQkGMYIC
# FTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmarFgZ+M
# on2KAAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODM5NDNaMCMGCSqGSIb3DQEJBDEWBBT4
# CAMnicbkEN4OqOkPmFLrL+2VKDANBgkqhkiG9w0BAQUFAASCAQANQkTrUYk1GB9N
# mxLejVcgU5JJrDEozSVpgVGyHu+gcfFG8H5ZzN4EsiPwbv9oiQgi1uCKPBP+fcin
# 7KVzhpEgrKVCfZ1apbQtniTOjb0p/EXn0TEvMOrc9bGm94Qa97pbNs60jxjukHd+
# yv1ZToYQK2gbvMgv3+80dzg4lvLeC+m6oVw5/U+N9QkjZ7qWAWkCc1he/jSYqZsc
# 4IXKCXVPQYQUjEanOWhPp6aQJT2colsMlEOBa8QcZSGy6byrCodfL+WRTx1VTCuH
# AiRDCYsBxMceab3yQnkyYz+KNkL+QDpBD7ep67lwiGlHajM1h8pGWxdpIXndCtth
# Np8ShWZZ
# SIG # End signature block
