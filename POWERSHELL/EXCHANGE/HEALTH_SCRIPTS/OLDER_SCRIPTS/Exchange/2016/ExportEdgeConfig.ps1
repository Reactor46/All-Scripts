﻿# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

#
# Synopsis: Exports Edge configuration to an XML file, use in conjunction with 
#    import-EdgeConfig script. 
# 
# Usage: 
# 
#    .\exportedgeconfig -cloneConfigData:"C:\cloneConfigData.xml" 
#                    -Key:"A9ABA4D2C21C4bc58B303EA47BBE3608" (32 byte string used
#                     for password encryption/decryption)
#

param($cloneConfigData, $key = "A9ABA4D2C21C4bc58B303EA47BBE3608")

#############################################################################
# Logs the information to the cloneLogFile.log
# isHost : If true Write the message to the host, inaddition to the log file.
# message : The message string to be logged. 
#############################################################################
function Write-Information([Boolean] $isHost, [String] $message)
{
    # Log all the messages to the log file.
    if($logger -ne $null)
    {  
        $logger.WriteLine();
        $logger.WriteLine($message)        
    }	
    
    # Write the error message to the host also.
    if($isHost -eq $true)
    {
        write-host $message 
    }
}

###########################################################################
# The function adds some useful information into the log file
# before the actual Export process starts.
###########################################################################
function Write-LogStartInformation()
{
    $startTime = ([System.DateTime]::Now).ToString()
    Write-Information $false "************ BEGIN EXPORT PROCESS **************"
    Write-Information $false $startTime
}

#############################################################
# Releases the file handles used by the Export Script.
#############################################################
function Export-ReleaseHandles()
{
    write-debug "Export-ReleaseHandles"
    if($logger -ne $null)
    {  
        Write-Information $false "************ END EXPORT PROCESS **************"
        $logger.flush()
        $logger.close()    
    }
    if($xmlWriter -ne $null)
    {  
        $xmlWriter.Close()
    }
}

###############################################################
# Finds the parameters of the objects the command returns. Gets 
# the objects using the command, iterates through the parameters
# of each object and creates the XML component for each object.
# getCommand: The command to be run
# writer: Writer to write exported data to
###############################################################
function ExportObjects([String] $getCommand, [System.Xml.XmlTextWriter] $writer)
{
    write-debug ("ExportObjects " + $getCommand)
    # Get all the object we need to export using the given command.
    # The command may return:
    #   - null if there are no objects to export
    #   - single object
    #   - array of objects.
    
    $resultObjects = invoke-expression $getCommand
    $resultObjectsCount = 0

    # If no objects returned Log and return.
    if (!$resultObjects)
    {
        Write-Information $false "$getCommand returned 0 objects"
        return
    }

    # If we are here, there is at least one object to export.
    # First, extract the type of objects being exported.
    $objectType = GetCommandNoun $getCommand

    foreach ($object in $resultObjects)
    {
        # If we have not done it yet, get all the object properties using the
        # Get-Member cmdlet.
        if ($objectPropertyNames -eq $null)
        {
            $objectPropertyNames = Get-Member -InputObject $object -MemberType Property
        }
        
        # Create a new XML element for the current object.
        $writer.WriteStartElement($objectType)
        
        # Save all the properties.
        foreach ($prop in $objectPropertyNames)
        {
            if ($object.($prop.Name) -eq $null)
            {
                $writer.WriteElementString($prop.Name, '$null')          
            }
            elseif ($prop.Definition.ToString().ToUpper().Contains("MULTIVALUEDPROPERTY"))
            {
                if ($object.($prop.Name).Count -eq 0)
                {
                    $writer.WriteElementString($prop.Name, '$null')
                    continue;    
                }
                $multiValuedDataItem = $null 
                foreach($multiValuedItem in $object.($prop.Name))
                {
                    if ($multiValuedDataItem -ne $null)
                    {
                        $multiValuedDataItem = $multiValuedDataItem + ", "
                    }
                    $multiValuedDataItem = $multiValuedDataItem + 
                                           "'" +
                                           $multiValuedItem.ToString() +
                                           "'"                
                } 
                $writer.WriteElementString($prop.Name, $multiValuedDataItem)
            }
            elseif ($prop.Definition.ToString().ToUpper().Contains("SYSTEM.MANAGEMENT.AUTOMATION.MSHCREDENTIAL"))
            {
 
                # Create a new XML element for the current Parameter.
                $writer.WriteStartElement($prop.Name)

                # Convert the PSCredential into a clear text "Domain\Username" string and 
                # a "Password" string encrypted with the Key parameter
                $credential = $object.($prop.Name)

                $securePassword = $credential.Password

#               $securePassword = ConvertTo-SecureString -String $encryptedPassword -key $encryptionKey
                $encryptedPassword = ConvertFrom-SecureString -SecureString $securePassword -key $encryptionKey

                $writer.WriteElementString("UserName", $credential.UserName)        
                $writer.WriteElementString("Password", $encryptedPassword)        

                write-debug ("credential.UserName=" + $Credential.UserName)
                write-debug ("credential.Password=" + $Credential.GetNetworkCredential().Password)

                # Finish the Parameter element.
                $writer.WriteEndElement()
            }
            else
            {
                $writer.WriteElementString($prop.Name, $object.($prop.Name))        
            } 
        }
        
        # Finish the element.
        $writer.WriteEndElement()

        # Increment the Objects Count counter.
        $resultObjectsCount++	
    }
    
    Write-Information $false "$getCommand returned $resultObjectsCount objects"

    return
}

############################################
# Extract a noun from a given command
# command: The command to be parsed for noun
# returns the noun part of the command.
############################################
function GetCommandNoun([String] $command)
{
    write-debug "GetCommandNoun"
    # Take the part that goes after '-' and if there are any parameters then
    # cut them off.
    
    $nounIndex = $command.IndexOf('-') + 1
    $nounEndIndex = $command.IndexOf(' ', $nounIndex)

    if ($nounEndIndex -eq -1)
    {
        $nounEndIndex = $command.Length
    }
    
    $nounLength = $nounEndIndex - $nounIndex
    return $command.Substring($nounIndex, $nounLength)
}

################################################################
# Return the list of items to clone 
# returns:  List of items to be cloned.
################################################################
function ReadCloneItems()
{
    write-debug "ReadCloneItems"

    $cloneItems = @(	
    "Get-TransportService $MachineName",
    "Get-AcceptedDomain",
    "Get-RemoteDomain",
    "Get-TransportAgent",
    "Get-SendConnector",
    "Get-ReceiveConnector",
    "Get-ContentFilterConfig",
    "Get-SenderIdConfig",
    "Get-SenderFilterConfig",
    "Get-RecipientFilterConfig",
    "Get-AddressRewriteEntry",
    "Get-AttachmentFilterEntry",
    "Get-AttachmentFilterListConfig",
    "Get-IPAllowListEntry | where {`$_.IsMachineGenerated -eq `$false}",
    "Get-IPAllowListProvider",
    "Get-IPAllowListConfig",
    "Get-IPAllowListProvidersConfig",
    "Get-IPBlockListEntry | where {`$_.IsMachineGenerated -eq `$false}",
    "Get-IPBlockListProvider",
    "Get-IPBlockListConfig",
    "Get-IPBlockListProvidersConfig",
    "Get-ContentFilterPhrase",
    "Get-SenderReputationConfig",
    "Get-TransportConfig"
    )

    return $cloneItems
}

######################################################
# Retrieves the Root setup registry entry.
# returns: return entry value of found else null
######################################################
function GetEdgeInstallPath()
{
    write-debug "GetEdgeInstallPath"
    # Get the root setup entires.
    $setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
    $setupEntries = get-itemproperty $setupRegistryPath
    if($setupEntries -eq $null)
    {
        return $null
    }

    # Try to get the Install Path.
    $installPath = $setupEntries.MsiInstallPath
    
    return $installPath
}

######################################################
# Retrieves the Setup Registry value of the given entry.
# setupProperty: Setup Setting to the retrieved
# returns: return entry value of found else null
######################################################
function GetEdgeScriptsPath([String] $scriptFileName)
{
    write-debug "GetEdgeScriptsPath"

    # Append the script path with file name.
    $scriptFilePath = $installPath + "Scripts\" + $scriptFileName
    
    return $scriptFilePath
}

################################################################
# Validates the input parameters.
# returns: True if validation succeeds else False.
################################################################
function ValidateInput()
{
    write-debug "ValidateInput"
    if($cloneConfigData -eq $null)
    {
        write-debug "ValidateInput"
        Write-Information $true "Input parameters are not set"
        return $false
    }

    return $true
}


#############################################################################
# InitializationCheck
# Setup script variables, check arguments etc.
#############################################################################
function InitializationCheck()
{
    write-debug "InitializationCheck"
    # Global variable for storing handle to the clone log file.
    # Change this to Actual Path relative to [C:\Program Files\Microsoft\Exchange Server\]

    $script:success = $false

    $script:installPath = GetEdgeInstallPath
    $script:logfilePath = $installPath + "Logging\SetupLogs\cloneLogFile.log"

    $script:logger = new-object System.IO.StreamWriter ($logfilePath , $true)

    if ($key -eq $uniqueKey)
    {
       write-host -ForegroundColor "red" "WARNING: Passwords will be encrypted with a default script encryption key"
    }

    # Check/Setup the SecurePassword encryption key 
    if (($key.Length -ne 32) -and
        ($key.Length -ne 24) -and
        ($key.Length -ne 16))
    {
       Write-Information $true "Key Length needs to be 16, 24 or 32 bytes long." 
       return
    }

    $script:encryptionKey = $key.ToCharArray()
    write-debug ("encryption key=" + $encryptionKey)

    Write-LogStartInformation

    # Return error if any extra parameters are supplied.
    if($script:args.Count -gt 0)
    {
        foreach ($arg in $script:args)
        {
            Write-Information $true "Invalid additional parameter $arg passed."
        }
        return
    }

    # Adam Service Name
    $AdamServiceName = "ADAM_MSExchange"

    # Make sure ADAM_MsExchange Service is running.
    $adamServiceStatus = get-service $AdamServiceName

    if(($adamServiceStatus -eq $null) -or 
       ($adamServiceStatus.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Running))
    {
        Write-Information $true "Adam service should be running, in order to export clone config."
        return
    }
    
    # Initialize "global" variables that can be used in the template.
    $script:MachineName = [System.Environment]::MachineName

    $script:edgeServer = get-ExchangeServer -Identity:$MachineName | where { $_.IsEdgeServer -eq $true }
    if (!$script:edgeServer)
    {
        Write-Information $true "Please run the Export script on the Edge Server." 
        return
    }

    $isValidInput = ValidateInput
    if ($isValidInput -eq $false)
    {
        return    
    }

    $script:success = $true	
}


#############################################################################
# Main Script starts here, validates the parameters. Defines the list of 
# cloneable items. Gets the data corresponding
# to each item and adds them to the clone data file.
#############################################################################

#Usage 
#:: ./exportedgeconfig 
#-cloneConfigData:"C:\cloneConfigData.xml"

# Do all the Initialization Checks
InitializationCheck
if (-not $success)
{
    Export-ReleaseHandles
    exit
}

write-debug "Main"

# Create and initialize XML text writer.

$xmlWriter = new-object System.Xml.XmlTextWriter(
    $cloneConfigData,
    [System.Text.Encoding]::UTF8)
    
$xmlWriter.Formatting = [System.Xml.Formatting]::Indented
$xmlWriter.WriteStartDocument()
$xmlWriter.WriteStartElement("ExportedEdgeConfiguration")

# Read all get commands that we want to use for exporting configuration data
$getCommands = ReadCloneItems

# Export each category of objects.
foreach ($getCommand in $getCommands)
{
    ExportObjects $getCommand $xmlWriter
}

# Close the XML file.
$xmlWriter.WriteEndElement()
$xmlWriter.WriteEndDocument()
$xmlWriter.Close()

# Exception handling.
trap
{
    Write-Information $true "Exporting Edge configuration information failed."
    Write-Information $true ("Reason: " + $error[0])
    Export-ReleaseHandles
    exit
}

Write-Information $true "Edge configuration is exported successfully to $cloneConfigData"
Export-ReleaseHandles

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCTvFldppIjOese3Kslkqr2Lu
# MdmgghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjaPiz4GL18u/
# A6Jg9jtt4tQYsDcF1Y02nA5zzk1/ohCyfEN7LBhXvKynpoZ9eaG13jJm+Y78IM2r
# c3fPd51vYJxrePPFram9W0wrVapSgEFDQWaZpfAwaIa6DyFyH8N1P5J2wQDXmSyo
# WT/BYpFtCfbO0yK6LQCfZstT0cpWOlhMIbKFo5hljMeJSkVYe6tTQJ+MarIFxf4e
# 4v8Koaii28shjXyVMN4xF4oN6V/MQnDKpBUUboQPwsL9bAJMk7FMts627OK1zZoa
# EPVI5VcQd+qB3V+EQjJwRMnKvLD790g52GB1Sa2zv2h0LpQOHL7BcHJ0EA7M22tQ
# HzHqNPpsPQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFJaVsZ4TU7pYIUY04nzHOUps
# IPB3MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACEds1PpO0aBofoqE+NaICS6dqU7tnfIkXIE1ur+0psiL5MI
# orBu7wKluVZe/WX2jRJ96ifeP6C4LjMy15ZaP8N0OckPqba62v4QaM+I/Y8g3rKx
# 1l0okye3wgekRyVlu1LVcU0paegLUMeMlZagXqw3OQLVXvNUKHlx2xfDQ/zNaiv5
# DzlARHwsaMjSgeiZIqsgVubk7ySGm2ZWTjvi7rhk9+WfynUK7nyWn1nhrKC31mm9
# QibS9aWHUgHsKX77BbTm2Jd8E4BxNV+TJufkX3SVcXwDjbUfdfWitmE97sRsiV5k
# BH8pS2zUSOpKSkzngm61Or9XJhHIeIDVgM0Ou2QwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBK4wggSqAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUopQsxGnpm/Fy5MrKcp5X3FmqqHMwYgYKKwYB
# BAGCNwIBDDFUMFKgKoAoAEUAeABwAG8AcgB0AEUAZABnAGUAQwBvAG4AZgBpAGcA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAIWDinOobniJEA5JLwZvl44s6AJce077VT+96rEoELW4
# jFE9WmkOASu/EbhPrGpFlAblXBLW/jdrAcUElpilM79JU3npUJDvt4tjH6i0dsdR
# fAwoNgqY2jfy3aCg6aMag6Gd+W75EFvv77FqRjuklB4/0CgeKvJR4+DkCr3uc7zC
# fQXHDn/SnD4yLMXuP4roLmiscciiXQT97uyrHogwWw2Z8tQc5ZX5GZoss58AZory
# m37fugnebqYZuBvb1+Aghaf064Q1gF6VGoiDEiuNjJxdkQ9dv3hQCQJJ2CKATc1k
# CDcbsnDhdsDTHUoXd+zLd4kdUxDKwoAhAzEwWNBaew+hggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAA
# m+B0N8s9TY0uAAAAAACbMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQzMThaMCMGCSqGSIb3DQEJ
# BDEWBBRod86stV6MoRFty4iVgT4p07UBbjANBgkqhkiG9w0BAQUFAASCAQBh3Gqp
# hb55m4WfGbAGnnQOFYPoCPIdD+VjIUkCQkDQ29s60QBxAo99H+M+27vEymlLMfup
# baUPM5VjXhOTPvmX3kr4jLyH3KQBcVDIAlIx1PB3CdNojGwEEYUaV3Ct3UAiFGKN
# yPnUN2BMK7yZilfSLbKyL6xMe609RqG/OXWKk+o8BCt4Y91RcNfBniPj0NVVgtNn
# CwSVi8I8Onrw0ng0lFQqZGmis/556tuAxunxb72ihozfBuIgi65NcTMzIrQIPJ14
# 9nFY08d8UPMW5b7U8JHH5viw9IjA4ePf3B1vYvF0Mr5L6399CJ5dGqYO/OtI5Nw1
# 5p26obs4L4+w4ijE
# SIG # End signature block