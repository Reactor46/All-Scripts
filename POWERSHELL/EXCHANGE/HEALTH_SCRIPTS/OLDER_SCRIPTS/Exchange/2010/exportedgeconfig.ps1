# Copyright (c) Microsoft Corporation. All rights reserved.  
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
    "Get-TransportServer $MachineName",
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
    $setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
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
# MIIaXAYJKoZIhvcNAQcCoIIaTTCCGkkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/WhjJIZ14RVQVoXciEZ/+BTg
# a2+gghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSg
# MIIEnAIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHCMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBRoR5CEYrIEnuiTLUvjsXRx+CBAzTBiBgorBgEEAYI3AgEMMVQwUqAqgCgARQB4
# AHAAbwByAHQARQBkAGcAZQBDAG8AbgBmAGkAZwAuAHAAcwAxoSSAImh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAuSIA
# vRRfW6viLf3uY8vzv3gMK8hKjT3DZDeVJX5/DaAyp4HXW1lE26qt90jWx8tNDJ3H
# Mc1qURsVIeqVsLNPepPvkVMAzUE1XZ/lc2UY+URkZBbfxA/RGOu8yVNlzXh1Iq80
# QcVzB0rUGd7jRZF5qsBX7twYUasaka+/Fsoywuz/i6xdv678Y0V+89dJx45dBkvN
# xfNv/EmicMuWnxnzKFQkuFPQW9nOMeC9EK448kMrFREm2srkD1PQCzzUMEumurDi
# tfk25A9pE12omPUbOtc53VgSav3syVrVF9LI+ymBdCJiqJRKK15lZSSga2C01Ilc
# XnL8DIkWBR8aA8vXDqGCAh8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhN
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECCmECjkIAAAAAAB8wCQYFKw4DAhoFAKBd
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIw
# NTA2MzcyM1owIwYJKoZIhvcNAQkEMRYEFGtNADa23oGWxfXrtC2OrJy95y1PMA0G
# CSqGSIb3DQEBBQUABIIBAFm4gHWTuS+qNPjQYpkC5DAdai0IAOXPPqfR9U6OsLhb
# iX8eg0CldhhT8iYFMFKfVZfS19MipvZvub5A1BLfVjgBadA37n2ZbTgNBCOGg4Ya
# LrNRXE430PKmVuBDDuW8X1tdUrR3Yos4NQZBxrI7CAR2ypXnbLKma8ENsgmjaPUn
# 17yyYhUxdSKu1bxh7K4JwQBFEZgzbK8LLWgM65UsohFATRaF6Jq1JNz07TG6xhqe
# BaxYFKSg91kaNKVieE+7XrY4TnChPMyvHVD4PpR6e4llXyXHU6sRTw6dHSDgyefO
# XnEqpVSbpR3biiWOd8ikQng3chXS4eVmMAxWDpdojME=
# SIG # End signature block
