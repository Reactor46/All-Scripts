
# Gets the base path for the Exchange Setup
function GetExchangeSetupPath
{
    $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("SOFTWARE\Microsoft\ExchangeServer\v15\Setup");
    
    return $key.GetValue("MsiInstallPath");
}

# Gets the base path for an application
# <param name="application">Name of the application: eg: ews</param>
function GetCafeApplicationBasePath
{
    param ([string] $application);
    
    $exchangeSetupPath = GetExchangeSetupPath;
    
    if ($application -eq 'RoutingUpdates')
    {
        return [System.IO.Path]::Combine($exchangeSetupPath, "FrontEnd\$application");
    }
    else
    {
        return [System.IO.Path]::Combine($exchangeSetupPath, "FrontEnd\HttpProxy\$application");
    }
}

# Gets the base path for an application
# <param name="application">Name of the application: eg: owa, airsync or ecp</param>
function GetBackendApplicationBasePath
{
    param ([string] $application);
    
    $exchangeSetupPath = GetExchangeSetupPath;
    
    if($application -eq 'servicehost')
    {
        return [System.IO.Path]::Combine($exchangeSetupPath, "bin\");
    }
    elseif($application -eq 'ews')
    {
        return [System.IO.Path]::Combine($exchangeSetupPath, "ClientAccess\exchweb\ews");
    }
    else
    {
        return [System.IO.Path]::Combine($exchangeSetupPath, "ClientAccess\$application");
    }
}

# Gets the full path of the web.config file for an application
# <param name="application">Name of the application: eg: ews</param>
function GetCafeWebConfigFullPath
{
    param ([string] $application);
    
    $applicationBasePath = GetCafeApplicationBasePath($application);
    return [System.IO.Path]::Combine($applicationBasePath, "web.config");
}

function GetBackendWebConfigFullPath
{
    param ([string] $application);
    
    $applicationBasePath = GetBackendApplicationBasePath($application);
    return [System.IO.Path]::Combine($applicationBasePath, "web.config");
}

# Create the CustomHeader nodes for Cafe Role
function CreateCustomHeadersNodeForCafe
{
    param ([string] $application);

    $webConfigFilePath = GetCafeWebConfigFullPath $application; 
    CreateCustomHeadersNodeForCafeInWebConfig $webConfigFilePath;
}

# Create the CustomHeader nodes for Cafe Role in Web config
function CreateCustomHeadersNodeForCafeInWebConfig
{
    param ([string] $webConfigFilePath);

    $xmlDocument = New-Object System.Xml.XmlDocument;
    $xmlDocument.Load($webConfigFilePath);

    CreateCustomHeadersNode $xmlDocument "X-FEServer" $env:COMPUTERNAME;
    DeleteCustomHeadersNode $xmlDocument "X-DiagInfo"
    
    $xmlDocument.Save($webConfigFilePath);
}

# Create the CustomHeader nodes for BE Role
function CreateCustomHeadersNodeForBackend
{
    param ([string] $application);

    $webConfigFilePath = GetBackendWebConfigFullPath $application;
    $xmlDocument = New-Object System.Xml.XmlDocument;
    $xmlDocument.Load($webConfigFilePath);

    CreateCustomHeadersNode $xmlDocument "X-BEServer" $env:COMPUTERNAME;
    CreateCustomHeadersNode $xmlDocument "X-DiagInfo" $env:COMPUTERNAME;
    
    $xmlDocument.Save($webConfigFilePath);
}

# Writes the specified string to the log
#<param name="message">String to Log</param>
#<param name="logType">Type of the log. It can be: "Error", "Warning", "Verbose"</param>
function WriteLog
{
    param ([string] $message, [string] $logType = "Verbose");

    $date = get-Date -format "MM/dd/yyyy hh:mm:ss";
    $script:logContent += "[$date] [$logType] $message" + [System.Environment]::NewLine;
    
    if($logType -eq "Error")
    {
        write-error $message;
    }
    elseif($logType -eq "Warning")
    {
        write-warning $message;
    }
    elseif($logType -eq "Verbose")
    {
        write-verbose $message;
    }
}

# Add an attribute to a given XML Element
# <param name="xmlDocument">Document where the attribute will be added</param>
# <param name="xmlElement">Element where the attribute will be added</param>
# <param name="attributeName">Name of the attribute</param>
# <param name="attributeValue">Value of the attribute</param>
function AddXmlAttribute
{
    param ([System.Xml.XmlDocument] $xmlDocument, [System.Xml.XmlElement] $xmlElement, [string] $attributeName, [string] $attributeValue);


    $attribute = $xmlDocument.CreateAttribute($attributeName);
    $temp = $attribute.set_Value($attributeValue);
    $temp = $xmlElement.SetAttributeNode($attribute);
}

# Create a CustomHeaders node
function CreateCustomHeadersNode
{
    param ([System.Xml.XmlDocument] $xmlDocument, [string] $headerName, [string] $value);

    # Finds the system.webServer/httpProtocol/customHeaders node. If it doesn't exist we create one
    $customHeadersNode = $xmlDocument.SelectSingleNode("configuration//system.webServer/httpProtocol/customHeaders");
    if($customHeadersNode -eq $null)
    {
        $httpProtocolNode = $xmlDocument.SelectSingleNode("configuration//system.webServer/httpProtocol");
            
        if ($httpProtocolNode -eq $null)
        {
            $httpProtocolNode = $xmlDocument.CreateNode([System.Xml.XmlNodeType]::Element, "httpProtocol", $null);
            $systemWebServerNode = $xmlDocument.SelectSingleNode("configuration//system.webServer");
            $temp = $systemWebServerNode.AppendChild($httpProtocolNode);
        }

        $customHeadersNode = $xmlDocument.CreateNode([System.Xml.XmlNodeType]::Element, "customHeaders", $null);
        $temp = $httpProtocolNode.AppendChild($customHeadersNode);
    }

    $DiagInfoHeaderNode = $xmlDocument.SelectSingleNode("configuration//system.webServer/httpProtocol/customHeaders/add[@name='$headerName']");
    if ($DiagInfoHeaderNode -eq $null)
    {
        WriteLog "Adding $headerName Custom Headers to Web.config file." "Warning";
        [System.Xml.XmlNode] $DiagInfoHeaderNode = $xmlDocument.CreateNode([System.Xml.XmlNodeType]::Element, "add", $null);
            
        AddXmlAttribute $xmlDocument $DiagInfoHeaderNode "name" "$headerName";
        AddXmlAttribute $xmlDocument $DiagInfoHeaderNode "value" "$value";

        $temp = $customHeadersNode.PrependChild($DiagInfoHeaderNode);
    }
    else
    {
        WriteLog "Web.config file already had $headerName CustomHeaders." "Warning";
    }
}

# Deletes a CustomHeaders node
function DeleteCustomHeadersNode
{
    param ([System.Xml.XmlDocument] $xmlDocument, [string] $headerName);

    # Finds the system.webServer/httpProtocol/customHeaders node. If it doesn't exist we create one
    $customHeadersNode = $xmlDocument.SelectSingleNode("configuration//system.webServer/httpProtocol/customHeaders");
    if($customHeadersNode -eq $null)
    {
        return;
    }

    $DiagInfoHeaderNode = $xmlDocument.SelectSingleNode("configuration//system.webServer/httpProtocol/customHeaders/add[@name='$headerName']");
    if ($DiagInfoHeaderNode -eq $null)
    {
        WriteLog "Web.config file does not have $headerName CustomHeaders." "Warning";
    }
    else
    {
        
        WriteLog "Removing $headerName Custom Headers to Web.config file." "Warning";
        $temp = $customHeadersNode.RemoveChild($DiagInfoHeaderNode);
    }
}
# SIG # Begin signature block
# MIIdwAYJKoZIhvcNAQcCoIIdsTCCHa0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0i0udUmp6BKEXwqubiVj8201
# upigghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMYwggTCAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB2jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUm9QHteAobpmflfgVeG52g/XCL24wegYKKwYB
# BAGCNwIBDDFsMGqgQoBAAEMAbwBuAGYAaQBnAHUAcgBlAEMAYQBmAGUAUgBlAHMA
# cABvAG4AcwBlAEgAZQBhAGQAZQByAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBADOACBH0xjiZ
# VzzgwzpJNIzfOEVeoOHz4cWAPuaZ7GvPzzLpR5drqqXexQlIQf2jaVEE4LSTKaaF
# tdJmeVrDhHdsfQbdZENzDabEYQPc1SFr1ouykkegKMuPTxen0fwsbhU9EljgUkkP
# iIAuFI7+6IUUInj22Aq//E+/Ru53C1CFeClLzVJ8D2q+jEMRonn2aynVgSWYHCDJ
# DTrKEh+Z+admB9/8u5A/pwOGlLI6YR1WPYGMwTJ2m9hN05QcCMgDF866hPv8NYvz
# zISb5IpEA9jdVPqf9fEhP2vK0YaiWBwqSN3SfMGVYaGHLO0loQWV3ayJ//LkhlXh
# EzFJdOw21p6hggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmarFgZ+Mon2KAAAAAACZMAkGBSsOAwIa
# BQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0x
# NjA5MDMxODQ0NDNaMCMGCSqGSIb3DQEJBDEWBBSVJMdrLT8BHvPHz6TQnmrmrxwQ
# TTANBgkqhkiG9w0BAQUFAASCAQBqtj2ZO5TlEc3QVhwj6IeY2SG/I/y0SiIgIkCV
# 07v30nt26u1DJLlr4W+nYjxbPewXqYv1FybVXqBLEAZ3Neao3jell+5dDprtGvKU
# zsMBJQPzl6JM6MFJXcT7SOJVH+bcSW76HhDmjKBXg7YUW5V6wtftTOd2Swu+rrCM
# jT3eyp755cJJr7wqlooLfFnnpltl2JGAogafPzqXwkspKVjnDeVUXu9DF8VbST+a
# fI41AX/GdNlO1lyabdvbBsb5ua2CG1/jBRUM9WhuaIZSuMaGM8wR7oQwIpTPURk8
# 9hctRA5qU2iPgZctf1Up59eb4DnM02rN97kfYY6Vy6O712jO
# SIG # End signature block
