# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

# Synopsis: This script creates a user that can be used for testing connectivity 
#           for Client Access Servers. This script has to be run by an admin who 
#           has permissions to create users in the Active Directory. 
#

#
# This function generates a test user name based on a name prefix and length.
#

function GenerateTestUserName
{
param($namePrefix, $totalChars)

    $HexDigits = "0123456789ABCDEF"
    $GenCount = $totalChars - $namePrefix.Length
                    
    $provider = new-object System.Security.Cryptography.RNGCryptoServiceProvider
    $builder = new-object char[] $GenCount
    $data = new-object byte[] 4

    for($num = 0; ($num -lt $GenCount); $num++)
    {
        $provider.GetBytes($data)
        $index = ([System.BitConverter]::ToUInt32($data, 0) % $HexDigits.Length)
        $builder[$num] = $HexDigits[$index]
    }
    
    $name = $namePrefix
    foreach ($char in $builder)
    {
        $name += $char
    }
    
    return $name
}

#
# This function is used to generate a cryptographically-secure random password.
#

function GenerateSecureRandomPassword
{
param($PasswordLength)

    $UpperCaseLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $LowerCaseLetters = "abcdefghijklmnopqrstuvwxyz"
    $NumberChars = "0123456789"
    $SymbolChars = "`~!@#$%^*()_+-={}|[]:`";`'?,./"
    $PasswordChars = $LowerCaseLetters + $UpperCaseLetters + $NumberChars + $SymbolChars
                    
    $provider = new-object System.Security.Cryptography.RNGCryptoServiceProvider
    $builder = new-object char[] $PasswordLength
    $data = new-object byte[] 4

    for($num = 0; ($num -lt $PasswordLength); $num++)
    {
        $provider.GetBytes($data)
        $index = ([System.BitConverter]::ToUInt32($data, 0) % $PasswordChars.Length)
        $builder[$num] = $PasswordChars[$index]
    }

    $provider.GetBytes($data)
    $num2 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)

    # Write some lower case characters to meet complexity requirements
    #
    for($num = 0; ($num -lt $num2); $num++)
    {
        $provider.GetBytes($data)
        $num3 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)
        $provider.GetBytes($data)
        $builder[$num3] = $LowerCaseLetters[([System.BitConverter]::ToUInt32($data, 0) % $LowerCaseLetters.Length)]
    }

    $provider.GetBytes($data)
    $num2 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)

    # Write some upper case characters to meet complexity requirements.
    #
    for($num = 0; ($num -lt $num2); $num++)
    {
        $provider.GetBytes($data)
        $num3 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)
        $provider.GetBytes($data)
        $builder[$num3] = $UpperCaseLetters[([System.BitConverter]::ToUInt32($data, 0) % $UpperCaseLetters.Length)]
    }

    $provider.GetBytes($data)
    $num2 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)

    # Write some symbols to meet complexity requirements.
    #
    for($num = 0; ($num -lt $num2); $num++)
    {
        $provider.GetBytes($data)
        $num3 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)
        $provider.GetBytes($data)
        $builder[$num3] = $SymbolChars[([System.BitConverter]::ToUInt32($data, 0) % $SymbolChars.Length)]
    }

    $provider.GetBytes($data)
    $num2 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)

    # Write some numbers to meet complexity requirements.
    #
    for($num = 0; ($num -lt $num2); $num++)
    {
        $provider.GetBytes($data)
        $num3 = (([System.BitConverter]::ToUInt32($data, 0) % ($PasswordLength-1)) + 1)
        $provider.GetBytes($data)
        $builder[$num3] = $NumberChars[([System.BitConverter]::ToUInt32($data, 0) % 10)]
    }

    $securePassword = new-object System.Security.SecureString
    $plaintext = ""
    foreach($char in $builder)
    {
      $securePassword.AppendChar($char)
      $plaintext += $char
    }
    
    write-host "Generated password:" $plaintext

    return $securePassword
}
    

#
# This function deletes all test user accounts from the test organization
#

function DeleteTestUserMailboxes
{
param($OrganizationId, $namePrefix)

    $mailboxes = get-mailbox -Organization $OrganizationId  -ErrorAction SilentlyContinue | where{$_.Name -like $NamePrefix+"*" }
    if ($mailboxes -ne $null)
    {
        foreach ($mailbox in $mailboxes)
        {
            write-host "Removing obsolete test mailbox:" $mailbox.Name
            remove-mailbox $mailbox -Confirm:$false
        }
    }
}

#
# This function creates the test user
#

function CreateTestUser
{
param($exchangeServer, $mailboxServer, $verbose)

  #
  # Create the CAS probing user with UPN.  The user will be searched for by the probing using UPN.  Note that this task must be run on
  # every mailbox server.
  #
  
  #
  #
  #
  
  $NamePrefix = "Extest_"
  $ADSiteName = (get-exchangeserver $(hostname)).Site.Name
  $SamAccountName = GenerateTestUserName $NamePrefix 20
  $UserName = $SamAccountName
  $ExchangeMonOrg = $ADSiteName + ".exchangemon.net"
  $UserPrincipalName =  $UserName + "@" + $ExchangeMonOrg
  
  write-host "===============================================================" 
  write-host "This script will create test user: $UserName"
  write-host "In the custom domain: $ExchangeMonOrg" 
  write-host "On mailbox server: $mailboxServer"
  write-host "===============================================================" 
  write-host 
  read-host "Type Control-Break to quit or Enter to continue"
  
  $OfferId = 2
  $ProgramId = "HostingSample"
  $Location = "us"
  
  $err= $null
  
  $orgCreated = $false

  $org = get-Organization $ExchangeMonOrg -ErrorAction SilentlyContinue
  if ($org -eq $null)
  {
    $AdminUser = "Administrator@" + $ExchangeMonOrg
    $org = new-organization -name $ExchangeMonOrg -domainname $ExchangeMonOrg -OfferId $OfferId -ProgramId $ProgramId -Location $Location -ErrorAction SilentlyContinue
    $orgCreated = $org -ne $null
  }
  
  if ($org -eq $null)
  {
     $err = "Could not find or create organization:" + $ExchangeMonOrg
  }
  else
  {
      # If organization already existed, delete test user mailboxes
      #
      if (!$orgCreated)
      {
          DeleteTestUserMailboxes $ExchangeMonOrg $NamePrefix
      }
      
      # Password length is 16 characters
      $MaxPasswordLength = 16

      "Generating secure password for the test user."
      $SecurePassword = GenerateSecureRandomPassword $MaxPasswordLength

      # Look up this user's mailbox
      #
      $newUser = $null
      $newUser = get-Mailbox -Organization:$ExchangeMonOrg -ErrorAction SilentlyContinue | where {$_.UserPrincipalName -eq $UserPrincipalName} 

      if ($newUser -ne $null)
      {
          $err = "There's an issue with the test user account. It already exists."
      }
      else
      {
        #
        # If there are multiple mailbox databases on this server, the user will be created in the last database returned
        #
        $mailboxDatabaseName = $null;
        get-MailboxDatabase -server $mailboxServer | foreach {$mailboxDatabaseName = $_.Guid.ToString()}
      
        if ($mailboxDatabaseName -ne $null)
        {
          write-host "Creating test user $UserName on:" $exchangeServer.Name
          write-host "In mailbox database:" $mailboxDatabaseName

          new-Mailbox -Name:$UserName -Alias:$UserName -SamAccountName:$SamAccountName -Password:$SecurePassword -Database:$mailboxDatabaseName -Organization:$ExchangeMonOrg -UserPrincipalName:$UserPrincipalName -ErrorVariable err -ErrorAction SilentlyContinue
          $newUser = get-Mailbox -Organization:$ExchangeMonOrg -ErrorAction SilentlyContinue | where {$_.UserPrincipalName -eq $UserPrincipalName} 
        }
        else
        {
          $err = "The server must have a mailbox database for creating the test user."
        }
      }
      
      if ($newUser -ne $null)
      {
        write-Host "UserPrincipalName: " $newUser.UserPrincipalName

        # Provide the newly creted user with Remote PowerShell support
        Set-User $newUser -RemotePowerShellEnabled:$true

        set-Mailbox $newUser -MaxSendSize:1000KB -MaxReceiveSize:1000KB -IssueWarningQuota:unlimited -ProhibitSendQuota:1000KB -ProhibitSendReceiveQuota:unlimited -HiddenFromAddressListsEnabled:$True

        #
        # Set the credentials and save them in the system
        #

        $Credentials = new-object System.Management.Automation.PSCredential ($UserPrincipalName, $securePassword)
      
        test-ActiveSyncConnectivity -ResetTestAccountCredentials -MailboxServer:($mailboxServer) -MailboxCredential:($Credentials) -Verbose:$Verbose -ErrorAction SilentlyContinue -ErrorVariable err
      }
  }

  #
  # Output any errors that may have occurred
  #
  if ($err -ne $null)
  {
    foreach ($e in $err)
    {
        if ($e.Exception -ne $null)
        {
            write-error $e.Exception
        }
        else
        {
            write-error $e
        }
    }
    
    return $false
  }
  
  return $true
  
}

#
# Script begins here
#

$Verbose = $true

# check for specified parameters.

if ($args.Count -gt 0) 
{
    $i = 0
    while($i -lt $args.Count)
    {
        switch($args[$i])
        {
            { $_ -eq "-Verbose" }
            { $Verbose = $args[$i + 1] }
        }
        $i = $i + 2
    }
}

$atLeastOneServer = $false
$pipedInput = $false
$expectedMailboxServerType = "Microsoft.Exchange.Data.Directory.Management.MailboxServer"
$mailboxServer = $null
$exchangeServer = $null

foreach ($mailboxServer in $Input)
{
  $pipedInput = $true
  if ($mailboxServer.GetType().ToString() -ne $expectedMailboxServerType)
  {
    write-Host "Skipping: " $mailboxServer " of type " $mailboxServer.GetType().ToString() ", expected type is " $expectedMailboxServerType
    continue;
  }
  $exchangeServer = get-ExchangeServer $mailboxServer
  if ($exchangeServer -ne $null)
  {
      $atLeastOneServer = $true
      break;
  }
}

$result = $true

if ($atLeastOneServer)
{
  $result = CreateTestUser $exchangeServer $mailboxServer $Verbose
}
elseif (!$pipedInput)
{
  $exchangeServer = get-ExchangeServer $(hostname.exe) -ErrorAction:SilentlyContinue
  if ($exchangeServer -ne $null)
  {
    if ($exchangeServer.IsMailboxServer)
    {
      $mailboxServer = get-MailboxServer $exchangeServer.Fqdn
      $result = CreateTestUser $exchangeServer $mailboxServer $Verbose
      $atLeastOneServer = $true
    }
  }
}

if (!$atLeastOneServer)
{
  write-Host
  write-Host "Please either run the command on an Exchange Mailbox Server or pipe at least one mailbox server into this task."
  write-Host "For example:"
  write-Host
  write-Host "  get-mailboxServer | new-TestCasConnectivityUser.Hosting.ps1"
  write-Host
  write-Host "or"
  write-Host
  write-Host "  get-mailboxServer MBXSERVER | new-TestCasConnectivityUser.Hosting.ps1"
  write-Host
}

# SIG # Begin signature block
# MIIdzwYJKoZIhvcNAQcCoIIdwDCCHbwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf39pTqDBS0X9yZDBCnJt1G7w
# TQGgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBNUwggTRAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB6TAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUGPrrN08VE7zJecEqqfudKWmPgNEwgYgGCisG
# AQQBgjcCAQwxejB4oFCATgBuAGUAdwAtAFQAZQBzAHQAQwBhAHMAQwBvAG4AbgBl
# AGMAdABpAHYAaQB0AHkAVQBzAGUAcgAuAEgAbwBzAHQAaQBuAGcALgBwAHMAMaEk
# gCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEB
# AQUABIIBAFZmqWQH1v8KxEmgUEoPvq6RMm/O+/M1VPSjj1/HXPbQGY2ZMFkSd3+c
# cQQahNWxxSSGHsA9iVvRWUXQ4xOW0cDkrkKRfXNA0Dc7+oj1xPi6xybG2NpXn0ut
# cQP4ubE8My1mf/ZS+cLVHZOAJekFtFM8VvDp0yxHCkTPO3CL6FZMzv6VO2UVHvj4
# TmtKoTiuVu26qZdrdbkbifVjOalHH17C2wLyFgKYfZ3ya/p4cbN0wT0/m/qrTxs+
# vbil96ovnZ9qZEWuJfZ9SIkKL1gSuP7GlTf5u2Cd6UoViZlpKLXs+UFioyDy4VUq
# dxagN/EMiriTjYdO0Im4vKjBpnuejPuhggIoMIICJAYJKoZIhvcNAQkGMYICFTCC
# AhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEh
# MB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmarFgZ+Mon2K
# AAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwG
# CSqGSIb3DQEJBTEPFw0xNjA5MDMxODQzNTNaMCMGCSqGSIb3DQEJBDEWBBROdWev
# PXJ0Emy3n3BtiJCctRpLgTANBgkqhkiG9w0BAQUFAASCAQB57eRSR6WHtwbxRGIk
# WChOvtfLApSHUasTyr71X80fI/lCyne4L1d2vwy1OJwLPifUVjMcSdlLb7POyIse
# 2vxp1odVtE3kg5ULFvZW/pgbwzuq07wjIMkotT218rWzR/ghXj+6kBDdspFMd+wB
# hJn4+TcZS9NGkDud9fpVOWxCXrbaZT//kY3gzcUtFCaw0LCqnlcz5R+L3yu2Ay6O
# K/Qn/qP34nuJJouAXRQU99xWYekiu4/xJqnxR2AzcYsybjRwtsvZDzGKzy7PC1Dz
# ViFFyQ5G0WjBJY+5EFd/tXdWH0+3AgcOFU90QTracIEfXAQnrRQIkb+765PZSm+M
# LHzu
# SIG # End signature block
