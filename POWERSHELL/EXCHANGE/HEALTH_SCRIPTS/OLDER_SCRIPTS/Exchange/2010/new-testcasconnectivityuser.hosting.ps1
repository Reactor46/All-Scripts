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
# MIIagwYJKoZIhvcNAQcCoIIadDCCGnACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf39pTqDBS0X9yZDBCnJt1G7w
# TQGgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggTH
# MIIEwwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHpMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQY+us3TxUTvMl5wSqp+50paY+A0TCBiAYKKwYBBAGCNwIBDDF6MHigUIBOAG4A
# ZQB3AC0AVABlAHMAdABDAGEAcwBDAG8AbgBuAGUAYwB0AGkAdgBpAHQAeQBVAHMA
# ZQByAC4ASABvAHMAdABpAG4AZwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAT1p1ISIPqoK+6aAS
# qYKyD2ksbw4fHFTZbGOZFKbXO3TbJyAT79fvtT1PR0hxDjNkwIxCMxNaTSDvAubl
# RW+Pwc8p3gFpahZXZVW0imJIM6PprK2J5/otwzGV99TSJEQ5rXpBkNhJI0kQG1oC
# YiF8gjRn6JAdMsQENwjTima6VJ6N6KTICR21IqVWKwaFa5uyXxYCLaX1vQ0jysd7
# P1+HkVNpKvWfqU7mqliEm10pR7UQiq0Rgz2W/UPpQgNJWhdwuA8H/HvJJC+APOUu
# ekw2Zq2YGF2EeuxR5ZH1PXVQ6l+4MwIta8ysO4SHs2DIDNbE9tHAtwNwaVdLYvO9
# AwSVOKGCAh8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0ECCmECjkIAAAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3
# DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIwNTA2MzcyNVow
# IwYJKoZIhvcNAQkEMRYEFEoBLTWt7x5c6EwjQFXnPh98mxTGMA0GCSqGSIb3DQEB
# BQUABIIBADHE1adJU5TVLzul3aWq4TxJUQpGu5eufeYVmnq0Zt+t8ZEn7hPePS1u
# 7RC8Oo7CgHZJy/rCPsvf8dS2hle+1aUwAURqtZQrY+RYqN1j0+RSU0TmInU0Ovr+
# U2Ekr4EpZwS2o3b6qlWLSqob67L5FUjsTdGBTDhYoxQR58dB8UCxKWhuHeaQJwDG
# BBpzmkY8jBxvQjgTx81g0DCgnrCnNbXAACAcCsCItYXuvwNuiFnL1vJIKloDdbnL
# 0zRHh40AIdIMzeLhGavhe4DQdQtEEA0yD0AdKatAVGysB+Rzs9vcAiODS3oB4idB
# lIIV7oFDt05ke/MyQebtR/9xnFjiZDQ=
# SIG # End signature block
