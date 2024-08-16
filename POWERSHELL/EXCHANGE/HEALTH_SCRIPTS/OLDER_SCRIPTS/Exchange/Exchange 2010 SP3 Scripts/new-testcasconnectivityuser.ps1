# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

# Synopsis: This script creates a user that can be used for testing connectivity 
#           for Client Access Servers. This script has to be run by an admin who 
#           has permissions to create users in the Active Directory. 
#


function CreateTestUser
{
param($exchangeServer, $mailboxServer, $securePassword, $OrganizationalUnit, $UMDialPlan, $UMExtension, $Prompt)

  #
  # Create the CAS probing user with UPN.  The user will be searched for by the probing using UPN.  Note that this task must be run on
  # every mailbox server.
  #
  $adSiteGuidLeft13 = $exchangeServer.Site.ObjectGuid.ToString().Replace("-","").Substring(0, 13);
  $UserName = "extest_" + $adSiteGuidLeft13;
  $SamAccountName = "extest_" + $adSiteGuidLeft13;
  $UserPrincipalName =  $SamAccountName + "@" + $exchangeserver.Domain
  
  $err= $null
  
  #
  # Create the mailbox if the user doen't exist.  Otherwise just look up this user's mailbox
  #
  $newUser = $null
  $newUser = get-Mailbox $UserPrincipalName -ErrorAction SilentlyContinue -ErrorVariable err
  
  if ($newUser -eq $null)
  {
    #
    # If there are multiple mailbox databases on this server, the user will be created in the last database returned
    #
    $mailboxDatabaseName = $null;
    get-MailboxDatabase -server $mailboxServer | foreach {$mailboxDatabaseName = $_.Guid.ToString()}
  
    if ($mailboxDatabaseName -ne $null)
    {
      write-host $new_testcasuser_LocalizedStrings.res_0000 $exchangeServer.Fqdn 
      if ($Prompt -eq $true)
      {
          read-host $new_testcasuser_LocalizedStrings.res_PromptToQuitOrContinue
      }
  
      new-Mailbox -Name:$UserName -Alias:$UserName -UserPrincipalName:$UserPrincipalName -SamAccountName:$SamAccountName -Password:$SecurePassword -Database:$mailboxDatabaseName  -OrganizationalUnit:$OrganizationalUnit -ErrorVariable err -ErrorAction SilentlyContinue
      $newUser = get-Mailbox $UserPrincipalName -ErrorAction SilentlyContinue
      
      if ($newUser -eq $null)
      {
          $err = "Mailbox could not be created. Verify that OU ( $OrganizationalUnit ) exists and that password meets complexity requirements."
      }
    }
    else
    {
      $err = $new_testcasuser_LocalizedStrings.res_0002
    }
  }
  else
  {
      write-host $new_testcasuser_LocalizedStrings.res_0003 $exchangeServer.Fqdn
      if ($Prompt -eq $true)
      {
          read-host $new_testcasuser_LocalizedStrings.res_PromptToQuitOrContinue
      }
  }
  
  if ($newUser -ne $null)
  {
    write-Host $new_testcasuser_LocalizedStrings.res_0005 $newUser.UserPrincipalName

    set-Mailbox $newUser -MaxSendSize:1000KB -MaxReceiveSize:1000KB -IssueWarningQuota:unlimited -ProhibitSendQuota:1000KB -ProhibitSendReceiveQuota:unlimited -HiddenFromAddressListsEnabled:$True

    # Provide the newly creted user with Remote PowerShell support
    Set-User $newUser -RemotePowerShellEnabled:$true

    #
    # Reset the credentials and save them in the system
    #
    test-ActiveSyncConnectivity -ResetTestAccountCredentials -MailboxServer:($mailboxServer) -ErrorAction SilentlyContinue -ErrorVariable err
  }
 
  # check if user is UM enabled. If not try to UM enable if passed correct parameters

  $simpleuser = get-mailbox -id $UserName -ErrorAction SilentlyContinue
  $umuser = get-ummailbox -id $UserName -ErrorAction SilentlyContinue
  if (($simpleuser -ne $null) -and ($umuser -eq $null))
  {
    # if user exists and not UM enabled 
    if ($UMDialPlan -ne $null)
    {
        # if UMDialPlan was specified - showing intent to UM enable the user

        while($true)
        {
            # loop until you find a valid dialplan 

            $dialplan = get-umdialplan -id $UMDialPlan -ErrorAction SilentlyContinue -ErrorVariable err
            if ($dialplan -ne $null)
            {
                $policy = $dialplan.UMMailboxPolicies[0]    
                break;  
            }
            else
            {
                write-host  
                $UMDialPlan= (read-host $new_testcasuser_LocalizedStrings.res_0006)  
            }
        }

        [int] $num = $dialplan.NumberOfDigitsInExtension
        
        while($true)
        {
            # loop until you find a valid UM extension

            if($UMExtension.Length -ne $num)
            {
                write-host  
                write-host ($new_testcasuser_LocalizedStrings.res_0007 -f $num)   
                $UMExtension = (read-host $new_testcasuser_LocalizedStrings.res_0008)    
            }
            else
            {
                break;
            }
        }
        
        # UM enable the user. Any error thrown from the task should be reported to the user by the err variable
        
        Enable-UMMailbox -id $UserName -Pin '12121212121212121212' -PinExpired $false -UMMailboxPolicy $policy -Extensions $UMExtension -ErrorAction SilentlyContinue -ErrorVariable err
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
Import-LocalizedData -BindingVariable new_testcasuser_LocalizedStrings -FileName new-TestCasConnectivityUser.strings.psd1

$UMDialPlan  = $null
[string]$UMExtension = 0

$OrganizationalUnit = "Users"

$ArgsErrorMessage = $new_testcasuser_LocalizedStrings.res_0009

# check whether admin wants to UM enable the test user. If so .. check if he has specified the right parameters.

$securePassword = $null
$Prompt = $true

if ($args.Count -gt 0) 
{
    write-Host
    if(($args.Count % 2) -ne 0)
    {
        write-Host $ArgsErrorMessage
        write-Host
        exit
    }
    $i = 0
    while($i -lt $args.Count)
    {
        switch($args[$i])
        {
            { $_ -eq "-OU" } 
            { $OrganizationalUnit = $args[$i + 1] }

            { $_ -eq "-Password" } 
            { $securePassword = $args[$i + 1] 
              $Prompt = $false
            }

            { $_ -eq "-UMDialPlan" } 
            { $UMDialPlan = $args[$i + 1]
              write-Host ($new_testcasuser_LocalizedStrings.res_0010 -f $UMDialPlan)
            }

            { $_ -eq "-UMExtension" }
            { $UMExtension = $args[$i + 1] 
              write-Host ($new_testcasuser_LocalizedStrings.res_0011 -f $UMExtension)
            }
            
            default         
            {   write-Host $ArgsErrorMessage
                write-Host
                exit
            }
        }
        $i = $i + 2
    }
    
    write-Host
}

if ($securePassword -ne $null)
{
    # Make sure that the password parameter is a SecureString
    if ($securePassword.GetType() -ne [System.Security.SecureString])
    {
        write-host $new_testcasuser_LocalizedStrings.res_0021
        write-host
        exit
    }
}
else
{
    $new_testcasuser_LocalizedStrings.res_0012
    # Enter password
    $securePassword = (read-host -asSecureString $new_testcasuser_LocalizedStrings.res_EnterPasswordPrompt)
}

$result = $true
$atLeastOneServer = $false
$pipedInput = $false
$expectedMailboxServerType = "Microsoft.Exchange.Data.Directory.Management.MailboxServer"
foreach ($mailboxServer in $Input)
{
  $pipedInput = $true
  if ($mailboxServer.GetType().ToString() -ne $expectedMailboxServerType)
  {
    write-Host ($new_testcasuser_LocalizedStrings.res_0014 -f $mailboxServer,$mailboxServer.GetType().ToString(),$expectedMailboxServerType)
    continue;
  }
  $exchangeServer = get-ExchangeServer $mailboxServer

  $result = CreateTestUser $exchangeServer $mailboxServer $securePassword $OrganizationalUnit $UMDialPlan $UMExtension $Prompt
  $atLeastOneServer = $true
}

if ((!$atLeastOneServer) -and (!$pipedInput))
{
  $exchangeServer = get-ExchangeServer $(hostname.exe) -ErrorAction:SilentlyContinue
  if ($exchangeServer -ne $null)
  {
    if ($exchangeServer.IsMailboxServer)
    {
      $mailboxServer = get-MailboxServer $exchangeServer.Fqdn
      $result = CreateTestUser $exchangeServer $mailboxServer $securePassword $OrganizationalUnit $UMDialPlan $UMExtension $Prompt
      $atLeastOneServer = $true
    }
  }
}

if (!$atLeastOneServer)
{
  write-Host
  write-Host $new_testcasuser_LocalizedStrings.res_0015
  write-Host $new_testcasuser_LocalizedStrings.res_0016
  write-Host
  write-Host $new_testcasuser_LocalizedStrings.res_0017
  write-Host
  write-Host $new_testcasuser_LocalizedStrings.res_0018
  write-Host
  write-Host $new_testcasuser_LocalizedStrings.res_0019
  write-Host
}

if($result -eq $true -and $UMDialPlan -eq $null)
{
    write-Host
    write-Host $new_testcasuser_LocalizedStrings.res_0020
    write-Host
}

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUxN/KzT9+b2ODFGO1duRsl9JL
# czmgghhqMIIE2jCCA8KgAwIBAgITMwAAAR4S9EGOKUc2xgAAAAABHjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM2
# WhcNMjAwMTEwMjEwNzM2WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046QUUyQy1FMzJCLTFBRkMxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDzb28mkTnE/Qx2JfKv+ykRWkxyTx2Gt1TN7wBh/G2D9Z9J
# oWGDsHVyLMxbJzirFVKII+4mj6qKjCbMhWWhrUcKaez6q7hh6tpNL/knhCj48lZe
# FYMbSNAcY+cgHt5i+xV3kh6sZz1x8VrNopcripG3IjrbIJ5a47/NMUVCZyLwn+V8
# yz0elFnqH52amVPanbS0Re4Mku1U9IEOAdhlFd1AMfNL4kumj3GucM+W1rL9jsRO
# 9kgSsMDFwsM5lDAhn3toZcapx0yMi961g05xhpSmZ/hI4+szlAyqH0HN1CXjq2XQ
# 6PhRYcn4o+BdUzYbJ8rSrU3VwUhOzhv7hdTl5R2DAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQU3NGCgrg8lSUew8D8IjHvi1eXPDUwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAWR2z6mBmVxajzjhf
# 10GitwpRJTFbD1PjMopI0EPjoQiZUNQk4pxBLpQMSTv983jET+IHM6tU58aH8zI9
# 6GRYxgqVC4fNWWmZZq+OJ8+kLC95j65TbHarqzjZeW8x7nbZBd+l27sDbgyE99YA
# m9LwKecAYJY4IOcC2vl1CwdBzVMvnwN+mbgHw5X1hEdrjRODR0Fq0p/Yp6olDZ+4
# 8Wytf1U2gnOxM+3oMIg5OMnZ36pvAU05trHyX3/sx4vv/iKnuenE4tnK7MVqF4Jd
# u49bNdNxrivTf7UIluolvjaOIfnePwHajCAKRQLcHcD9LgtFg5PFEFhx64v52YnZ
# YAWspDCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# vDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUbEDsVPR/EpHCfsHzUfwisr8Xb7kw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAD1O2fDeDZKZZHDoIlvJk7q7j+V5Yh/3PVpiJjMUi9lx
# YR997Y2HEb0JSYH4MwPOiZ0Giipc8BCTBf9K7j+Tnca3M0LKBCz6h6tZBdyWf46c
# hMn6y2GDDVxoTcXoXNB4qEgw2Ut3z3Gl6Qg/2FhGOtHNblxcIIzmoCqNySU0rhQ1
# FsxZP2n9m49HA/IWPYW62K1JN/NF4BmWiL7DfEr8LZpZ+9vpnrXErM7NIakdQq/x
# zG3I/JutxN7ia86JydsF/QNbQJTE7J3xgysigfQTCvAdnUZZxz4ymcBB08dcJlzk
# jbsy9VO/+5nK5MND6gPzmYOD64B4a/jyr8k2FQfagN6hggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# HhL0QY4pRzbGAAAAAAEeMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBQCYVwlZqCdR1BxdOuEP3dc//ecOzANBgkqhkiG9w0BAQUFAASCAQAT4TvR
# eqo61QgElfcAB8Rmjkh/19ga+RdxbPi+QBa5JZkjuzaVuSHcWxYMbGBjSQUambYd
# 6NXe5pY15IUQ8TxKjkuoN72XHJyl2cWLDcZ0hEZUkS4xcuEA3SsbDtcqXJONEWka
# tKG3UyncdaXAoiU5IIBTTxZfhdw29xvdsBmgf6BtqSnrufxZef2v3n22PoZwALkE
# 7AuwU71rCjDalUJpIg1wdcMc//TsllbN0wEj9pVQ+1RIZWUs5LhG199X0YdlGVAV
# oWO4WxswGk0i7Fu5FnwGO6+a3O3fbz3t00tIuPQbmwoGtBbezgmYO3WQKXMJfkh/
# bSM/1CFLR95Nui5/
# SIG # End signature block
