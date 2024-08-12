# Copyright (c) Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Synopsis: Installs and enables default transport agents on Edge and Hub roles.
#
# Usage:
#
#    .\ReinstallDefaultTransportAgents.ps1
#

$forceConfirm = $true
# Exchange install path
$serverPath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MSiInstallPath

# transport agent install path
$agentPath = $serverPath + "TransportRoles\agents\"

# transport agent configuration path
$agentConfigPath = $serverPath + "TransportRoles\Shared\"

# transport agent configuration file
$agentConfigFile = "agents.config"

# transport agent configuration file with full path
$agentConfigFilePath = $agentConfigPath + $agentConfigFile

# old transport agent configuration
$agentOldConfigFile = $agentConfigFile + ".old"

# old transport agent configuration file with full path
$agentOldConfigFilePath = $agentConfigPath + $agentOldConfigFile

# show warning before changing anything
$originalPreference = $WarningPreference
if ($forceConfirm)
{
    $WarningPreference = "Inquire"
}
Write-Warning ("This script will restore the default Exchange Transport Agent configuration. The current configuration will be backed up to '" + $agentOldConfigFilePath + "'.")
$WarningPreference = $originalPreference

# local server's name
$localservername = hostname

$exchangeServer = Get-ExchangeServer -Identity $localservername

# server object associated with the local server
if ($exchangeServer -eq $null)
{
    # log an error if exchange server not found
    Write-Host ("Failed to find the Exchange server: " + $localservername)
    return
}

# if server has Edge role
$isEdge = $exchangeServer.IsEdgeServer

#if server has Hub role
$isHub = $exchangeServer.IsHubTransportServer

# create the folder if it does not exist
if (!(Test-Path -PathType Container $agentConfigPath))
{
    New-Item $agentConfigPath -Type directory > $null
}

# rename the original config file if it already exists
if (Test-Path -PathType Leaf $agentConfigFilePath)
{
    if (Test-Path -PathType Leaf $agentOldConfigFilePath)
    {
        Remove-Item $agentOldConfigFilePath
    }
    Rename-Item $agentConfigFilePath $agentOldConfigFile
}

$ConnectionFilteringAgent =
    @("Connection Filtering Agent",
      "Microsoft.Exchange.Transport.Agent.ConnectionFiltering.ConnectionFilteringAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$ContentFilterAgent =
    @("Content Filter Agent",
      "Microsoft.Exchange.Transport.Agent.ContentFilter.ContentFilterAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$SenderIdAgent =
    @("Sender Id Agent",
      "Microsoft.Exchange.Transport.Agent.SenderId.SenderIdAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$SenderFilterAgent =
    @("Sender Filter Agent",
      "Microsoft.Exchange.Transport.Agent.ProtocolFilter.SenderFilterAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$RecipientFilterAgent =
    @("Recipient Filter Agent",
      "Microsoft.Exchange.Transport.Agent.ProtocolFilter.RecipientFilterAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$ProtocolAnalysisAgent =
    @("Protocol Analysis Agent",
      "Microsoft.Exchange.Transport.Agent.ProtocolAnalysis.ProtocolAnalysisAgentFactory",
      "Hygiene\Microsoft.Exchange.Transport.Agent.Hygiene.dll",
      $true)
$AddressRewritingInboundAgent =
    @("Address Rewriting Inbound Agent",
      "Microsoft.Exchange.MessagingPolicies.AddressRewrite.FactoryInbound",
      "EdgeMessagingPolicies\Microsoft.Exchange.MessagingPolicies.EdgeAgents.dll",
      $true)
$EdgeRuleAgent =
    @("Edge Rule Agent",
      "Microsoft.Exchange.MessagingPolicies.EdgeRuleAgent.EdgeRuleAgentFactory",
      "EdgeMessagingPolicies\Microsoft.Exchange.MessagingPolicies.EdgeAgents.dll",
      $true)
$AttachmentFilteringAgent =
    @("Attachment Filtering Agent",
      "Microsoft.Exchange.MessagingPolicies.AttachFilter.Factory",
      "EdgeMessagingPolicies\Microsoft.Exchange.MessagingPolicies.EdgeAgents.dll",
      $true)
$AddressRewritingOutboundAgent =
    @("Address Rewriting Outbound Agent",
      "Microsoft.Exchange.MessagingPolicies.AddressRewrite.FactoryOutbound",
      "EdgeMessagingPolicies\Microsoft.Exchange.MessagingPolicies.EdgeAgents.dll",
      $true)
$TransportRuleAgent =
    @("Transport Rule Agent",
      "Microsoft.Exchange.MessagingPolicies.TransportRuleAgent.TransportRuleAgentFactory",
      "Rule\Microsoft.Exchange.MessagingPolicies.TransportRuleAgent.dll",
      $true)
$JournalingAgent =
    @("Journaling Agent",
      "Microsoft.Exchange.MessagingPolicies.Journaling.JournalAgentFactory",
      "Journaling\Microsoft.Exchange.MessagingPolicies.JournalAgent.dll",
      $true)
$PrelicensingAgent =
    @("Prelicensing Agent",
      "Microsoft.Exchange.MessagingPolicies.RmSvcAgent.PrelicenseAgentFactory",
      "RmSvc\Microsoft.Exchange.MessagingPolicies.RmSvcAgent.dll",
      $true)
$UMPlayonPhoneAgent =
    @("UM PlayOnPhone Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.UMPlayonPhoneAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$UMPartnerMessageAgent =
    @("UM Partner Message Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.UMPartnerMessageAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$ContentAggregationAgent =
    @("Content Aggregation Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.ContentAggregationAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$InboundSmsDeliveryAgent =
    @("Mobile Message Receive Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.SmsDeliveryAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$MailboxRulesAgent =
    @("Mailbox Rules Agent",    
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.MailboxRulesAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)      
$MeetingMessageProcessingAgent =
    @("Meeting Message Processing Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.MeetingMessageProcessingAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$MeetingForwardNotificationAgent =
    @("Meeting Forward Notification Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.MfnSubmitterAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$ApprovalProcessingAgent =
    @("Approval Processing Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.ApprovalProcessingAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$ApprovalSubmitAgent =
    @("Approval Submit Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.ApprovalSubmitterAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$ConversationsProcessingAgent =
    @("Conversations Processing Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.ConversationsProcessingAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)
$MRMDeliveryAgent =
    @("Message Records Management Delivery Agent",
      "Microsoft.Exchange.MailboxTransport.StoreDriver.Agents.RetentonPolicyTagProcessingAgentFactory",
      "..\..\bin\Microsoft.Exchange.StoreDriver.dll",
      $true)

$EdgeAgents =
    @($ConnectionFilteringAgent,
      $AddressRewritingInboundAgent,
      $EdgeRuleAgent,
      $ContentFilterAgent,
      $SenderIdAgent,
      $SenderFilterAgent,
      $RecipientFilterAgent,
      $ProtocolAnalysisAgent,
      $AttachmentFilteringAgent,
      $AddressRewritingOutboundAgent)

$HubAgents =
    @($InboundSmsDeliveryAgent,
      $ConversationsProcessingAgent,
      $TransportRuleAgent,
      $JournalingAgent,
      $PrelicensingAgent,
      $UMPlayonPhoneAgent,
      $UMPartnerMessageAgent,
      $MailboxRulesAgent,
      $MeetingMessageProcessingAgent,
      $MeetingForwardNotificationAgent,
      $ApprovalProcessingAgent,
      $ApprovalSubmitAgent,
      $ContentAggregationAgent,
      $MRMDeliveryAgent)

# generate agent list
$agents = @()
if ($isEdge)
{
    $agents += $EdgeAgents
}
if ($isHub)
{
    $agents += $HubAgents
}

# install agents
$originalPreference = $WarningPreference
$WarningPreference = "SilentlyContinue"
foreach ($agent in $agents)
{
    $name = $agent[0]
    $factory = $agent[1]
    $agentAssembly = $agentPath + $agent[2]
    $enabled = $agent[3]
    Install-TransportAgent -Name:$name -TransportAgentFactory:$factory -AssemblyPath:$agentAssembly > $null
    if ($enabled)
    {
        Enable-TransportAgent -Identity:$name
    }
}
$WarningPreference = $originalPreference

# display the current agent status
Get-TransportAgent
Write-Host ""
Write-Warning "The Transport Agents shown above have been re-installed. Please exit Powershell and restart the MS Exchange Transport service for the change to take effect."

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfXNuDCFiiE+mlOouCH7mEs3h
# gqmgghhqMIIE2jCCA8KgAwIBAgITMwAAASDzON/Hnq4y7AAAAAABIDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM4
# WhcNMjAwMTEwMjEwNzM4WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MjI2NC1FMzNFLTc4MEMxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCO1OidLADhraPZx5FTVbd0PlB1xUfJ0J9zuRe1282yigKI
# +r7rvHTBllcSjV+E6G3BKO1FX7oV2CGaAGduTl2kk0vGSlrXC48bzR0SAb1Ui49r
# bUJTA++yfZA+34s8vYUye1XX2T5D0GKukK1hLkf8d7p2A5nygvMtnnybzmEVavSd
# g8lYzjK2EuekiLzL/lYUxAp2vRNFUitr7MHix5iU2nHEG4yU8crlXjYFgJ7q3CFv
# Il1yMsP/j+wk+1oCC1oLV6iOBcpq0Nxda/o+qN78nQFoQssfHoA9YdBGUnRHk+dK
# Sq5+GiV3AY0TRad2ZRzLcIcNmUJXny26YG+eokTpAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUIkw9WwdWW+zV8Il/Jq7A7bh6G7cwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAE4tuQuXzzaC2OIk4
# ZhJanhsgQv9Tk8ns/9elb8pAgYyZlSwxUtovV8Pd70jtAt0U/wjGd9n+QQJZKILM
# 6WCIieZFkZbqT9Ut9zA+tc2eQn4mt62PlyA+YJZNHEiPZhwgbjfLIwMRsm845B4N
# KN7WmfYwspHdT/mPgLWaBsSWS80PuAtpG3N+o9eTHskT+qauYAMqhZExfI8S2Rg4
# kdqAm7EU/Nroe4g0p+eKw6CAQ2ZuhuqHMMPgcQlSejcEbpS5WAzdCRd6qDXPHh0r
# C3FayhXrwu/KKuNW2hR1ZCx/ieNiR8+lWt1JxXgWAttgaRtR3VqGlL4aolg41UCo
# XfN1IjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU4dv62k1avSYAUK/RboCeSG4nBdgw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBADYVg+T4vhwMtYLWtXDI2Rut8FRydrD70B2TJWK1h7J8
# DnLL+Cs+6EUqYtGKSh1tOy0vR2DJBdpi/WjqQCSK2Uw+q0ThSGrFhT4KXPPwLI8t
# /SslUr7nUAL1Y0Gc3qKBYUnYt4BJv8BfeCBX3iPCs03I1mAsTrwUrtU+278zJsHZ
# JbWx0rfOrmziLTDJurJjgEL08/d5nhcT7HSv7guvEt6b7oKNIzwPRBRDlNe7RorG
# LXHezmSRbPrJLqQpNMNO3G5Moa1HlMz/mBO7qZdWHXUNJs5kO4DlVpCgnakgaSgs
# 7nSOjmJiNQsBlBdlzsGxLs5mnr4DifEPczMRj3zWJ2ShggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# IPM438eerjLsAAAAAAEgMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDJaMCMGCSqGSIb3DQEJ
# BDEWBBSo+yrAjnX2j+/rcOLmxhLdkDYBbDANBgkqhkiG9w0BAQUFAASCAQARww4+
# c4NfhX/IwnhuL9DsyHDp3cODJMqnQP1HUIQZTs+PFr1tAaTYeVEmS7JymFaJypxu
# sndS+OoHmLm8zUXqJyVdvR+Xf/eewgRkiNvQ0K6p25L6j6VREGL5V6YlSq+79lEH
# Bqf7VjnmTzfDmqpD8O8S8frOKagHEicjiFfvIyAekkeHYjyTDAvn+nA5VHFre9fP
# rQBDNhMlD5YlPld5RxlRqC3x92QOQkv3Jmh4m+AN4Dfeqrc/2JsVX+8Le3lygN9V
# +Hn5cOBhifz3wt/TrlmM+8zrX1nLeEPfC+Ie6L7aX6PIjVpkE6BMN+Xi8NNcyBHG
# AGSGx6XHopli81YO
# SIG # End signature block
