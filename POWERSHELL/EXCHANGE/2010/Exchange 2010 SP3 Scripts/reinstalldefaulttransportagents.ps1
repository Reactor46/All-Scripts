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
# MIIaewYJKoZIhvcNAQcCoIIabDCCGmgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfXNuDCFiiE+mlOouCH7mEs3h
# gqmgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggS/
# MIIEuwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHhMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBTh2/raTVq9JgBQr9FugJ5IbicF2DCBgAYKKwYBBAGCNwIBDDFyMHCgSIBGAFIA
# ZQBpAG4AcwB0AGEAbABsAEQAZQBmAGEAdQBsAHQAVAByAGEAbgBzAHAAbwByAHQA
# QQBnAGUAbgB0AHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# ZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAAZBRsrUsn8c/qJwRR+KKIEt771C
# zf1SZs3p8d2YpuLfEp9rHVOi/ncED34nWd39u4eCqSUF8s3nB58LfPRU+fjeHl8P
# yv4Ydn4PKs10Qfbw+K4G4WzX1Oqc5A992y2gAj0k51Xm5qGrT22K+No9mmUrE1U2
# RQWzO0kECdxb3WK+n+YtfSeTB/Ga32PIGsMM8PJJaAGJt4eeRbnnd5K8DyryJ7+i
# OuhViBjghuvnPcxgk+He34GT0IakMsfbL1Yf5eXnkaPBWGwtwUVnDCVJbl0579FO
# qBQivVwVGkAbWSAXcKuiE/GxzOmw7p4AZ+Zx+cPJnRf6u85kOINfENmd8I+hggIf
# MIICGwYJKoZIhvcNAQkGMYICDDCCAggCAQEwgYUwdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAgphAo5CAAAAAAAfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJ
# KoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xMzAyMDUwNjM3MjVaMCMGCSqGSIb3
# DQEJBDEWBBTK8TNT1R99ibzRyMD8oTSY3ZuG8DANBgkqhkiG9w0BAQUFAASCAQB+
# W+igL8G0LBph0BxEu09ysBP8wlCDwFnk3Ps7lnuTe/va3uk7Dhx7/medHGZa1mwo
# nPo8h/2tiCGVEMsFs9nY9lkq7H8lwRlco7LoYDEAr6ZuB48PlAg0uBDK29h/IZER
# 5Ogz7lIHuFVv84jEqtnsZ5N3LnPmi3P1DM2fB1yzvwIClusVMLKzcdWwTfNWe0Ig
# j1zHzEhXoUI872P2VHi+jK2FFl9/SPg6ib+3xaqW7oXl9lPEo2+g4Kob+2KNAD+W
# Me/hOu2gPVKS+AwEHNOtyWeMQ8PX6TD+dsiER6VZjZSj3CoWF5WYA91qO4G4DxUM
# VJDXnwF1mHUilqGlsgQE
# SIG # End signature block
