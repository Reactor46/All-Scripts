<#
.EXTERNALHELP ExchUCUtil-help.xml
#>

# Copyright (c) 2006 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
param($Forest = $null,[Switch]$Verify)

Import-LocalizedData -BindingVariable ExchUCUtil_LocalizedStrings -FileName ExchUCUtil.strings.psd1

# Constants
$RTC_CS14_POOL_VER = 0x50000
$EUM_DP_CONTAINER = "UM DialPlan Container"
$EUM_AA_CONTAINER = "UM AutoAttendant Container"
$NotFound = "(not found)"
$ErrorList = $null;
$Verify = $false;

# Error messages
$ReadRTCPoolError = $ExchUCUtil_LocalizedStrings.res_0000;
$ReadUMDialPlanError = $ExchUCUtil_LocalizedStrings.res_0001;
$WriteUMGatewayError = $ExchUCUtil_LocalizedStrings.res_0002;
$ReadUMIPGatewayError = $ExchUCUtil_LocalizedStrings.res_0003;
$ReadPermissionsError = $ExchUCUtil_LocalizedStrings.res_0004;
$WritePermissionsError = $ExchUCUtil_LocalizedStrings.res_0005;

####################################################################################################
# Script starts here
####################################################################################################

# Helper functions
#
function HasErrors($e)      { return (($e -ne $null) -and ($e.Count -gt 0)); }
function ThrowIfError()     { if(HasErrors($ErrorList)){ throw $ErrorList[0]; } }
function WriteObject($obj)  { $obj | format-table; }

function GetRTCPools()
{
	# Execute the Get-UCPool script
	#
	$pools = .\get-ucpool.ps1 -forest:$forest
	return ($pools);
}

function GetUMIPGateways()
{
	$gateways = get-umipgateway -ev:ErrorList -ea:SilentlyContinue;
	ThrowIfError

	# Put the gateways in a hash table using the Address as the key
	$ht=@{};
	if($gateways)
	{ 
		foreach($gw in $gateways){ $ht[$gw.Address.ToString()] = $gw; } 
	}
	return ($ht);
}

function CreateConfigObject()
{
	$configEntry = new-object System.Management.Automation.PSObject;
	add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:PoolFqdn -value:$null
	add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:UMIPGateway -value:$null
	add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:DialPlans -value:$null
	return ($configEntry);
}

function ShowGateways($pools, $gateways, $dialplans)
{
	$configTable = @();

	# Verify that all UC pools have a corresponding UMIPGateway
	#
	foreach($pool in $pools.Values)
	{
		# Create a dynamic config entry object:
		#
		$configEntry = CreateConfigObject

		$configEntry.PoolFqdn = $pool.Fqdn
		if( $gateways.ContainsKey($configEntry.PoolFqdn) )
		{
			$gw = $gateways[$configEntry.PoolFqdn];
			$configEntry.UMIPGateway = $gw.Name;

			$dpList = @();
			if( $gw.HuntGroups )
			{
				foreach($hg in $gw.HuntGroups)
				{ $dpList += $hg.UMDialPlan.Name; }
			}
			else
			{
				$dpList += $NotFound;
			}
			$configEntry.DialPlans = $dpList;
		}
		else
		{
			$configEntry.UMIPGateway = $NotFound;
			$configEntry.DialPlans = $NotFound;
		}
	
		$configTable += $configEntry;
	}

	WriteObject $configTable
}

function ConfigureGateways($pools, $gateways, $dialplans)
{
	write-host $ExchUCUtil_LocalizedStrings.res_0006;
	
	$configTable = @();
	
	if( $pools.Count -eq 0 )
	{
		write-host $ExchUCUtil_LocalizedStrings.res_0007;
		return ($configTable);
	}

	foreach($pool in $pools.Values)
	{
		# Create a dynamic config entry object:
		#
		$configEntry = CreateConfigObject

		[string]$poolName = $pool.Name;
		$configEntry.PoolFqdn = $pool.Fqdn
		
		write-host ($ExchUCUtil_LocalizedStrings.res_0008 -f $configEntry.PoolFqdn)
		if( !$gateways.ContainsKey($configEntry.PoolFqdn) )
		{
			write-host $ExchUCUtil_LocalizedStrings.res_0009
			$gw = new-umipgateway -Name:$poolName -Address:$configEntry.PoolFqdn -ev:ErrorList -ea:SilentlyContinue
			ThrowIfError
			
			$gateways[$gw.Address.ToString()] = $gw;
			$configEntry.UMIPGateway = $gw.Name;
		}
		else
		{
			write-host $ExchUCUtil_LocalizedStrings.res_0010
			$gw = $gateways[$configEntry.PoolFqdn];
			$configEntry.UMIPGateway = $gw.Name;
		}
		
		# Determine whether pool represents a Branch Office Appliance (BOA) pool
		$isBranchRegistrar = $pool.Data -icontains "ExtendedType=RemoteRegistrar";
		write-host IsBranchRegistrar: $isBranchRegistrar

		# MWI support: Enabled for CS14+ and disabled for earlier versions, Disabled for BOA pools
		$mwiEnabled = ($pool.Version -ge $RTC_CS14_POOL_VER) -and ($isBranchRegistrar -eq $false);
		write-host MessageWaitingIndicatorAllowed: $mwiEnabled
		
		# Outcalling: Disabled for BOA pools
		$outcallsAllowed = ($isBranchRegistrar -eq $false);
		write-host OutcallsAllowed: $outcallsAllowed
		
		$gw | set-umipgateway -MessageWaitingIndicatorAllowed:$mwiEnabled -OutcallsAllowed:$outcallsAllowed -ev:ErrorList -ea:SilentlyContinue
		ThrowIfError
		
		# OCS pools must be linked to all dial-plans.
		$gw.HuntGroups | remove-umhuntgroup -Confirm:$false -ev:ErrorList -ea:SilentlyContinue

		# Verify this gateway is associated with all sip name dial-plans
		$dpList = @();
		if( $dialplans )
		{
			foreach($dp in $dialplans)
			{
				write-host -NoNewLine ($ExchUCUtil_LocalizedStrings.res_0011 -f $dp.Name)
				# The phone context must not change during the lifetime of the dial-plan
				$phoneCtx = $dp.PhoneContext.Split(".")[0];
				$hg = new-umhuntgroup -Name:$dp.Name -PilotIdentifier:$phoneCtx -UMDialPlan:$dp.Identity -UMIPGateway:$gw.Identity -ev:ErrorList -ea:SilentlyContinue
				ThrowIfError
				$dpList += $phoneCtx;
			}
			write-host ""
		}
		else
		{
			write-host -NoNewLine $ExchUCUtil_LocalizedStrings.res_0012
			$dpList += $NotFound;
		}
		$configEntry.DialPlans = $dpList;
		$configTable += $configEntry;
		write-host ""
	}
	return ($configTable);
}

function ContainsAccessRight($obj, [System.DirectoryServices.ActiveDirectoryRights]$r)
{
	foreach($ar in $obj.AccessRights)
	{  if(($ar -band $r) -eq $r) {return $true;}  }
	return $false;
}

function GetRTCGroup($group)
{
	if($forest -ne $null) {$group = $forest + "\" + $group;}
	return ($group);
}

function GetObjectPermissions($group, $identity, [System.DirectoryServices.ActiveDirectoryRights]$rightsNeeded)
{
	$perm = get-adpermission -Identity:$identity -User:$group -ev:ErrorList -ea:SilentlyContinue
	$result = new-object System.Management.Automation.PSObject;
	add-member -InputObject:$result -MemberType:NoteProperty -Name:ObjectName -value:$identity
	add-member -InputObject:$result -MemberType:NoteProperty -Name:AccessRights -value:$rightsNeeded
	add-member -InputObject:$result -MemberType:NoteProperty -Name:Configured -value:$false
	$result.AccessRights = $result.AccessRights.ToString()
	$result.Configured = $perm -and (ContainsAccessRight $perm $rightsNeeded)
	$result.Configured = $result.Configured.ToString()
	return ($result);
}

function SetObjectPermissions($group, $obj, [System.DirectoryServices.ActiveDirectoryRights]$rightsNeeded, $inheritanceType)
{
	if( $obj.Configured -eq $false )
	{
		write-host ($ExchUCUtil_LocalizedStrings.res_0013 -f $obj.ObjectName)
		$perm = add-adpermission -Identity:$obj.ObjectName -AccessRights:$rightsNeeded -InheritanceType:$inheritanceType -User:$group -ev:ErrorList -ea:SilentlyContinue
		ThrowIfError;
	}
	else
	{
		write-host ($ExchUCUtil_LocalizedStrings.res_0014 -f $obj.ObjectName)
	}
}

function GetPermissions($group)
{
	$resultList = @();

	$exorg = get-organizationconfig

	[System.DirectoryServices.ActiveDirectoryRights]$rights = "ListChildren"
	$result = GetObjectPermissions $group $exorg.Identity $rights
	$resultList += $result;
	
	$rights = "ListChildren,ReadProperty"
	$result = GetObjectPermissions $group $EUM_DP_CONTAINER $rights
	$resultList += $result;
	
	$rights = "ListChildren,ReadProperty"
	$result = GetObjectPermissions $group $EUM_AA_CONTAINER $rights
	$resultList += $result;

	return ($resultList);
}

function ShowPermissions($group)
{
	write-host ($ExchUCUtil_LocalizedStrings.res_0015 -f $group)
	$permissions = GetPermissions $group
	WriteObject $permissions
}

function ConfigurePermissions($group, $permissions)
{
	write-host ($ExchUCUtil_LocalizedStrings.res_0016 -f $group)
	
	#Set permissions on the Exchange Organization container
	[System.DirectoryServices.ActiveDirectoryRights]$rights = "ListChildren";
	[System.DirectoryServices.ActiveDirectorySecurityInheritance]$inheritance = "None";
	SetObjectPermissions $group $permissions[0] $rights $inheritance

	#Set permissions on the UM DialPlan and AutoAttendant Containers
	$rights = "ListChildren,ReadProperty"
	$inheritance = "All";
	SetObjectPermissions $group $permissions[1] $rights $inheritance
	SetObjectPermissions $group $permissions[2] $rights $inheritance
}

# Executes this script
function RunScript()
{
	write-host ""
	
	[System.String]$message = $null;
	trap
	{
		write-host -NoNewLine $message
		write-host -NoNewLine $ExchUCUtil_LocalizedStrings.res_0017
		write-host $_.Exception.Message;
		exit;
	}

	# get all RTC pools
	$message = $ReadRTCPoolError
	$ucObjects = GetRTCPools
	$pools = $ucObjects.RTCPools;
	
	# get admin permissions
	$message = $ReadPermissionsError
	$adminGroup = $ucObjects.RTCUniversalServerAdmins.Identity;
	$adminPermissions = GetPermissions $adminGroup
	
	# get server permissions
	$message = $ReadPermissionsError
	$serverGroup = $ucObjects.RTCComponentUniversalServices.Identity;
	$serverPermissions = GetPermissions $serverGroup

	# get all UM IP gateways
	$message = $ReadUMIPGatewayError;
	$gateways = GetUMIPGateways;
	
	# get all SIP dial-plans
	$message = $ReadUMDialPlanError;
	$dialplans = get-umdialplan | where { $_.UriType -eq "SipName" };

	if( !$Verify )
	{
		# Ensure read permissions are set for OCS admin group
		$message = $WritePermissionsError
		ConfigurePermissions $adminGroup $adminPermissions
		write-host "";
		
		# Ensure read permissions are set for OCS server group
		$message = $WritePermissionsError
		ConfigurePermissions $serverGroup $serverPermissions
		write-host "";

		# Ensure gateways exist for OCS pools
		$message = $WriteUMGatewayError;
		$configTable = ConfigureGateways $pools $gateways $dialplans;
		write-host "";
	}

	# Check whether permissions have been set for OCS
	ShowPermissions $adminGroup
	ShowPermissions $serverGroup

	# Check whether gateways exist for OCS pools
	ShowGateways $pools $gateways $dialplans;
}

# Run the script
RunScript

# SIG # Begin signature block
# MIIaYgYJKoZIhvcNAQcCoIIaUzCCGk8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzWPHta5b3EFIBPd60vp9aBC5
# 6RSgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggSdMIIEmQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIG2MBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBQOJFEiDnti3Mj0bJEdYWpsxmL3bzBWBgorBgEEAYI3AgEMMUgw
# RqAegBwARQB4AGMAaABVAEMAVQB0AGkAbAAuAHAAcwAxoSSAImh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAqsSQY3q+
# cs65LS09DdldsQpGl99ARZDR24hjWC72mdDNc4bfd3kbUCDUtVJ0TCZe6X8pxI3G
# wDjKbS7pbaj9uyu9w0dmn+pm02RloVJ6xb3+W0bQeSrWGXSLt66ndFy/Y5DZviA+
# tIkdMcvqAiOvU8J3Lk0tNY+E2x7UusCBK5ANT9XBeiGzYhKTbvYEN45orMO8P3+X
# aqEpVWJJJFiwseC7u+AT6/IiVjV1wPhLzxCnwfVbGQEUPOuOF9Bdf9iTFR7IYEF9
# 4dQHQugaoR0AEzK+SlHaw32KlUbsgsjq+XNCl1bSWJGVjWxGWvabPuihBHfof88a
# L4bgjxCCbICK26GCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAAArOTJIwbLJSPMAAAAAACswCQYFKw4D
# AhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTEzMDIwNTA2MzcyM1owIwYJKoZIhvcNAQkEMRYEFA+0sywD8FRkYpu8dUVpfOLB
# uBI3MA0GCSqGSIb3DQEBBQUABIIBAIroS916W8EHQGPTQPrrCngdPU5lq/ByTzC6
# /gJYi689CYheBbHwha/hKDhV+Vo9M6QUlumDonOVpRTLdkRVXlbvtNU245OD2lVK
# pe4ojoxOY5uLvAoRRSQsX+zYSMviZmgOBQaBztSY/dtUGaYt4qCdY4Iiga8ogqZi
# k7zQ9g7XO18kV7dMS2THmMpR7VgMkFu0ZZOTNcCisf1gwFEuwnz5/ZEzl1KCyakw
# rHSzDrdSDqqwswBqoei5ynx9ARMIXJ1kI0oOhVOmiNRK/ou1djqNs+lhzW7dpASw
# 4iArpMx9Ok7IOpyAA44CHTYYmuta+BpT4c+bMrmwnYG5V8+r1OA=
# SIG # End signature block
