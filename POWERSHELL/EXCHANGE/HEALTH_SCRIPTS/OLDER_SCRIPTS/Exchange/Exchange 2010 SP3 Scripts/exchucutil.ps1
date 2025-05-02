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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzWPHta5b3EFIBPd60vp9aBC5
# 6RSgghhqMIIE2jCCA8KgAwIBAgITMwAAASIn72vt4vugowAAAAABIjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RUFDRS1FMzE2LUM5MUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDRU0uHr30XLzP0oZTW7fCdslb6OXTQeoCa/IpqzTgDcXyf
# EqI0fdvhtqQ84neZE4vwJUAbJ2S+ajJirzzZIEU/JiTZpJgeeAMtN+MuAbzXrdyS
# ohyUDuGkuN+vVSeCnEZkeGcFf/zrNWWXmS7JsVK2BJR8YvXk0sBUbWVpdj0uvz68
# Y+HUyx8AKKE2nHRu54f6fC4eiwP/hs+L7NejJm+sNo7HXV4Y6edQI36FdY0Sotq8
# 7Lh3U96U4O6X9cD0iqKxr4lxYYkh98AzVUjiiSdWUt65DAMbdjBV6cepatwVVoET
# EtNK/f83bMS3sOL00QMWoyQM1F7+fLoz1TF7qlozAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUtlFVlkkUKuXnuF3JZxfDlHs2paYwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAZldsd7vjji5U30Zj
# cKiJvhDtcmx0b4s4s0E7gd8Lp4VnvAQAnpc3SkknUslyknvHGE77OSdxKdrO8qnn
# T0Tymqvf7/Re2xJcRVcM4f8TeE5hCaffCkB7Gtu90R+6+Eb1BnBDYMbj3b42Jq8K
# 42hnDG0ntrgv4/TmyJWIvmGQORWMCWyM/NraY3Ldi7pDpTfx9Z9s4eNE/cxipoST
# XHMIgPgDgbZcuFBANnWwF+/swj69cv87x+Jv/8HM/Naoawrr8+0yDjiJ90OzLGI5
# RScuGfQUlH0ESbzevO/9PFpoUywmNYhHoEPngLJVT2W6y13jFUx3IS9lnR0r1dCh
# mynB8jCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUDiRRIg57YtzI9GyRHWFqbMZi928w
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAEfEBkrWFV8YB3R0DZYjZ7/NZhJSOxNgtWK3vL2EhIAY
# +8kahytTyyrk+yY9RJ+hcHFF2cR8VvsKv00U4srH4V9J4KF7DiUFItG5CJuIiuhF
# sam8hNDLbhxl0m+asyF6TFMABbDiZ5WTjQbrklODbG4BLdHehtbF9zEd+avkPrdj
# BXb/k7VJ6xCxDvq1EkDylZmEyLbx7dl/6YVGB1oohsNX5YYg6/nDGlrCCm5n9yi2
# Q+BIimq5MIZ6OP/J2XtpEAIuOJz/D5CBdWFQCU95s39HXvmOjMIEMVvE2OviOdld
# tB2MeFSG2+mj/AEoiNwo82Lo50xs75RRgeLk8m660uehggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# Iifva+3i+6CjAAAAAAEiMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBQwM7XcOYF6rDoZGm4rjGJKyvCiMjANBgkqhkiG9w0BAQUFAASCAQC+vMiK
# SKhOzREpVyQz8ZyDHC6P9QADD54mJL8ZSP9j5pTwawUaBphHDmrIOpXFquOftb8Q
# hPyuxo7bkMXGqDtIJCTk3Jf7ojX0wssaoHL/zLySp88iP6PMedJ63s2cVcko6s4o
# DK8z+Kh6bROp84j/lP1mkvTiRY5MAyB2Jqv9voVX5Ja9xQhMCgFvYC7M4vlVH9Cs
# hh1v3XWnfQoDr6mOe/epI0JPjZHLyhGM26BQeTxW88gqrsRIPV179TLEfAtQvnrH
# SlZa6eFAwzeyme3Rp3iLMdVwSkibok0I1d7CoFqAvGU9KFalAih/qhCjZ67uHu0o
# 5ifp7S+gNfQRB0gv
# SIG # End signature block
