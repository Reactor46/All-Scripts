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
$EUM_SERVER_CONTAINER = "Administrative Groups"
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

	$rights = "ListChildren,ReadProperty"
	$result = GetObjectPermissions $group $EUM_SERVER_CONTAINER $rights
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

	#Set permissions on the UM DialPlan and AutoAttendant and Server Containers
	$rights = "ListChildren,ReadProperty"
	$inheritance = "All";
	SetObjectPermissions $group $permissions[1] $rights $inheritance
	SetObjectPermissions $group $permissions[2] $rights $inheritance
	SetObjectPermissions $group $permissions[3] $rights $inheritance
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
# MIIdnAYJKoZIhvcNAQcCoIIdjTCCHYkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUT+ST3euZohod2t/5udib4KKk
# JtagghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI3
# WhcNMTcwNjMwMTkyMTI3WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjdBRkEtRTQxQy1FMTQyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jclqAQB7jVZ
# CvOuH5jFixrRTGFtwMHws1sEZaA3ciobVIdWIejc5fBu3XdwRLfxjsmyou3JeTaa
# 8lqA929q2AyZ9A3ZBfxf8VqTxbu06wBj4o4g5YCsz0C/81N2ESsQZbjDxbW5sKzD
# hhT0nTzr82aepe1drjT5dvyU/AvEkCzaEDU0dZTq2Bm6NIif11GzA+OkD0bdZG+u
# 4EJRylQ4ijStGgXUpAapb0y2RtlAYvZSpLYzeFFcA/yRXacCnoD++h9r66he/Scv
# Gfd/J/5hPRCtgsbNr3vFJzBWgV9zVqmWOvZBPGpLhCLglTh0stPa/ZxZjTS/nKJL
# a7MZId131QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPPCI5/SvSWNvaj1nBvoSHO7
# 6ZPBMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD+xPVIhFl30XEe39rlgUqCCr2fXR9o0aL0Oioap6LAUMXLK
# 4B+/L2c+BgV32joU6vMChTaA+7XEw7pXCRN+uD8ul4ifHrdAOEEqOTBD7N5203u2
# LN667/WY71purP2ezNB1y+YAgjawEt6VjjQcSGZ9bTPRtS2JPS5BS868paym355u
# 16HMxwmhlv1klX6nfVOs1DYK5cZUrPAblCZEWzGab8j9d2ZIGLQmTEmStdslOq79
# vujEI0nqDnJBusUGi28Kh3Hz1QIHM5UZg/F5sWgt0EobFGHmk4KH2vreGZArtCIB
# amDc5cIJ48na9GfA2jqJLWsbvNcwC486g5cauwkwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKIwggSeAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBtjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUtvkRHAsRAZA/xUJnAfLT7aBV4qIwVgYKKwYB
# BAGCNwIBDDFIMEagHoAcAEUAeABjAGgAVQBDAFUAdABpAGwALgBwAHMAMaEkgCJo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUA
# BIIBACk/OQd4flLh90SQsGkQQi4pcCRaROcJzYIajBeaebFHOkQfMbrHg8oMPOiM
# 8o7ZnMSUxGiSZ9ix/IeMItd1p6wJIjbvwSmHSRUf9FUt6aRMJU/pp925vhAApcn0
# TRNqyr7e5NUOI4h6yd8DAnoAw3+A7cqY0krrW1/Qgxpd12KJoe7XihmLihEDdPRm
# MRgBJwS7/YpyS2pS1r53PxMeVn3rkqgeHCL+lGtrWk7+budEu0Ny2mrpY2jNZP8m
# 02rbPLn57h9cejUmNSmZj8C2RPi6fPRL0WibATmN3GbytNGPLtWSqCz+2t6CNGKj
# zP6+VO3j7E4DItmxg65BdkRn6YOhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhEC
# AQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8G
# A1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmARYy38jCbCeAAAA
# AACYMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0xNjA5MDMxODQzMTVaMCMGCSqGSIb3DQEJBDEWBBTN5G/5Vw7x
# MDnhngbXaJizUy2CsDANBgkqhkiG9w0BAQUFAASCAQArWj0ftI4x5sYau6GHt5EL
# nlmzl3stQy9BQ8wDed1Rqvo5ZGctb5/i20mRJjNcQJ5q6QpOCtPC1+Jgdz3i+Rk1
# kBzeQ2Wqcp8TzMgne69TAu6u/jLE3RmsWJb0hK4/ZGZy/8/wKfuwAUHFRCdgmXnD
# 3bm5WwcOVTZomZI6nqdZSEFxa2DvN+K6AsXIdBhxHWFdsSmnD9l6Qagy4WqaM9/y
# rcIkvncLEmYCBfH1jdoXxlm9qpDkBf4htSChbcXm2t6ftMN8BVXf2ExIsCq8JSQh
# 9A36mw8a9W0EVzzcAw2jnDr93O8z2g9fm/VosSLK8ElJq7vvreKsOo2K7eVRWtWJ
# SIG # End signature block
