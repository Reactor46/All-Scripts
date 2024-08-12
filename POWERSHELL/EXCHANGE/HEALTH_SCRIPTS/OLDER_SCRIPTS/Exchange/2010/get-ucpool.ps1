<#
.EXTERNALHELP Get-UCPool-help.xml
#>

# Copyright (c) 2006 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
param($Forest = $null)

#load hashtable of localized string
Import-LocalizedData -BindingVariable GetUCPool_LocalizedStrings -FileName Get-UCPool.strings.psd1

# Constants
#
$E_NOT_FOUND          = 0x80072030;
$E_NOT_OPERATIONAL    = 0x8007203A;
$E_INVALID_CREDENTIAL = 0x8007052E;
$RTC_MIN_POOL_VER     = 0x30000
$RTC_ADMIN_GROUP      = "RTCUniversalServerAdmins"
$RTC_SERVER_GROUP     = "RTCComponentUniversalServices";


####################################################################################################
# Script starts here
####################################################################################################

# Helper functions
#
function GetCOMProp($obj,$m){ return ($obj.psbase.GetType().InvokeMember($m,"GetProperty",$null,$obj,$null)); }
function GetLdapPath($path) { if($forest -ne $null){ return ("LDAP://"+$forest+"/"+$path); } else { return("LDAP://"+$path); } }

# Executes the LDAP query and prompts for credentials if necessary
#
function ExecuteQuery($entry)
{
	trap
	{ 
		$ec = $_.Exception.ErrorCode;
		if( !$entry.psbase.UserName -and ($ec -eq $E_INVALID_CREDENTIAL))
		{
			write-output $_.Exception.Message
			$cred = get-credential
			$nc = $cred.GetNetworkCredential();
			if($nc.Domain){$entry.psbase.UserName = $nc.Domain + "\" + $nc.UserName;} 
			else {$entry.psbase.UserName = $nc.UserName;}
			$entry.psbase.AuthenticationType = "secure" ;
			$entry.psbase.Password = $nc.Password;
			$entry.psbase.RefreshCache();
			continue;
		}
		else
		{
			throw $_.Exception;
		}
	}
	$entry.psbase.RefreshCache();
}

function ReadEntryFromSearchResult($searchResult)
{
	[System.DirectoryServices.DirectoryEntry]$entry = $null;
	if( $searchResult -ne $null )
	{
		$entry = $searchResult.GetDirectoryEntry();
		$objectDN = $entry.psbase.InvokeGet("distinguishedName");
		$entry.psbase.Path = GetLdapPath($objectDN);
		ExecuteQuery $entry
	}
	return ($entry);
}

function FindUCObjects()
{
	[string]$ldapQuery = $null;
	[System.DirectoryServices.DirectoryEntry]$entry = $null;
	[System.DirectoryServices.DirectorySearcher]$searcher = $null;

	# Initialize the result object
	#
	$result = CreateResultObject
	$result.RTCPools = @{};
	
	if( $forest -ne $null )
	{
		$ldapQuery = "GC://" + $forest + "/RootDSE"
	}
	else
	{
		$ldapQuery = "GC://RootDSE"
	}

	$entry = new-object System.DirectoryServices.DirectoryEntry $ldapQuery
	ExecuteQuery $entry;
	
	$globalCatalog = $entry.psbase.Properties["rootDomainNamingContext"].Value;
	$entry.psbase.Path = "GC://" + $globalCatalog
	write-host Using Global Catalog: $entry.psbase.Path
	write-host 

	# Find the RTC security groups using the global catalog port
	#
	$searcher = new-object System.DirectoryServices.DirectorySearcher $entry

	# Set the referral chasing:All so we can find objects in child domains
	# We also need the canonicalName to construct the object's Identity
	#
	$searcher.ReferralChasing = "All";
	$property = $searcher.PropertiesToLoad.Add("canonicalName");

	$result.RTCUniversalServerAdmins = FindRTCGroup $searcher $RTC_ADMIN_GROUP
	$result.RTCComponentUniversalServices = FindRTCGroup $searcher $RTC_SERVER_GROUP

	# Find the pool container using the global catalog port
	# 
	$entry = FindUCPools($searcher);

	# Put the pools in a hash table using the DnsHostName as the key
	#
	$pools = $entry.psbase.children;
	if($pools)
	{
		foreach($pool in $pools)
		{
			$psPool = CreatePoolObject($pool)
			[int]$version = $psPool.Version;
			if($psPool.Fqdn -and $version -ge $RTC_MIN_POOL_VER){ $result.RTCPools[$psPool.Fqdn] = $psPool; }
		}
	}
	return ($result);
}

# Create a dynamic PowerShell object to hold the pool data
#
function CreatePoolObject($adsiPool)
{
	$obj = new-object System.Management.Automation.PSObject;
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:Name -value:$null
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:Fqdn -value:$null
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:Version -value:$null
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:Data -value:$null

	$obj.Name       = $adsiPool.psbase.InvokeGet("CN");
	$obj.Fqdn       = $adsiPool.psbase.InvokeGet("dnsHostName");
	$obj.Version    = $adsiPool.psbase.InvokeGet("msRTCSIP-PoolVersion");
	$obj.Data       = $adsiPool.psbase.InvokeGet("msRTCSIP-PoolData");
	return ($obj);
}

function CreateGroupObject($searchResult)
{
	$obj = $null;
	[string]$canonicalName = $searchResult.Properties.canonicalname;
	$entry = ReadEntryFromSearchResult($searchResult);
	if( $entry -ne $null )
	{
		$obj = new-object System.Management.Automation.PSObject;
		add-member -InputObject:$obj -MemberType:NoteProperty -Name:Name -value:$null
		add-member -InputObject:$obj -MemberType:NoteProperty -Name:DistinguishedName -value:$null
		add-member -InputObject:$obj -MemberType:NoteProperty -Name:Identity -value:$null
		
		$obj.Name = $entry.psbase.InvokeGet("sAMAccountName");;
		$obj.DistinguishedName = $entry.psbase.InvokeGet("distinguishedName");
		# Get-ADPermission expects domain\account as the identity
		# Canonical name is domain/users/account
		$obj.Identity = $canonicalName.Split("/")[0] + "\" + $obj.Name;
	}
	return ($obj);
}

function FindRTCGroup($searcher, $groupName)
{
	$searcher.Filter = "(&(objectCategory=group)(SamAccountName=" + $groupName + "))";
	$searchResult = $searcher.FindOne();
	$groupObject = CreateGroupObject($searchResult);
	if( $groupObject -eq $null )
	{
		throw ($GetUCPool_LocalizedStrings.res_0000 -f $groupName);
	}
	return ($groupObject);
}

# Find the pool container using the global catalog port
# 
function FindUCPools($searcher)
{
	$searcher.Filter = "(objectCategory=msRTCSIP-Pools)";
	$searchResult = $searcher.FindOne();
	$entry = ReadEntryFromSearchResult($searchResult);
	if( $entry -eq $null )
	{
		throw $GetUCPool_LocalizedStrings.res_0001;
	}
	return ($entry);
}

function CreateResultObject()
{
	$obj = new-object System.Management.Automation.PSObject;
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:RTCPools -value:$null
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:RTCUniversalServerAdmins -value:$null
	add-member -InputObject:$obj -MemberType:NoteProperty -Name:RTCComponentUniversalServices -value:$null
	return ($obj);
}

# Run the script
#
$result = FindUCObjects;
return ($result);

# SIG # Begin signature block
# MIIaYgYJKoZIhvcNAQcCoIIaUzCCGk8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXnJKzTHi0huGNS/aMSpVieAW
# ilKgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# SIb3DQEJBDEWBBRJIz/Sa/FIkd/l3eACyC6Qdk4OuDBWBgorBgEEAYI3AgEMMUgw
# RqAegBwARwBlAHQALQBVAEMAUABvAG8AbAAuAHAAcwAxoSSAImh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAoauWQpzI
# LUGv8SKoryTFEdlqKaGgV8VWY0zI+38fFyf2rx7PY38DMDDW9LJobkSLacDHcnod
# Iv2ukjFbBokXzfadQSC4Gw8EpB9BLalaBn196R7xJ2z00CKhamcWa8Zr3Ps3h78q
# Mu2dsJWmPF1ibBmxfflxY4+2c7Hs+qVxq+wyM2zq+HjvnXprv8AQN0Nt0Ynxc1WO
# 7TRJvXR0XBFft1A9KS+4ZUC4E/Xm2V1HkWQ1dSiblcaDg/zC1FymzapDG2xnY6YZ
# +KytNN4wo4+r1/RQ3ixcbgHaOEGpMhdglQFCwQwn3kJ88+QIAnZTrMkXFMfLAQvZ
# bORUp6BDuUCWu6GCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAAArOTJIwbLJSPMAAAAAACswCQYFKw4D
# AhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTEzMDIwNTA2MzcyNFowIwYJKoZIhvcNAQkEMRYEFJxdE4sXLC8oSU7K2P8TqY17
# pqmWMA0GCSqGSIb3DQEBBQUABIIBAA9w3Vny00goUkZE9JJ7E9uZpuCdp1Gb2g1D
# rjjzQwZYV8B/DkNGy9Rv1xDa3UIitm0ZPsGujy0XMO6SkEjgVrwYZ1lwAJiftP/t
# WxPaoipFJvIQEkwIAkxLIlD29mv03ZFPxbIqapNb1GcElaBVoOGrz3g8aRVlx3k8
# BLHXNIHLzXVFi5vSpYmdIZm9IT9sQCOyFhsyiK0JwJBCmIjYG10rTp50B9fsBft/
# tXhiU3ceCzi7y/TBEANyg3xtpmWN116eejZxOZEYU8Fkn2v2CzoCcOQJOOBK+swk
# URK0L2UY5+LKIu3+GWTH/V84KaZGOCZpJ+83ZQxAHWjlmEaerzg=
# SIG # End signature block
