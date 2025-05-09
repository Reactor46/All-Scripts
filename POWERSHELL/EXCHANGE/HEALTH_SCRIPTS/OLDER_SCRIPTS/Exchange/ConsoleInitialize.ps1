# Copyright (c) 2008 Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 
#
# This script is used by Exchange Management Console only to aid filter out based on some conditions.
#    

$global:ConfirmPreference = "None"

# Get logon user identity
# 
# For the case of multi-domain topology which has the same prefix for the domain name, such as abc.com and abc.cn,
# userInfo.Identity.Name is not unique since it only contains the short domain name, for example abc\administrator.
# 
# So EMC uses the unique security identifier userInfo.WindowsIdentity.User to identify the user when the WindowsIdentity is available.
function global:Get-LogonUserIdentity ()
{
	$senderInfo = (Get-Variable PSSenderInfo).Value
    if (($senderInfo -eq $null) -or ($senderInfo.UserInfo -eq $null))
    {
        throw 'Could not retrieve the information of the user making the call'
    }

    $sid = [Microsoft.Exchange.Configuration.Authorization.ExchangeAuthorizationPlugin]::GetExecutingUserSecurityIdentifier($senderInfo.UserInfo, $senderInfo.ConnectionString) 
	
	$sid.ToString()
}

# Get logon user
function global:Get-LogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}	
	PROCESS 
	{
		Set-ADServerSettings -ViewEntireForest:$true
		Get-User (Get-LogonUserIdentity)
	}
}

# Get ManagementRoleAssignment for logon user
function global:Get-ManagementRoleAssignmentForLogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}	
	PROCESS 
	{
		Set-ADServerSettings -ViewEntireForest:$true
		Get-ManagementRoleAssignment -RoleAssignee (Get-LogonUserIdentity)
	}
}

# Get ManagementRole for logon user
function global:Get-ManagementRoleForLogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}	
	PROCESS 
	{
		Set-ADServerSettings -ViewEntireForest:$true
		foreach ($o in Get-ManagementRoleAssignment -RoleAssignee (Get-LogonUserIdentity))
		{
			$o.Role | Get-ManagementRole
		}
	}
}

# Get ManagementScope for logon user
function global:Get-ManagementScopeForLogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}
	PROCESS 
	{
		Set-ADServerSettings -ViewEntireForest:$true
		
		$scopes = @{}
		foreach ($o in Get-ManagementRoleAssignment -RoleAssignee (Get-LogonUserIdentity))
		{
			If ($o.CustomRecipientWriteScope -ne $null -and !$scopes.ContainsKey($o.CustomRecipientWriteScope))
			{
				$scopes.add($o.CustomRecipientWriteScope, $null)
				$o.CustomRecipientWriteScope | Get-ManagementScope -ErrorAction SilentlyContinue
			}

			If ($o.CustomConfigWriteScope -ne $null -and !$scopes.ContainsKey($o.CustomConfigWriteScope ))
			{
				$scopes.add($o.CustomConfigWriteScope , $null)
				$o.CustomConfigWriteScope | Get-ManagementScope -ErrorAction SilentlyContinue
			}
		}
	}
}

# Get Exclusive ManagementScope for logon user
function global:Get-ExclusiveManagementScopeForLogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		Get-ManagementScope -Exclusive:$true
	}
}

# Get ADServerSettings for logon user
# The reason to write this wrapper is that we cannot public the Get-ADServerSettings to any user
# without a role assignment, otherwise it won't pass the check during setting Rbac Scope in task.cs.
function global:Get-ADServerSettingsForLogonUser ()
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}	
	PROCESS 
	{
		Get-ADServerSettings
	}
}


# Set ADServerSettings for logon user
# The reason to write this wrapper is that we cannot public the Get-ADServerSettings to any user
# without a role assignment, otherwise it won't pass the check during setting Rbac Scope in task.cs.
function global:Set-ADServerSettingsForLogonUser ([object]$RunspaceServerSettings)
{
	BEGIN
	{
		set-variable VerbosePreference -value SilentlyContinue
	}	
	PROCESS 
	{
		Set-ADServerSettings -RunspaceServerSettings $RunspaceServerSettings
	}
}

# Filter out object which NameProperty equals an input SearchText
function global:Filter-PropertyStringContains ([string]$Property, [string]$SearchText)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.$Property.ToString().ToUpper().Contains($SearchText.ToUpper())} 		
		}
		while ($false) #connectScope
	}
}

# Filter out object wich $Property not contains $SearchText
function global:Filter-PropertyStringNotContains ([string]$Property, [string]$SearchText)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.$Property.ToString().ToUpper().Contains($SearchText.ToUpper()) -eq $false} 		
		}
		while ($false) #connectScope
	}
}

# Sort the pipeline objects with specified property
function global:Sort-Objects ([string]$Property)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | Sort-Object -Property $Property
		}
		while ($false) #connectScope
	}
}

# Equal to
function global:Filter-PropertyEqualTo ([string]$Property, [object]$Value=$null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.$Property -eq $Value} 		
		}
		while ($false) #connectScope
	}
}

# Not Equal to
function global:Filter-PropertyNotEqualTo ([string]$Property, [object]$Value=$null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.$Property -ne $Value} 		
		}
		while ($false) #connectScope
	}
}

# Equal or Greater Than
function global:Filter-PropertyEqualOrGreaterThan ([string]$Property, [object]$Value=$null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.$Property -ge $Value} 		
		}
		while ($false) #connectScope
	}
}

# Filter out recipients without primary smtp address
function global:Filter-Recipient ()
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.PrimarySmtpAddress.Length -ne 0} 		
		}
		while ($false) #connectScope
	}
}

# Resolve a bunch of objects 
function global:Filter-PropertyInObjects ([string]$ResolveProperty, [string[]]$inputObjects = $null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			if ($inputObjects -eq $null)
			{
				return
			}
			
			if ($inputObjects -contains $_.$ResolveProperty)
			{
				$_
			}
		}
		while ($false) #connectScope
	}
}

# Will be removed with Delegation Feature
function global:Filter-Delegation ()
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{
			# All Delegation feature will be cut soon.
			$_ 
		}
		while ($false) #connectScope
	}
}

# Filter mailboxes which version equal or greater than the specific
function global:Filter-Mailbox ([object]$Value=$null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			$_ | where-object {$_.ExchangeVersion.ToInt64() -ge $Value} 		
		}
		while ($false) #connectScope
	}
}

# Specific Filter for DatabaseMaster Picker to get all Databases or Master of its DatabaseCopies
function global:Filter-DatabaseMasterServer ()
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{				
			foreach ($copy in $_.DatabaseCopies)
			{
				$copy.HostServerName | Get-ExchangeServer
			}

		}
		while ($false) #connectScope
	}
}

# Specific Filter for ServersInSameDag Picker to get servers in the same DAG as $dagMemberServer
function global:Filter-ServersInSameDag ([string]$dagMemberServer)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{		
			if ($_.Servers -Contains $dagMemberServer)
			{
				foreach ($server in $_.Servers)
				{
					if ($server -ne $dagMemberServer)
					{
						$server | Get-ExchangeServer
					}
				}
			}
		}
		while ($false) #connectScope
	}
}

# Get all PublicFolder Installed Exchange Server
function global:Filter-PublicFolderInstalledExchangeServer ([int]$minVersion = 8, [string]$excludeServer = $null)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			if ($_.Server -eq $excludeServer)
			{
				return
			}
			else
			{
				$_.Server | Get-ExchangeServer | where-object {$_.AdminDisplayVersion.Major -ge $minVersion} 		
			}
		}
		while ($false) #connectScope
	}
}

# Filter for ExchangeCertificate picker
function global:Filter-ExchangeCertificate ([object]$isSelfSigned, [object]$hasKeyIdentifier, [object]$privateKeyExportable, [object]$status)
{
    BEGIN
    {
        set-variable VerbosePreference -value Continue
    }
    PROCESS
    {
        :connectScope do
        {
            if (($status -eq $null -or $_.Status -eq $status) -and ($isSelfSigned -eq $null -or $_.IsSelfSigned -eq $isSelfSigned) -and ($hasKeyIdentifier -eq $null -or [string]::IsNullOrEmpty($_.KeyIdentifier) -ne $hasKeyIdentifier) -and ($privateKeyExportable -eq $null -or $_.PrivateKeyExportable -eq $privateKeyExportable))
            {
                $_
            }
        }
        while ($false) #connectScope
    }
}

# A generic ExchangeServer Filter for all ExchangeServer Pickers
function global:Filter-ExchangeServer ([int]$minVersion = 8, [int]$maxVersion = 2147483647, [string[]]$serverRoles, [switch]$includeLegacyServer, [switch]$backendServerOnly, [string[]]$excludedServers, [Microsoft.Exchange.Data.ExchangeBuild]$exactVersion)
{
	BEGIN
	{
		set-variable VerbosePreference -value Continue
	}	
	PROCESS 
	{
		:connectScope do
		{	
			if ($excludedServers -contains $_.Identity)
			{
				return
			}
			
			if ($includeLegacyServer -and $_.AdminDisplayVersion.Major -lt 8)
			{
				# only return backend server
				if ($backenServerOnly -and ($_.ExchangeLegacyServerRole -ne 0))
				{
					return
				}
				
				$_
			}
			else
			{
				$targetServerRole = $false
				foreach($role in $serverRoles)
				{
					if ($_.ServerRole -like '*'+$role+'*')
					{
						$targetServerRole = $true
					}
				}
			
				if ($targetServerRole -and $_.AdminDisplayVersion.Major -ge $minVersion -and $_.AdminDisplayVersion.Major -le $maxVersion)
				{
					if ($exactVersion -ne $null)
					{
						if (($_.AdminDisplayVersion.Major -eq $exactVersion.Major) -and ($_.AdminDisplayVersion.Minor -eq $exactVersion.Minor) -and ($_.AdminDisplayVersion.Build -eq $exactVersion.Build))
						{
							$_
						}
					}
					else
					{
						$_
					}
				}
			}
		}
		while ($false) #connectScope
	}
}
# SIG # Begin signature block
# MIIdqgYJKoZIhvcNAQcCoIIdmzCCHZcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULbZM91bgPH+uwSOWYDXwxYVv
# r86gghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLAwggSsAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBxDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUsV6DANbm8KBVIpq/0XKrlfCO27swZAYKKwYB
# BAGCNwIBDDFWMFSgLIAqAEMAbwBuAHMAbwBsAGUASQBuAGkAdABpAGEAbABpAHoA
# ZQAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAw
# DQYJKoZIhvcNAQEBBQAEggEAKjpY/MDAEdZF5BGmAvrGK9R2ZhlxrYN11S7I0Egm
# wtrImFbp/TGDmM+xrRe/x97i5lic+lN3RomNyMoH0SoARDQXWr3cFEyggndaJty7
# sIJFMthiHh09nKalr+ySWv4TqQ7mBBGt0gAoCEe5DZyLIBA1VGQVyD6K5hELCzNi
# 3Af0HoP0SuyesXBOm1/0a70Em3IaKYcEIGamVT2FrvFaJanKDdASpaQ9UfTbJTo8
# 3a1Zg+lOH5C/xu3Ci8czCu0YQe5yobUBUTsXNgQGV17dSbJ4qlfe4zcQNBacukVy
# UA73Ef055wmvsWfrhqV9rsJj9NzgmqtVAXlp1dIAC/EsgaGCAigwggIkBgkqhkiG
# 9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMA
# AACYBFjLfyMJsJ4AAAAAAJgwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQzNVowIwYJKoZIhvcN
# AQkEMRYEFH/FGSjSiPGu2Xleqo7tAhfoUH17MA0GCSqGSIb3DQEBBQUABIIBAIhk
# +NjTKmzX9ZgvMXRga+2NVx965TsjGVCpqJianeByhzxn3kZo6QXZcdfyuOoU7CpD
# 6puMTJWzRcQNyFym4g0hW4ZzpU+Qr2eYKppbZYKmcczqUM8bQzVHk/aZBF79ZbVl
# dzR67C0QRLnSFBHV4LrpwL5RUiu9DtIg0/FMhMsBbtl1z7qGrzwMlZO55aDCv9Mv
# 7XEGb+NQ0aLudcR01I9FvuXyyZ3Z80b33iaLqchSQGldAoFvj75dvUlK/JA2NCce
# 4jycxxM0b6A++hpvkcz39sYhmTgkUp8Pvedvl6VcSVkZil0TkgBIPfUfmqVdQ78b
# 9yKyV2cyLaWPkt9jkgg=
# SIG # End signature block
