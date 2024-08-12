################################################################################
#
#  Import-RetentionTags.ps1
#
#  $Path: File from which Retention Policy Tags and Policies are imported.
#  $Organization: Name of organization. Required only if new location is in  datacenter. 
#  $DomainController: Domain Controller. Optional.
#  $Update: Set true for updating. Default is True.
#  $Delete: Set true for deleting. Default is True.
#  $Confirm: Set true for asking input for update and delete. If it is false, then update 
#			and delete actions are performed based on Update and Delete parameters above.
################################################################################

Param($Path, $Organization, $DomainController, $Update, $Delete, $Confirm)

Import-LocalizedData -BindingVariable MigrateTags_Strings -FileName MigrateRetentionTags.strings.psd1

$parsedConfirm  = $True
$parsedUpdate = $True
$parsedDelete = $True

if ($Confirm -ne $null)
{
	if ($Confirm.ToLower().Equals("false"))
	{
		$parsedConfirm = $False;
	}
}

if ($Update -ne $null)
{
	if ($Update.ToLower().Equals("false"))
	{
		$parsedUpdate = $False;
	}
}

if ($Delete -ne $null)
{
	if ($Delete.ToLower().Equals("false"))
	{
		$parsedDelete = $False
	}
}

function EscapeXml([String]$text)
{
	return $text.replace('&', '&amp;').replace("'", '&apos;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;') 
}

function CreateIdentity([String]$name)
{
	if($DCAdmin)
	{
		return ($Organization +  '\' + $name)
	}
	
	return $name;
}

function RemoveRetentionTags([String[]]$tagNames)
{
 	foreach($tagName in $tagNames)
	{
		$identity = CreateIdentity -name:$tagName
		Remove-RetentionPolicyTag $identity -Confirm:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
	}
}

function RemoveRetentionPolicies([String[]]$policyNames)
{
 	foreach($policyName in $policyNames)
	{
		$identity = CreateIdentity -name:$policyName	
		Remove-RetentionPolicy $identity -Confirm:$false -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
	}
}

function ArrayToString($array)
{
	$isFirst = $true;
	$str = "";
	foreach($i in $array)
	{
		if ($isFirst)
		{
			$isFirst = $false;
			$str = " " + $i;
		}
		else
		{
			$str += ", " + $i;
		}
	}
	
	$str += "."
	return $str
}

function FindInArray([String[]] $array, [String] $str)
{
	foreach($item in $array)
	{
		if ($item -eq $str)
		{
			return $true;
		}
	}
	
	return $false;
}


function CompareArrays([String[]]$a, [String[]]$b)
{
	if ($a -eq $null -and $b -eq $null)
	{
		return $true;
	}
	
	if ($a -eq $null -or $b -eq $null)
	{
		return $false;
	}
	
	if ($a.Length -eq 0 -and $b.Length -eq 0)
	{
		return $true;
	}

	if ($a.Length -eq 0 -or $b.Length -eq 0)
	{
		return $false;
	}
	
	if ($a.Length -ne $b.Length)
	{
		return $false;
	}

	$c = $a | Sort-Object
	$d = $b | Sort-Object

	$enum1 = $c.GetEnumerator()
	$enum2 = $d.GetEnumerator()

	while ($enum1.MoveNext() -and $enum2.MoveNext())
	{
		if($enum1.Current -ne $enum2.Current)
	  	{
			return $false;
		}
	}

	return $true;
}

if(!$Path)
{
	Write-Host $MigrateTags_Strings.SpecifyImportFile  -ForegroundColor:Red
   	exit;
}

$DCAdmin = $false
$OrgParam = Get-ManagementRole -Cmdlet Get-RetentionPolicyTag -CmdletParameters Organization
#	if any management roles have access to Organization paramenter, then this is a DC Admin.
$DCAdmin = !!$OrgParam

if($DCAdmin)
{
    if (!$Organization)
    {
	# Setting $DCadmin to false, so that Organization parameter is not used. 
		$DCAdmin = $false;
    }
}

$GetParams = @{
    ErrorAction = "SilentlyContinue"
	WarningAction = "SilentlyContinue"
}

if($DCAdmin)
{
    $GetParams["Organization"] = $Organization
}

[xml]$tags = Get-Content $Path

if ($tags)
{
    if ($tags.RetentionData.RetentionPolicyTag)
    {
		$duplicateTags = @()
		$readConfirmation = "n"

		foreach($tag in $tags.RetentionData.RetentionPolicyTag)
		{
			[Boolean]$RetentionEnabled = [System.Convert]::ToBoolean($tag.RetentionEnabled)
			[Boolean]$SystemTag = [System.Convert]::ToBoolean($tag.SystemTag)
			[Boolean]$MustDisplayCommentEnabled = [System.Convert]::ToBoolean($tag.MustDisplayCommentEnabled)
	    	
			$RetentionId = $tag.RetentionId
			$TagExists = Get-RetentionPolicyTag @GetParams  | where {$_.RetentionId -eq "$RetentionId" }
			if ($TagExists)
			{
				$tagName = $tag.Name
				
				if ($tagExists.Comment -eq $tag.Comment -and
					$tagExists.Name -eq $tag.Name  -and
					$tagExists.LabelForJournaling -eq $tag.LabelForJournaling -and
					$tagExists.MessageClass -eq $tag.MessageClass -and
					$tagExists.MessageFormatForJournaling -eq $tag.MessageFormatForJournaling -and
					$tagExists.MustDisplayCommentEnabled -eq $MustDisplayCommentEnabled -and
					$tagExists.RetentionAction -eq $tag.RetentionAction -and
					$tagExists.Type -eq $tag.Type -and
					$tagExists.RetentionEnabled -eq $RetentionEnabled -and
					$tagExists.SystemTag -eq $SystemTag -and 
					$tagExists.DomainController -eq $DomainController)
				{
					$commentArray = @()
					if ($tag.LocalizedComment.Comment)
					{
						foreach($comment in $tag.LocalizedComment.Comment)
						{
							$commentArray = $commentArray + $comment
						}
					}
					
					if (CompareArrays -a:$commentArray -b:$tagExists.LocalizedComment)
					{
						$locNameArray = @()
						
						if ($tag.LocalizedRetentionPolicyTagName.LocalizedName)
						{
							foreach($locName in $tag.LocalizedRetentionPolicyTagName.LocalizedName)
							{
								$locNameArray  = $locNameArray  + $locName
							}
						}
						
						if (CompareArrays -a:$locNameArray -b:$tagExists.LocalizedRetentionPolicyTagName)
						{
							if(!$tagExists.AgeLimitForRetention -and !$tag.AgeLimitForRetention)
							{
								continue;
							}
							
							if($tagExists.AgeLimitForRetention -eq $tag.AgeLimitForRetention)
							{
								continue;
							}
						}
					}
				}
				
				$duplicateTags = $duplicateTags + $tagName
			}
		}

		$tagsUpdatedMessage = ''
		$updateTags = $parsedUpdate
		if ($duplicateTags.Length -gt 0)
		{
			$tagsUpdatedMessage = ArrayToString($duplicateTags);
			if ($parsedConfirm -and -$parsedUpdate) 
			{
				do 
				{
					Write-Host ($MigrateTags_Strings.TagsAlreadyExist -f $tagsUpdatedMessage) -ForegroundColor:Yellow
					$readConfirmation = Read-Host
				} while ($readConfirmation.ToLower() -ne "y" -and $readConfirmation.ToLower() -ne "n")

				$updateTags = ($readConfirmation -eq "y")
			}
		}
		
		$tagsAdded = @()
		foreach($tag in $tags.RetentionData.RetentionPolicyTag)
		{
			[Boolean]$RetentionEnabled = [System.Convert]::ToBoolean($tag.RetentionEnabled)
			[Boolean]$SystemTag = [System.Convert]::ToBoolean($tag.SystemTag)
			[Boolean]$MustDisplayCommentEnabled = [System.Convert]::ToBoolean($tag.MustDisplayCommentEnabled)
	    	
			$RetentionId = $tag.RetentionId
			$TagExists = Get-RetentionPolicyTag @GetParams  | where {$_.RetentionId -eq "$RetentionId" }

			if (FindInArray -array:$duplicateTags -str:$tag.Name)
			{
			# Tag exists and different from on file.
				if (!$updateTags)
				{
				# User opted for not changing
					continue;
				}
			}
			elseif ($TagExists)
			{
			# Tag exists and same as on file.
				continue;
			}
			else
			{
				$tagsAdded += $tag.Name
			}
	        
			$newRetentionPolicyTagParameters = @{
				Name = $tag.Name
				Comment = $tag.Comment
				RetentionId = $tag.RetentionId 
				MessageClass = $tag.MessageClass 
				MustDisplayCommentEnabled = $MustDisplayCommentEnabled
				RetentionAction = $tag.RetentionAction 
				Type = $tag.Type 
				RetentionEnabled = $RetentionEnabled 
				SystemTag = $SystemTag 
				ErrorAction = "SilentlyContinue"
				WarningAction = "SilentlyContinue"
			}
	    	

			if($tag.AgeLimitForRetention)
			{
				$newRetentionPolicyTagParameters["AgeLimitForRetention"] = $tag.AgeLimitForRetention
			}

			if ($DomainController)
			{
				$newRetentionPolicyTagParameters["DomainController"] = $DomainController
			}

			if($DCAdmin)
			{
				$newRetentionPolicyTagParameters["Organization"] = $Organization
			}

			$commentArray = @()
			if ($tag.LocalizedComment.Comment)
			{
				foreach($comment in $tag.LocalizedComment.Comment)
				{
					$commentArray = $commentArray + $comment
				}
			}
			$newRetentionPolicyTagParameters["LocalizedComment"] = $commentArray		

			$locNameArray = @()
			if ($tag.LocalizedRetentionPolicyTagName.LocalizedName)
			{
				foreach($locName in $tag.LocalizedRetentionPolicyTagName.LocalizedName)
				{
					$locNameArray  = $locNameArray  + $locName
				}
			}	    		
			$newRetentionPolicyTagParameters["LocalizedRetentionPolicyTagName"] = $locNameArray 
			
			if ($TagExists)
			{
				$GetParams.Remove("Identity")
				$newRetentionPolicyTagParameters.Remove("Type")
				$newRetentionPolicyTagParameters.Remove("Organization")
				$newRetentionPolicyTagParameters.Remove("RetentionId")
				$RetentionId = $tag.RetentionId
				Get-RetentionPolicyTag @GetParams  | where {$_.RetentionId -eq "$RetentionId" } | Set-RetentionPolicyTag @newRetentionPolicyTagParameters -Confirm:$false
			}
			else
			{
				New-RetentionPolicyTag @newRetentionPolicyTagParameters
			}
		}

		if($tagsAdded -ne $null -and $tagsAdded.Length -gt 0)
		{
			$message = ArrayToString -array:$tagsAdded
			Write-Host ($MigrateTags_Strings.TagsCreated  -f $message) -ForegroundColor:Yellow
		}
		
		if ($updateTags)
		{
			if($tagsUpdatedMessage)
			{
				Write-Host ($MigrateTags_Strings.TagsUpdated  -f $tagsUpdatedMessage) -ForegroundColor:Yellow
			}	
		}
    }
	
	$GetParams.Remove("Identity")
	$destinationTags = Get-RetentionPolicyTag @GetParams
	$tagsToDelete = @();	
	if(!!$destinationTags)
	{
		foreach ($destinationTag in $destinationTags)
		{
			$tagRetentionId = $destinationTag.RetentionId
			$tagPath = "/RetentionData/RetentionPolicyTag[RetentionId='$tagRetentionId']"
			$t = $tags.SelectNodes($tagPath);
			if ($t.Count -eq 0)
			{
				$tagsToDelete += $destinationTag.Name;
			}
		}
	}
	
	$deleteTags = $parsedDelete
	if ($tagsToDelete.Length -gt 0)
	{
		if ($parsedConfirm -and $parsedDelete)
		{
	
			$tagMessage = ArrayToString($tagsToDelete);
			do 
			{
				Write-Host ($MigrateTags_Strings.TagsToBeDeleted -f $tagMessage) -ForegroundColor:Yellow
				$readConfirmation = Read-Host
 			 } while ($readConfirmation.ToLower() -ne "y" -and $readConfirmation.ToLower() -ne "n")

			$deleteTags = $readConfirmation -eq "y"
		}
		 
		if ($deleteTags)
		{
			RemoveRetentionTags -tagNames:$tagsToDelete 
			Write-Host ($MigrateTags_Strings.TagsDeleted -f $tagMessage) -ForegroundColor:Yellow
		}
	}

    if ($tags.RetentionData.RetentionPolicy)
    {
		$duplicatePolicies = @()
		$readConfirmation = "n"

		foreach($policy in $tags.RetentionData.RetentionPolicy)
		{
			$RetentionId = $policy.RetentionId
			$policyExists = Get-RetentionPolicy @GetParams | where {$_.RetentionId -eq "$RetentionId" }

			if ($policyExists)
			{
				$policyName = $policy.Name
				
				if ($policyExists.Name -eq $policy.Name  -and
					$policyExists.DomainController -eq $DomainController)
				{
					$tagArray = @()
					if ($policy.RetentionPolicyTagLinks.TagLink)
					{
						foreach($tagLink in $policy.RetentionPolicyTagLinks.TagLink)
						{
							$tagArray = $tagArray + $tagLink
						}
					}
	
					$ADPolicyTagLinks = @()	
			        foreach($tagLink in $policyExists.RetentionPolicyTagLinks)
					{
						$ADPolicyTagLinks += $tagLink.Name;
					}

					if (CompareArrays -a:$tagArray -b:$ADPolicyTagLinks)
					{
						continue;
					}
				}
				
				$duplicatePolicies= $duplicatePolicies+ $policyName 
			}

		}

		$policiesUpdatedMessage = ''
		$updatePolicies = $parsedUpdate
		if ($duplicatePolicies.Length -gt 0)
		{
			if ($parsedConfirm -and $parsedUpdate)
			{
				$policiesUpdatedMessage = ArrayToString($duplicatePolicies);
				do 
				{
					Write-Host ($MigrateTags_Strings.PoliciesAlreadyExist -f $policiesUpdatedMessage) -ForegroundColor:Yellow
					$readConfirmation = Read-Host
				} while ($readConfirmation.ToLower() -ne "y" -and $readConfirmation.ToLower() -ne "n")

				$updatePolicies = ($readConfirmation -eq "y")
			}
		}

		$policiesAdded = @()
		foreach($policy in $tags.RetentionData.RetentionPolicy)
		{
			$RetentionId = $policy.RetentionId
			$policyExists = Get-RetentionPolicy @GetParams | where {$_.RetentionId -eq "$RetentionId" }
			if (FindInArray -array:$duplicatePolicies -str:$policy.Name)
			{
			# Policy exists but different from on file.
				if (!$updatePolicies)
				{
				# User opted for not updating.
					continue;
				}
			}
			elseif ($PolicyExists)
			{
			# Policy exists but is not different from in file.
				continue;
			}
			else
			{
			
				$policiesAdded += $policy.Name
			}

			$newRetentionPolicyParameters = @{
				Name = $policy.Name
				RetentionId = $policy.RetentionId 
				ErrorAction = "SilentlyContinue"
				WarningAction = "SilentlyContinue"
			}
	        
			$tagArray = @()
			foreach($tagLink in $policy.RetentionPolicyTagLinks.TagLink)
			{
				$tagArray = $tagArray + $tagLink
			}
	        
			$newRetentionPolicyParameters["RetentionPolicyTagLinks"] = $tagArray
	                    
			if ($DomainController)
			{
				$newRetentionPolicyParameters["DomainController"] = $DomainController
			}

			if($DCAdmin)
			{
				$newRetentionPolicyParameters["Organization"] = $Organization
			}

			if ($PolicyExists)
			{
				$GetParams.Remove("Identity")
				$newRetentionPolicyParameters.Remove("Organization")
				$newRetentionPolicyParameters.Remove("RetentionId")
				$RetentionId = $policy.RetentionId
				Get-RetentionPolicy @GetParams  | where {$_.RetentionId -eq "$RetentionId" } | Set-RetentionPolicy @newRetentionPolicyParameters -Confirm:$false
			}
			else
			{
				New-RetentionPolicy @newRetentionPolicyParameters
			}
		}

		if($policiesAdded.Length -gt 0)
		{
			$message = ArrayToString -array:$policiesAdded
			Write-Host ($MigrateTags_Strings.PoliciesCreated -f $message) -ForegroundColor:Yellow
		}
		
		if ($updatePolicies)
		{
			if ($policiesUpdatedMessage)
			{
				Write-Host ($MigrateTags_Strings.PoliciesUpdated -f $policiesUpdatedMessage) -ForegroundColor:Yellow
			}
		}
	}
	
	$GetParams.Remove("Identity")
	$destinationPolicies = Get-RetentionPolicy @GetParams | where {$_.Name -ne "ArbitrationMailbox"}
	$policiesToDelete = @();	
	if(!!$destinationPolicies)
	{
		foreach ($destinationPolicy in $destinationPolicies)
		{
			$policyRetentionId = $destinationPolicy.RetentionId;
			$policyPath = "/RetentionData/RetentionPolicy[RetentionId='$policyRetentionId']"
			
			$p = $tags.SelectNodes($policyPath);
			if ($p.Count -eq 0)
			{
				$policiesToDelete += $destinationPolicy.Name;
			}
		}
	}
	
	if ($policiesToDelete.Length -gt 0)
	{
		$tagMessage = ArrayToString($policiesToDelete);
		if ($parsedConfirm -and $parsedDelete)
		{
			do 
			{
				Write-Host ($MigrateTags_Strings.PoliciesToBeDeleted -f $tagMessage) -ForegroundColor:Yellow
				$readConfirmation = Read-Host
			} while ($readConfirmation.ToLower() -ne "y" -and $readConfirmation.ToLower() -ne "n")

			$parsedDelete = ($readConfirmation -eq "y")
		}
		 
		if ($parsedDelete)
		{
			RemoveRetentionPolicies -policyNames:$policiesToDelete 
			Write-Host ($MigrateTags_Strings.PoliciesDeleted -f $tagMessage) -ForegroundColor:Yellow
		}
	}
}

# SIG # Begin signature block
# MIIdsAYJKoZIhvcNAQcCoIIdoTCCHZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3U+9a2GnNxKnszWdy+7Sl467
# PZygghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBLYwggSyAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCByjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUGT3m6eP6sIb7mIQxcatXkFQRGOIwagYKKwYB
# BAGCNwIBDDFcMFqgMoAwAEkAbQBwAG8AcgB0AC0AUgBlAHQAZQBuAHQAaQBvAG4A
# VABhAGcAcwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNo
# YW5nZSAwDQYJKoZIhvcNAQEBBQAEggEASJ4SqcOOCHIVLMCZTEyYFa9OL6Kt87Na
# ZfoqzQfty74E0m8QW0OXM1WuIqUR5JDpn/W4VSdZr/JCKMkj0InMF0xDnKEBTnQ/
# 3oENLEhOuuFbd2y0+hlGaiiwZ7CAP14nWOhAygEw6kHxBS8mZV7QWcKAC0rM0CcS
# 19SJ1ENCF5GCa/n0OMXnQP13uU7/zAlF13y8a3unlh9e7nEG7Z11+h5Pf7ghIeC3
# hJWaylWtoe8PR4FILKYG355QpZndfSoUMb89XqcSOGbqu5xU4qy0hnOMkK8JpxYx
# clein2msd7ru+Pnm4w5PUYWOssA7+gp8aNajXPvhwfuhB9XxCoeEf6GCAigwggIk
# BgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0ECEzMAAACampsWwoPa1cIAAAAAAJowCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQyM1owIwYJ
# KoZIhvcNAQkEMRYEFMGa3TSbgatV/8JaVeOI4iTRd4AkMA0GCSqGSIb3DQEBBQUA
# BIIBAHbXgEO4x6MHn7rEN/woY6ILyJhDus07stnAFZGMWZ2fu1XvnhnMcjpLZA+o
# /OM6nLob9Rsaip1r4UmzErRCbG5QsQQ87SuVfJwog0LkfVRGPPdDtKMMLspo0Uzr
# I4gll6EAVA8HMTIjN+TXbxL8gKiHZcBoTGhFCbqsiPgMTBFu7i3s2HNZAevM070c
# y1TTWc/aEPZCNSubpejF3Tds8tjtj8RhfyeMla+M5Q7+uMVqq4ETCLwXWaHl0O+K
# MIUlh4FLpVbqsvAmqu53ZXJ4lP3gy900DyM6gIV1LsH1oHNYF59Sv5ROaDETOvzk
# bj6iMRNQrkrlMwbGUk7b4hSTzIw=
# SIG # End signature block
