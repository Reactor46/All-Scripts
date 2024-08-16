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
				LabelForJournaling = $tag.LabelForJournaling
				MessageClass = $tag.MessageClass 
				MessageFormatForJournaling = $tag.MessageFormatForJournaling 
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
# MIIadgYJKoZIhvcNAQcCoIIaZzCCGmMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUH0QmYATRS+j6lhcTvwoED/yF
# GbagghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggSxMIIErQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHKMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTcS0IJrrnCvov3hV3T68ZYbuBcIDBqBgorBgEEAYI3AgEMMVww
# WqAygDAASQBtAHAAbwByAHQALQBSAGUAdABlAG4AdABpAG8AbgBUAGEAZwBzAC4A
# cABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkq
# hkiG9w0BAQEFAASCAQCqgFNXU+2BZOg0g1E2EJ34P9FXbSlfuVBslwIVhX33FThV
# KM6vIbZqNkJDtiGncM3ZUJB614lyf12j/4WwomDyTMwRucTEMZXInQCm0f4/XW+T
# 92/kUL7GcIWNt2PBW213wuAjeduXmnnu1xINY5J5nOqc/sYMX04HTGvmY3s2bxw7
# Za2v8CTnb8Y8tGJoRkn8jm7umpgobh48oE+Bn63kayeHM2BkMdJrvrSWMH0OJBTw
# x/Pn//mIx07qsA4d7LqW4usSxPnPkwNiCX9njVdjbMzyPSZfg6Cxx2NMnFfp0oq+
# 8IT7aTk194nrf+xgqSnmZH652QFDsHuKf7kteTC3oYICKDCCAiQGCSqGSIb3DQEJ
# BjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAACs5
# MkjBsslI8wAAAAAAKzAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTMwMjA1MDYzNzIzWjAjBgkqhkiG9w0BCQQx
# FgQUeiL/sgXGsbBf9ECk6iuEY/9kzCMwDQYJKoZIhvcNAQEFBQAEggEAJl2vaezl
# hiVi4GROMWnBNvLp6JmYA2iPFXRym/HM6BveQZnjLuxDTRcZkTcPzgTlb7trWT3n
# aXsxPXxQ7DhYTfEqekHtUDpYFJvAovjo1UTEuOLXyCwJfeMeehbtoWhBUY4EbLXb
# U0iwrhblk8RejdvDJzYuAcK2PZ+9KR5Px5ndfDGRgy6jRcTeE8+8W6O7qXztxSpH
# 1ZPU3Dx8MuVPX/tnZ9mxk6NYVP7kI9z3G/iHj9ngc34RDu6T8t2Y35F/exDL3KZv
# eemGJSuk9kSmDGaUy0i0ZNr1riEfyEI/U1HqPof3JtoP5u1N/fh9Nwa1BoN3CV5V
# WxhyWSsa7szQ4g==
# SIG # End signature block
