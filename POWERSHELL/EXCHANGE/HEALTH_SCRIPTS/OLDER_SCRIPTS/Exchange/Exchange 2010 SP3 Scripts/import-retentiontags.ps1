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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUH0QmYATRS+j6lhcTvwoED/yF
# GbagghhqMIIE2jCCA8KgAwIBAgITMwAAARzbbpm3tnP6bwAAAAABHDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM1
# WhcNMjAwMTEwMjEwNzM1WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RDJDRC1FMzEwLTRBRjExJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCxqRuPkgvAvJMVHxyEsWMAs/pxAn3vnvfWrFqQj2NkG9kP
# E3XXn9Xn7n7WsHbuuVdpi4nSyPfLTriA2kzbF+eco/ZTVRbanYk8BXwZGgUzRgF4
# LxQq4INdpNmH2zBti8HK7xURC8HoBB82c5VnZp1AZvgnWRs+6wbzXnauqbwoGuTJ
# XPzaPXivUjL2W+W9G9NMJ5nrmkcNcmq/ncqA88qrofMBqly6y+SL1EdCR0oVYl1A
# ZOgf+ALrh/TMeA1Bld+EFzJa/rEo1QB3IPcwm3xQfW26SYOyQFPIfLjXkBs+VYrc
# S27bByATdjsOJ06krz5tc2fKLv+ao5r1sOIvFDcFAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUb8nAx97t5y1LdYL20QwUPKqBH8UwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAWVKU4uhqdIGVX+vj
# MkduTPqjk59ZxNeOrJX/O7MP5OkObcq6T+vqTyjmeTsiNoO0btyofj9bUJUAic8z
# 10V/rwlvvsYUyzlnTos7+76NU86PoQuMGTLuPfmEAQD4rpUs1kyJchz2m0q7/AbI
# usbsTTLzJ8TW7vyEluJG9LhLAxvAz7dvWdcWQBmh52egoL84XvUq4g0lFNqkiSIV
# 7z7IFsXbvXzhS2NnOLIdpHjGfxhIvRCTFNKCxflV+O8/AqERd6txTeBFpWPRvN0U
# S+GOJvA77FxAvGH2vaH3zQ3WeQxVBAJ6LrUCiKkKm+gJFwE/2ftF5zEMuZS9Zg/F
# EnmzLDCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU3EtCCa65wr6L94Vd0+vGWG7gXCAw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAFIa4ILs25eHLpRAYV2oWOnmEy6s2XJv5z/aa3bBlpk/
# B+wPJFqj6v0F1/f3xyBNCy96E3TsTByqzQxm+54AbhTsD1k/qwquVpiDsmt+aV8Z
# /FygLkcsG5U7dFt1fqGOizIRJMPyEoK1NNxi0BdC/pZuXMo+kAaLNcLQhSZUBxTP
# fdkq9HIi28/n1WhJeVsgndgndwP9YMKfzstYyhCGYyjnu/1FuqeNvK64um+th4xU
# 3gwbSgG2n6e8tQd5rysHGx4650GNTiSmUrEt9P0xvTzBOhFcwkHjFobM7tZcKBp5
# gdZW2ZArCkw+QurFdpLji0Lxdkx5yOHYMbtl5YlFuJmhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# HNtumbe2c/pvAAAAAAEcMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBQDzbhh9qJIpXXFnLvC0nldxsuwETANBgkqhkiG9w0BAQUFAASCAQACMP5q
# bbgF6DtNb6EyczNnng+gvxzPQl+ZUTiXUbq7yJdH3k76UAUDpOKpuzz1IUeI5LFH
# VtgkXSHi2fC8fxzzoh+g32ZBZK8/qUiByh/CA4y7RfHy+J46D5eZwCH8fpPQDS7N
# aPX02c/E3hXjEXIzZyIm7hifg1JepqvHv6u3VnsS4giju1w8mYspfpQNM+quD+PV
# qFZM4TUaiCHuj0b3fueO/I1VUTj8YeD96reSFHxBI/zyA9RKM1KXk42Kz90U2NyL
# 3ND2e8XtdZJlTInq37m/2ZH/MHVKLoKoZdIaYdYYI4n5IuBy8fblxfOXCnJaeV8P
# Y19PmPA2BROE5XEv
# SIG # End signature block
