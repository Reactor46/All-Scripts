function Get-SPRMySiteSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPServiceApplicationUserProfiles,
    [Parameter(Mandatory = $true)]
    [object[]]$SPUserProfileManager
  )

  foreach ($userProfileServiceApplication in $SPServiceApplicationUserProfiles)
  {
    $properties = [ordered]@{
      'ServiceApplication'    = $userProfileServiceApplication.DisplayName
      'Multilingual'          = $SPUserProfileManager.IsPersonalSiteMultipleLanguage
      'ReadPermissionLevel'   = $SPUserProfileManager.PersonalSiteReaders
      'SenderEmailAddress'    = $SPUserProfileManager.MySiteEmailSenderName
      'EnableNetbiosdomainName' = $userProfileServiceApplication.NetBIOSDomainNamesEnabled
      'MySiteManagedPath'     = $SPUserProfileManager.PersonalSiteInclusion
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}
