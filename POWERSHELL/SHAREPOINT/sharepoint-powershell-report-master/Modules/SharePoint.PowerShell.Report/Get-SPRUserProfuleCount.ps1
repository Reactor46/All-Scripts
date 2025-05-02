function Get-SPRUserProfuleCount
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
      'ServiceApplication' = $userProfileServiceApplication.Name
      'UserProfileCount' = $SPUserProfileManager.Count
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}
