function Get-SPRUserProfuleProperties
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPUserProfileManager
  )
	
  $userProfileProperties = $SPUserProfileManager.Properties
  $output = @()
  foreach ($userProfileProperty in $userProfileProperties)
  {
    $properties = [ordered]@{
      'DisplayName'      = $userProfileProperty.Name
      'PropertyType'     = $userProfileProperty.Type
      'DefaultPolicy'    = $userProfileProperty.DefaultPrivacy
      'PrivacyPolicy'    = $userProfileProperty.PrivacyPolicy
      'UserOverridePolicy' = $userProfileProperty.UserOverridePrivacy
      'Replicable'       = $userProfileProperty.IsReplicable
      'Multivalue'       = $userProfileProperty.IsMultivalued
      'Alias'            = $userProfileProperty.IsAlias
    }
    $entry = New-Object -TypeName PSObject -Property $properties
    $output += $entry
  }
  Write-Output -InputObject $output
}
