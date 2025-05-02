function Get-SPRAudiences
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPServiceApplicationUserProfiles
  )
	
  $output = @()
  foreach ($userProfileServiceApplication in $SPServiceApplicationUserProfiles)
  {
    foreach ($audience in $userProfileServiceApplication.AudienceManager.Audiences)
    {
      $properties = [ordered]@{
        'ServiceApplication' = $userProfileServiceApplication.Name
        'Name'             = $audience.AudienceName
        'Owner'            = $audience.OwnerAccountName
        'SatisfyAllRules'  = $audience.AudienceRules
        'LastCompiled'     = $audience.LastCompilation
        'Members'          = $audience.MemberShipCount
      }
      $entry = New-Object -TypeName PSObject -Property $properties
      $output += $entry
    }
    Write-Output -InputObject $output
  }
}
