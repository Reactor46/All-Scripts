function Get-SPRWebApplicationFeatures
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPFeature,
    [Parameter(Mandatory = $true)]
    [object[]]$SPwebApplication
  )
	
  $features = $SPFeature | Where-Object { $_.Scope.Value -eq 'WebApplication' }
  
  foreach ($webApplication in $SPwebApplication)
  {
    foreach ($feature in $features)
    {
      $title = $feature.Name
      $isActive = ($webApplication.Features | Where-Object {$_.DefinitionId -eq $feature.Id}) -ne $null
      
      $properties = @{
        'WebApplication'      = $webApplication.DisplayName
        'FeatureName'      = $feature.DisplayNameEN
        'CompatibilityLevel' = $feature.CompatibilityLevel
        'IsActive'         = $isActive
      }
    
      $output = New-Object -TypeName PSObject -Property $properties
      
      Write-Output -InputObject $output
    }
  }
}