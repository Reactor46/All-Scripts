function Get-SPRWebApplicationManagedPaths
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPManagedPath
  )
	
  foreach ($managedPath in $SPManagedPath)
  {
if($managedPath.Name -eq '')
      {
        $name = '/'
      }
      else 
      {
        $name = '/{0}/' -f $managedPath.Name
      }
      
      if ($managedPath.Type.Value -eq 'ExplicitInclusion')
      {
        $prefixType = 'Explicit Inclusion'
      }
      elseif($managedPath.Type.Value -eq 'WildcardInclusion') 
      {
        $prefixType = 'Wildcard Inclusion'
      }
      
    
      $properties = @{
        'WebApplicationUrl' = $managedPath.WebApplicationUrl
        'Name'      = $name
        'PrefixType' = $prefixType
      }

    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}