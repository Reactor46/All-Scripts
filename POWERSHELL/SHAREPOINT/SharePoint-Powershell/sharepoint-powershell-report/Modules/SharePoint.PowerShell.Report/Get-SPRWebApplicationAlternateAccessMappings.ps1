function Get-SPRWebApplicationAlternateAccessMappings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPAlternateUrl
  )
	
  foreach ($alternateUrl in $SPAlternateUrl)
  {
    $properties = [ordered]@{
      'DisplayName' = $alternateUrl.WebApplicationDisplayName
      'InternalUrl' = $alternateUrl.IncomingUrl
      'Zone'      = $alternateUrl.Zone
      'Url'       = $alternateUrl.PublicUrl
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}
