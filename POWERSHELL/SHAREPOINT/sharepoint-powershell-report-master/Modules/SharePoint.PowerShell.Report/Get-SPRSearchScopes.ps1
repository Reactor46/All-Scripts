function Get-SPRSearchScopes
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'ServiceApplication'      = $XXX
      'DisplayName'             = $XXX
      'LastModifiedBy'          = $XXX
      'AlternateResultsPage'    = $XXX
      'DefaultSearchResults'    = $XXX
      'DifferentPageForSearching' = $XXX
      'RuleConunt'              = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}