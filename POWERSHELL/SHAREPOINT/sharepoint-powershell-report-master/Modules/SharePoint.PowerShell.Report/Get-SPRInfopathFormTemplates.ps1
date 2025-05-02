function Get-SPRInfopathFormTemplates
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Name'          = $XXX
      'Version'       = $XXX
      'Modified'      = $XXX
      'Category'      = $XXX
      'WorkflowEnabled' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}