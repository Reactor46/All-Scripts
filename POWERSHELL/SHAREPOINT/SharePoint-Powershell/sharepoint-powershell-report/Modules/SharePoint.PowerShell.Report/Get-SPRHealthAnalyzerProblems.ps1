function Get-SPRHealthAnalyzerProblems
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Category'      = $XXX
      'Title'         = $XXX
      'Severity'      = $XXX
      'FailingServers' = $XXX
      'FailingServices' = $XXX
      'LatModified'   = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}