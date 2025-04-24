function Get-SPREventThrottling
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPLogLevels
  )
	
  foreach ($SPLogLevel in $SPLogLevels)
  {
    $properties = [ordered]@{
      'Name'        = $SPLogLevel.Name
      'EventSeverity' = $SPLogLevel.EventSeverity
      'TraceSeverity' = $SPLogLevel.TraceSeverity
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}