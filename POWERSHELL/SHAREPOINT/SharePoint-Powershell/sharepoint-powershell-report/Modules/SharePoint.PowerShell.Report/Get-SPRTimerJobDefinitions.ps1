function Get-SPRTimerJobDefinitions
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Name'            = $XXX
      'JobTitle'        = $XXX
      'LastRunTime'     = $XXX
      'StartTime'       = $XXX
      'RecurringSchedule' = $XXX
      'Enabled'         = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}