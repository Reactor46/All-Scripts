function Get-SPRSqlServerProcessorAffinity
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server'                         = $XXX
      'Processor'                      = $XXX
      'ProcessorAffinityRunningValue'  = $XXX
      'ProcessorAffinityConfiguredValue' = $XXX
      'IOAffinityRunningValue'         = $XXX
      'IOAffinityConfiguredValue'      = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}