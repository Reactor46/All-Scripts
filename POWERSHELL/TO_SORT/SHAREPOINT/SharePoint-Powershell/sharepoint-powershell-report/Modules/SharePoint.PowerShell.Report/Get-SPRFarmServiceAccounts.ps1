function Get-SPRFarmServiceAccounts
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Component' = $XXX
      'Account' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}