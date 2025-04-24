function Get-SPRFarmManagedAccounts
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Username'         = $XXX
      'AutomaticChange'  = $XXX
      'LastChange'       = $XXX
      'NextPasswordChange' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}