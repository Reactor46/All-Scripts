function Get-SPRFarmPassphrase
{
    $properties = [ordered]@{
	#TODO: find a way to get the passphrase
      'Passphrase'    = 'Get the passphrase'
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
}