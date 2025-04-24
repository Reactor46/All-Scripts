function Get-SPRFarmAdministrators
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPFarmAdmins
  )
	
  foreach ($SPFarmAdmin in $SPFarmAdmins)
  {
    $properties = [ordered]@{
      'Login' = $SPFarmAdmin.UserLogin
      'DisplayName'     = $SPFarmAdmin.DisplayName
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}