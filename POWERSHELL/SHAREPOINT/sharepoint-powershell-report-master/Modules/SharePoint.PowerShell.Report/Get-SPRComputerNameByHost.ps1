function Get-SPRComputerNameByHost
{
  param
  (
    [String][Parameter(Mandatory)]
    $Hostname
  )
	
  $output = ([Net.Dns]::GetHostByName($Hostname)).Hostname
	
  Write-Output -InputObject $output
}