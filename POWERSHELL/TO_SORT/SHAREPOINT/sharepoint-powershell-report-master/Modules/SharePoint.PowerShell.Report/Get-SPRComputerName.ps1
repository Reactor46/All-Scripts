function Get-SPRComputerName
{
  param
  (
    [String][Parameter(Mandatory)]
    $Address
  )
	
  try
  {
    $output = Get-SPRComputerNameByHost -Hostname $Address
  }
  catch
  {
    $output = Get-SPRComputerNameByAlias -Alias $Address
  }
	
  Write-Output -InputObject $output
}