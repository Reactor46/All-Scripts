function Get-SPRWebApplicationList
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication
  )
	
  foreach ($webApplication in $SPWebApplication)
  {
    $properties = [ordered]@{
      'WebApplication' = $webApplication.DisplayName
      'Url'          = $webApplication.Url
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}