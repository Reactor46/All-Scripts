function Get-SPRWebApplicationGeneralSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication
  )
	
  foreach ($webApplication in $SPWebApplication)
  {
  
    $timeZone = Get-SPRTimeZoneById -Id $webApplication.DefaultTimeZone
    $quotaTemplate = Get-SPRWebApplicationDefaultQuotaTemplate -WebApplication $webApplication
    
    $properties = @{
      'DisplayName'        = $webApplication.DisplayName
      'DefaultTimeZone'    = $timeZone
      'DefaultQuotaTemplate' = $quotaTemplate
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}




