function Get-SPRDataRetrievalServices
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplications
  )
	
  foreach ($WebApp in $SPWebApplications)
  {
    $properties = [ordered]@{
      'WebApplication'                 = $WebApp.DisplayName
      'InheritTheGlobalSettings'       = $WebApp.InheritDataRetrievalSettings
      'EnableTheseDataRetrievalServices' = $WebApp.DataRetrievalProvider.Enabled
      'EnableUpdateQuerySupport'       = $WebApp.DataRetrievalProvider.UpdateAllowed
      'EnableTheseDataSourceControles' = $WebApp.DataRetrievalProvider.OleDbProviders
      'ResponseSizeLimit'              = $WebApp.DataRetrievalProvider.MaximumResponseSize
      'RequestTimeOut'                 = $WebApp.DataRetrievalProvider.RequestTimeOut
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}