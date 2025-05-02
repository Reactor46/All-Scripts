function Get-SPRWebApplicationWorkflowSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPWebApplication
  )
	
  foreach ($webApplication in $SPWebApplication)
  {
    $properties = [ordered]@{
      'DisplayName' = $webApplication.DisplayName
      'AllowInternalUsers' = $webApplication.UserDefinedWorkflowsEnabled
      'AllowExternalUsers' = $webApplication.ExternalWorkflowParticipantsEnabled
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}