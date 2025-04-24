function Get-SPRQuotaTemplates
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPQuotaTemplate
  )
   

    foreach ($quotaTemplate in $SPQuotaTemplate){
    
    $properties = [ordered]@{
      'TemplateName'          = $quotaTemplate.Name
      'StorageMaximumLevel'   = $quotaTemplate.StorageMaximumLevel
      'StorageWarningLevel'   = $quotaTemplate.StorageWarningLevel
      'InvitedUserMaximumLevel' = $quotaTemplate.InvitedUserMaximumLevel
      'UserCodeMaximumLevel'  = $quotaTemplate.UserCodeMaximumLevel
      'UserCodeWarningLevel'  = $quotaTemplate.UserCodeWarningLevel
      #'WarningLevelEmail'     = $xxx
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
  
}