function Get-SPRWebApplicationDefaultQuotaTemplate
{
  param
  (
    [Parameter(Mandatory)]
    $WebApplication
  )
  
  $quotaTemplateName = $WebApplication.DefaultQuotaTemplate
  
  if ($quotaTemplateName -eq '')
  {
    $output = 'No Quota'
  }
  else 
  {
    $output = $quotaTemplateName
  }
  
  Write-Output -InputObject $output
}