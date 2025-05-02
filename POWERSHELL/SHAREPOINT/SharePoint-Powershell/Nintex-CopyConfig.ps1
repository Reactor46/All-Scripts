Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop  


$webapp = Get-SPWebApplication -identity https://apps.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 

$webapp = Get-SPWebApplication -identity https://depts.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 

$webapp = Get-SPWebApplication -identity https://teams.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 

$webapp = Get-SPWebApplication -identity https://pulse.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 

$webapp = Get-SPWebApplication -identity https://wiki.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 

$webapp = Get-SPWebApplication -identity https://eflipchart.kscpulse.com/ 
$webapp.UpdateWorkflowConfigurationSettings() 