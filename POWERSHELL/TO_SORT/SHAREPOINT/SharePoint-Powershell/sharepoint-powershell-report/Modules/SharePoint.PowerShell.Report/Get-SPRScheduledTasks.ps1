function Get-SPRScheduledTasks
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$ScheduledTasks
  )
  
  foreach ($t in $ScheduledTasks)
  {
	#this check is added so that we do not display the default microsoft windows scheduled tasks
	$path_var = $t.taskpath
	if ($path_var) {
		if (-Not ($t.taskpath.Contains('Microsoft\Windows'))) {
			$properties = [ordered]@{
			  'TaskName' = $t.TaskName
			  'State'     = $t.State
			  'Description'     = $t.Description
			  'Source'     = $t.Source
			  'URI'     = $t.URI
			  'Version'     = $t.Version
			}
			$output = New-Object -TypeName PSObject -Property $properties
			Write-Output -InputObject $output
		}
	}
  }
}