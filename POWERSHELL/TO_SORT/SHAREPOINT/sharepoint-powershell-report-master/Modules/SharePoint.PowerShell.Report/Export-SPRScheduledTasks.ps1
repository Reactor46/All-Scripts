function Export-SPRScheduledTasks {
  param(
    [string]$path
  )
  $file = '{0}\SPTasks.xml' -f $Path
  Start-Job -ScriptBlock {
    $selectedTasks = @()
    $ScheduledTasks = Get-ScheduledTask
	$ScheduledTasks | Export-Clixml -Path $args[0] 
  } -ArgumentList $file 
}