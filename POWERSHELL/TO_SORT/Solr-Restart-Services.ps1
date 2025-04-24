$servers = @("FBV-SCSLR10-P01", "FBV-SCSLR10-P02", "FBV-SCSLR10-P03")
$scriptPathLocal = "D:\scripts\RestartServices.ps1"  # Path on your local machine
$scriptPathRemote = "C:\Scripts\RestartServices.ps1"  # Target path on the remote servers

# Script to restart services (content of RestartServices.ps1)
$restartScriptContent = @'
# RestartServices.ps1
$services = @('zookeeper-1', 'solr-8983')

foreach ($service in $services) {
    try {
        Write-Host "Restarting service: $service"
        Restart-Service -Name $service -Force -ErrorAction Stop
        Write-Host "$service restarted successfully."
    } catch {
        Write-Host "Failed to restart $service : $_"
    }
}
'@

# Time offsets in minutes
$times = @("00:00", "00:10", "00:20")

# Create task on each server
foreach ($i in 0..2) {
    $server = $servers[$i]
    $taskTime = $times[$i]

    # Ensure C:\Scripts directory exists on the remote server
    Invoke-Command -ComputerName $server -ScriptBlock {
        $dir = "C:\Scripts"
        if (-not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory
            Write-Host "Directory created : $dir"
        } else {
            Write-Host "Directory already exists : $dir"
        }
    }

    # Copy the RestartServices.ps1 script to the remote server
    Invoke-Command -ComputerName $server -ScriptBlock {
        param($scriptContent, $scriptPath)
        
        $scriptContent | Set-Content -Path $scriptPath
        Write-Host "Script copied to $scriptPath"
    } -ArgumentList $restartScriptContent, $scriptPathRemote

    # Create the scheduled task on the remote server
    Invoke-Command -ComputerName $server -ScriptBlock {
        param($taskTime, $scriptPath)

        # Task action to run the PowerShell script
        $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File $scriptPath"

        # Trigger: Daily at the specified time
        $taskTrigger = New-ScheduledTaskTrigger -Daily -At $taskTime

        # Task settings: Allow it to run on battery power
        #$taskSettings = New-ScheduledTaskSettingsSet
        
        # Register the scheduled task
        #  Register-ScheduledTask -TaskName "StartupScript1" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest –Force
        $taskName = "Restart_Zk_Solr_$taskTime"
        Register-ScheduledTask -TaskName $taskName -Trigger $taskTrigger -User "NT AUTHORITY\SYSTEM" -Action $taskAction -RunLevel Highest  -Force
       

        Write-Host "Scheduled task created: $taskName"
    } -ArgumentList $taskTime, $scriptPathRemote
}
