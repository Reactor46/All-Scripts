# RestartServices.ps1

# Define the services you want to restart
$services = @('zookeeper-1', 'solr-8983')

foreach ($service in $services) {
    try {
        Write-Host "Restarting service: $service"
        Restart-Service -Name $service -Force -ErrorAction Stop
        Write-Host "$service restarted successfully."
    } catch {
        Write-Host "Failed to restart $service: $_"
    }
}