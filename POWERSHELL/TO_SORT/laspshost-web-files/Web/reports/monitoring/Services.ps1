$runtimeFiles =
@{
 "influxdb" = [PSCustomObject] @{name="InfluxDB run-time files"; id="InfluxDB"; daemonExe="influxd.exe"; ...};
"kapacitor" = [PSCustomObject] @{name="Kapacitor run-time files"; id="Kapacitor"; daemonExe="kapacitord.exe"; ...};
  "grafana" = [PSCustomObject] @{name="Grafana run-time files"; id="Grafana"; daemonExe="grafana-server.exe"; ...};
}

foreach ($key in $($runtimeFiles.keys)) {
    $id = $runtimeFiles[$key].id
    if (!(IsServiceInstalled $key)) {

        # GetCmdLine returns something like "c:\influxdb\influxd.exe -config `"c:\influxdb\custom.conf`""
        $cmdLine = GetCmdLine $key
        $Arguments =  @()
        $Arguments += "install"
        $Arguments += $key
        $Arguments += $cmdLine

        Write-Host "... Trying to install service `"$key`" with: $nssmExe $Arguments ..."
        Start-Process "$nssmExe" -ArgumentList $Arguments -Wait -PassThru -NoNewWindow | Out-Null
        if (IsServiceInstalled $key) {
            Write-Host "Success: installed service: $key"
        } else {
            Write-Host "FAILED: did not install service: $key"
            ExitWithCode $errorCode
        }
    }

    $description = "Provides $id windows service"
    $display = "$id"
    Write-Host "... Trying to configure service `"$key`" ..."
    Start-Process "$nssmExe" -ArgumentList "set $key AppExit Default Restart" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key AppNoConsole 1" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key AppRestartDelay 60000" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key ObjectName LocalSystem" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key Start SERVICE_AUTO_START" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key Type SERVICE_WIN32_OWN_PROCESS" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key DisplayName `"$display`"" -Wait -PassThru -NoNewWindow | Out-Null
    Start-Process "$nssmExe" -ArgumentList "set $key Description `"$description`"" -Wait -PassThru -NoNewWindow | Out-Null
}