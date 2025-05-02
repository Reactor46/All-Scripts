while ($true) {
    # Capture the CPU usage of dsa.exe before the 1-second interval
    $before = Get-Process -Name "dsa" | Select-Object Name, CPU

    # Wait for 1 second to measure the CPU usage over a short period
    Start-Sleep -Seconds 1

    # Capture the CPU usage of dsa.exe after the 1-second interval
    $after = Get-Process -Name "dsa" | Select-Object Name, CPU

    # Calculate the CPU usage delta for the process
    $cpuBefore = $before.CPU
    $cpuAfter = $after.CPU
    $cpuDiff = $cpuAfter - $cpuBefore

    # Display the CPU usage for dsa.exe
    if ($cpuDiff -gt 0) {
        Write-Host "dsa.exe CPU Usage: $cpuDiff seconds"
    }
    else {
        Write-Host "dsa.exe is idle or using no CPU in the last second."
    }

    # Optional: add a pause before the next iteration (e.g., 1 second)
    Start-Sleep -Seconds 1
}