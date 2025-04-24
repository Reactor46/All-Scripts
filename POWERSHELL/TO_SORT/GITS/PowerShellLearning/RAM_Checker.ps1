# Get compute resource information, included but not limited to RAM
$ram_info = Get-CimInstance -ClassName Win32_ComputerSystem

# As the above gives you your RAM in bytes, we can make it more readable by converting it to GB:
$total_memory_gb = [math]::Round($ram_info.TotalPhysicalMemory / 1GB, 2)

# Display the RAM in GB and provide different messages based on different levels
if ($total_memory_gb -lt 4) {
    Write-Host "Your total physical memory is low at $total_memory_gb GB. Consider upgrading for better performance."
} elseif ($total_memory_gb -lt 8) {
    Write-Host "Your total physical memory is moderate at $total_memory_gb GB. It's suitable for most tasks."
} else {
    Write-Host "Your total physical memory is high at $total_memory_gb GB. You have ample resources for demanding applications."
}
