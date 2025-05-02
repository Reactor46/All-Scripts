#$Servers = "LASWEB01", "LASWEB02", "LASWEB03", "LASWEB04", "LASWEB05", "LASWEB06", "LASWEB07", "LASWEB08", "LASWEB10", "LASWEB11", "LASWEB12","LASDMZWEB01","LASDMZWEB02","LASDMZWEB03","LASDMZWEB04","LASDMZWEB05","LASDMZWEB06","LASDMZWEB07","LASDMZWEB08","LASDMZWEB09","LASDMZWEB10","LASDMZWEB11","LASDMZWEB12"
$Servers= Get-Content C:\LazyWinAdmin\WebServers\Configs\IISLogs.txt
foreach ($Server in $Servers){

    $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $Server -Filter "DeviceID='D:'" | Select-Object Size,FreeSpace
    $percentUsed = [math]::Round(100 - (($disk.FreeSpace / $disk.Size) * 100))
    Write-Host $Server "D:\ - % Used =>" $percentUsed % |fl


}

