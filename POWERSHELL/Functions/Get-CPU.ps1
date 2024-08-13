function Get-CPU
{
    $CPUPercent = @{
        Name = 'CPUPercent'
        Expression = {
            $TotalSec = (New-TimeSpan -Start $_.StartTime).TotalSeconds
            [Math]::Round( ($_.CPU * 100 / $TotalSec), 2)
        }
    }
 
    Get-Process | 
    Select-Object -Property Name, $CPUPercent, Description |
    Sort-Object -Property CPUPercent -Descending |
    Select-Object -First 5
}