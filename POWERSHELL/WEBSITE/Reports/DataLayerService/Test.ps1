param([string] $Servers)
    Foreach ($server in (get-content $servers)){
      $Uptime = Get-CimInstance Win32_Process -ComputerName $server -Filter "name = 'DataLayerService.exe'" 
      #$DateCalc = [System.Management.ManagementDateTimeConverter]::ToDateTime($Uptime.CreationDate)
      $memuse = [String]([int]($Uptime.WorkingSetSize/1MB))+" MB"
      $CalcRun = New-TimeSpan -Start ($Uptime.CreationDate) -End (Get-Date)
      New-Object PsObject -Property @{Days = $CalcRun.Days; Hours = $CalcRun.Hours; Minutes = $CalcRun.Minutes;Server = $server;Memory = $memuse} 
      } 
      