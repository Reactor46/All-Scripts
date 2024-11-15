﻿  $ComputerName = Get-Content -Path C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Windows7.txt
  if(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet){
        $explorer = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ?{$_.name -eq "explorer.exe"}
        
    if($explorer){
      $Hive = [long]$HIVE_HKU = 2147483651
      $sid = ($explorer.GetOwnerSid()).sid
      $owner  = $explorer.GetOwner()
      $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object {$_.Name -eq "StdRegProv"}
      $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
      
      #If the SID network has mapped drives iterate and report on said drives
      if($DriveList.sNames.count -gt 0){
        "$($owner.Domain)\$($owner.user) on $($ComputerName)"
        foreach($drive in $DriveList.sNames){
          "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
        }
      }else{"No mapped drives on $($ComputerName)"}
    }else{"explorer.exe not running on $($ComputerName)"}
  }else{"Can't connect to $($ComputerName)"}

Out-File -FilePath "C:\LazyWinAdmin\Logs\Contoso.corp\Computers\Mapped\" + $owner + ".txt"