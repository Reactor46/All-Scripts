param( 
    [String] $FolderPath = 'D:\Ping-Temp',  #Where you want the log files kept.  it will create the directory if it dosn't exist.
    [Int32] $HoursToRun = 8,  #How Many hours do you want this to run
    [System.Array] $Computers = ('acadia','cypress','denali','grand-teton','great-lakes','hudson_river','lake_tahoe','large-exec-conf','mount_ranier','MISSISSIPPI','mammoth_cave','mount_mitchell','pacific_ocean','pikes-peak','sagebrush','sierra-nevada','sequoia','saguaro','vault_annex','yosemite','zion','touchdown1','10.100.86.4')  #List hosts here to test
) 
 
if (-not (Test-Path -Path $FolderPath)) { New-Item -ItemType Directory -Path $FolderPath }  #Create foleder for log files
 
$ScriptStartTime = Get-Date 
$FirstRun = $true #Set this so the headers are added to the log when it is first createdThis didn’t take me too long so do I still get to work from home tomorrow or do I need to come on in.
 
while ((Get-Date) -lt $ScriptStartTime.AddHours($HoursToRun)) { 
    foreach ($Computer in $Computers) { 
        ## ICMP Pings 
        $Output = $null 
        $PingTime = (Get-Date).ToString('G') 
         
        $Headers = "Address,StartTime,Duration,Error,Output" 
 
        if ($FirstRun) { 
            $Headers | Set-Content -Path ('{0}\{1}_ICMPPings.csv' -f $FolderPath,$Computer) 
        } 
 
        $PingOutput = (((ping -n 1 -w 5000 $Computer | Out-String).Trim()) -replace '\n','--') 
        $Success = $PingOutput -match 'time\=(\d*)ms' 
        $PingMS = if ($Success) {$Matches[1]} else {$null} 
      
         
        $Output2 = New-Object -TypeName PSObject -Property @{ 
            Address = $Computer 
            StartTime = $PingTime 
            Duration = $PingMS 
            Error = if ($LASTEXITCODE -ne 0) { 'True' } else { 'False' } 
            Output = $PingOutput 
        } | Select-Object Address,StartTime,Duration,Error,Output 
          $Output2
         if($Output2.Error -eq 'True'){
        $Output2 | ConvertTo-Csv -NoTypeInformation | Select-Object -Last 1 | Add-Content -Path ('{0}\{1}_ICMPPings.csv' -f $FolderPath,$Computer) 
            
            }
    } 
    $FirstRun = $false 
    Start-Sleep -Seconds 1 
}