
$Today=(Get-Date -format dd-MM-yyyy) # In Date Month Year fomat
$ReportPath = "C:\inetpub\wwwroot\index2.html"
$ReportTitle= "Contoso DataLayer Service Status $today"
$Servers = GC  C:\inetpub\wwwroot\DataLayerService\Configs\Servers.txt
$OutPut = @{Expression={$_.CSName};Label="Server"}, @{Label="Memory Usage";Expression={[String]([int]($_.WorkingSetSize/1MB))+" MB"}}
Import-Module -Name pshtmltable


$Style = @"
 <header>
<meta http-equiv="refresh" content="5" >
</header>
<style>
BODY{font-family:Calibri;font-size:12pt;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:black;background-color:#0BC68D;text-align:center;}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align:center;}
</style>
"@


function Get-ContosoMemUsage {
    Foreach ($server in $servers){
      $Uptime = Get-CimInstance Win32_Process -ComputerName $server -Filter "name = 'DataLayerService.exe'" 
      #$DateCalc = [System.Management.ManagementDateTimeConverter]::ToDateTime($Uptime.CreationDate)
      $memuse = [String]([int]($Uptime.WorkingSetSize/1MB))+" MB"
      $CalcRun = New-TimeSpan -Start ($Uptime.CreationDate) -End (Get-Date)
      New-Object PsObject -Property @{Days = $CalcRun.Days; Hours = $CalcRun.Hours; Minutes = $CalcRun.Minutes;Server = $server;Memory = $memuse} 
      
      } 
    
Start-Sleep -Seconds 5
}
 
while($true){
 Get-ContosoMemUsage | Select Server, Memory, Days, Hours  | New-HTMLTable -
 
 #| ConvertTo-HTML -Body $Style -Title $ReportTitle | Out-File $ReportPath
 }