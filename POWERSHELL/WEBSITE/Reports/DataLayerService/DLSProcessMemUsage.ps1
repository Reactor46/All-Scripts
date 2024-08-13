# Settings
$Today=(Get-Date -format dd-MM-yyyy) # In Date Month Year fomat
$ReportPath = "C:\inetpub\wwwroot\Reports\Contosodls.html"
$ReportTitle= "Contoso DataLayer Service Status $today"
$Servers = GC  C:\inetpub\wwwroot\Reports\DataLayerService\Configs\Servers.txt
$Service = "ContosoDataLayerService"
$Process = "DataLayerService"
$OutPut = @{Expression={$_.CSName};Label="Server"}, @{Label="Memory Usage";Expression={[String]([int]($_.WS/1MB))+" MB"}}


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
Get-CimInstance Win32_Process -ComputerName $Servers -Filter "name = 'DataLayerService.exe'" -ErrorAction SilentlyContinue -Property * | Sort-Object CSName |             
Select $OutPut | ConvertTo-HTML -Body $Style -Title $ReportTitle | Out-File $reportpath


Start-Sleep -Seconds 5
}

 
while($true){
 Get-ContosoMemUsage
 }