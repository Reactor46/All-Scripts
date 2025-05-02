# Settings
$Today=(Get-Date -format dd-MM-yyyy) # In Date Month Year fomat
$ReportPath = "C:\inetpub\wwwroot\DSLMemoryUsage.html"
$ReportTitle= "Contoso DataLayer Service Status $today"
$Servers = GC  C:\inetpub\wwwroot\DataLayerService\Configs\Servers.txt
$OutPut = @{Expression={$_.CSName};Label="Server"}, @{Label="Memory Usage";Expression={[String]([int]($_.WS/1MB))+" MB"}}, @{Expression={$_.CreationDate};Label="Running Since"}


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
Select $OutPut | ConvertTo-HTML -Body $Style | Out-File $reportpath


#Start-Sleep -Seconds 5
}

#>
#function Show-ContosoMemUsage{

<#
function Get-MemUsage {
$procs = Get-CimInstance Win32_Process -ComputerName $Servers -Filter "name = 'DataLayerService.exe'" -Property *
foreach ($proc in $procs) {
$props = @{'Proccess Name'=$proc.name;'Memory Usage in MB'=$proc.WS / 1MB -as [int]}
New-Object -TypeName PSObject -Property $props
}

function Show-MemUsage {
$params = @{'CssStyleSheet'=$style;'Title'="Contoso DataLayer Serviec Memory Usage";'HTMLFragments'=@($html_pr)}
ConvertTo-EnhancedHTML @params

function Make-HTML {
$params = @{'As'='Table';'PreContent'='<h2>&diams; Processes</h2>';'MakeTableDynamic'=$true;'TableCssClass'='grid'}
$html_pr = Get-MemUsage -ComputerName $Servers |
ConvertTo-EnhancedHTMLFragment @params
}
#>
#Start-Sleep -Seconds 5
#}
#}
#}


 
#while($true){
# Get-ContosoMemUsage
# }