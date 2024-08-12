$Servers = GC -Path $PSScriptRoot\All-Server-List.txt 
ForEach($server in $Servers){
Get-Service -ComputerName $server -ErrorAction SilentlyContinue | Select -Property * | Select MachineName, DisplayName, Status, StartType | Export-CSV -Path $PSScriptRoot\Services\$server-Services.csv -Append -NoTypeInformation}

$files = Get-ChildItem $PSScriptRoot\Services\*.csv 
Get-Content $files | Set-Content $PSScriptRoot\Services\All-Servers-Services.csv
