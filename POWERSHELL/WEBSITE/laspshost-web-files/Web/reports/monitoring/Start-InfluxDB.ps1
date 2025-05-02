Set-Location $PSScriptRoot
$GetExe = get-childitem -Path $PSScriptRoot -Filter *.exe -Recurse | where {$_.Name -match '(influxd.exe)|(influx\.exe)'}
[Environment]::SetEnvironmentVariable('HOME','C:\Influx\')

foreach($i in $GetExe)
{
        if(@(Get-process -name $i.BaseName -ErrorAction SilentlyContinue).count -eq 0) 
        { start-process -filepath $i.FullName -Verbose }

}