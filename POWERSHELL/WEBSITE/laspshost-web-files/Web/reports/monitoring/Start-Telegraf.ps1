Set-Location $PSScriptRoot
$GetExe = get-childitem -Path $PSScriptRoot -Filter *.exe -Recurse | where {$_.Name -match '(telegraf.exe)'} | select -first 1
$Config =  (  get-childitem -Path $PSScriptRoot -Filter *.conf -Recurse | where {$_.Name -match 'telegraf.conf'} | select -first 1).FullName
write-debug "Configuration File Path: $($Config)"
write-debug "Matching exes: `n$($GetExe | format-table -AutoSize | out-string)"
Get-Process -Name telegraf -ErrorAction SilentlyContinue | Stop-Process -Verbose 
foreach($i in $GetExe)
{
        if(@(Get-process -name $i.BaseName -ErrorAction SilentlyContinue).count -eq 0) 
        { 
                set-variable -name WorkingDir -value (split-path -Path $i.FullName -Parent) -Verbose
                [Environment]::SetEnvironmentVariable('TELEGRAF_CONFIG_PATH',$Config)
                start-process -filepath $i.FullName -Verbose 
        }

}
