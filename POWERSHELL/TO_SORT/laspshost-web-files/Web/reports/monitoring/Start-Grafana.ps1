Set-Location $PSScriptRoot
$GetExe = get-childitem -Path $PSScriptRoot -Filter *.exe -Recurse | where {$_.Name -match '(grafana\-server.exe)'} | select -first 1
$GetOriginalConfig = (get-childitem -Path $PSScriptRoot -Filter *.ini -Recurse | where {$_.Name -match '(defaults.ini)'} | select -first 1).FullName

$Getcustom = (get-childitem -Path $PSScriptRoot -Filter *.ini -Recurse | where {$_.Name -match '(custom.ini)'} | select -first 1)
write-debug "Matching exes1`n$($GetExe | format-table -AutoSize | out-string)"
if(@($Getcustom).Count -eq 0)
{ 

        (Get-Content -path $GetOriginalConfig -Raw) |  Out-File -FilePath ([io.path]::Combine([io.path]::GetDirectoryName($GetOriginalConfig),'custom.ini')) -NoClobber -Encoding utf8
}
foreach($i in $GetExe)
{
        if(@(Get-process -name $i.BaseName -ErrorAction SilentlyContinue).count -eq 0) 
        { 
                $RootDir = (split-path -path $i.fullname -Parent | Split-Path -Parent)
                [Environment]::SetEnvironmentVariable('HOME', $RootDir )
                start-process -filepath $i.FullName -Verbose -WorkingDirectory  $RootDir 
        }



}

Start-process  http://localhost:3000