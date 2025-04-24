Set-Location $PSScriptRoot
$GetExe = get-childitem -Path $PSScriptRoot -Filter *.exe -Recurse | where {$_.Name -match '(chronograf.exe)'} | select -first 1

write-debug "Matching exes1`n$($GetExe | format-table -AutoSize | out-string)"

foreach($i in $GetExe)
{
        if(@(Get-process -name $i.BaseName -ErrorAction SilentlyContinue).count -eq 0) 
        { 

                start-process -filepath $i.FullName -Verbose 
        }

}