param( 
[Parameter(Mandatory=$true)] 
[string]$server,
[string]$domain,
[string]$save
) 
 
 
$machines = $server 
$ErrorActionPreference = "SilentlyContinue" 
$date = Get-Date -Format "MM-dd-yyyy" 
 
#Clear Job Cache 
Get-Job | Remove-Job -Force 
 
$jobscript = { 
 
Param($machine) 
     
    $pingstate = "" 
 
    $pingstate = Test-Connection $machine -Count 3 -Quiet 
 
    $IISVersionString = "" 
     
    $IISVersionString = ((reg query \\$machine\HKLM\SOFTWARE\Microsoft\InetStp\ | findstr VersionString).Replace(" ","")).Replace("VersionStringREG_SZ","") 
 
    if ($IISVersionString -eq ""){ 
 
        $IISVersionString = "Not Installed." 
     
    } 
 
    $results = [PSCustomObject]@{ 
     
        ServerName = $machine 
        IISVersion = $IISVersionString 
        Pingable = $pingstate 
 
    } 
 
    $results 
 
} 
 
 
Foreach($machine in $machines){ 
 
Start-Job -ScriptBlock $jobscript -ArgumentList $machine 
 
} 
 
$output = Get-Job | Wait-Job | Receive-Job  
 
$output | Export-Csv $PSScriptRoot\$save\$domain-$date.csv -NoTypeInformation -Append 
 
 
 
$output