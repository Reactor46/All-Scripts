param( 
[Parameter(Mandatory=$true,Position=0)] 
[string]$serverlistTXT) 
 
 
$machines = (Get-content "$serverlistTXT").Trim() 
$ErrorActionPreference = "SilentlyContinue" 
$datetime = Get-Date -Format "MMddyy-HHmmss" 
 
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
 
$output | Export-Csv $PSScriptRoot\IISversionQuery_$datetime.csv -NoTypeInformation -Append 

 
$output