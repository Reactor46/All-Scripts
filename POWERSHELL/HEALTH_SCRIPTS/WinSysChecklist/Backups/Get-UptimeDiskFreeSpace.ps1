# ----------------------------------------------------------------------------- 
# Script: QueryAD_HTML_Uptime_FreespaceReport.ps1 
# Author: ed wilson, msft 
# Date: 08/07/2012 15:11:03 
# Keywords: Scripting Techniques, Web Pages and HTAs 
# comments: added freespace to the script  
# Get-Wmiobject, New-Object, Get-Date, Convertto-HTML, Invoke-Item 
# HSG-8-7-2012, HSG-8-8-2012 
# ----------------------------------------------------------------------------- 
Param( 
  [string]$path = "C:\LazyWinAdmin\WinSysChecklist\Logs\uptimeFreespace.html") 
 
Function Get-UpTime 
{ Param ([string[]]$servers) 
  Foreach ($s in $servers)  
   {  
     if(Test-Connection -cn $s -Quiet -BufferSize 16 -Count 1) 
       { 
        $os = Get-WmiObject -class win32_OperatingSystem -cn $s  
        New-Object psobject -Property @{computer=$s; 
        uptime = (get-date) - $os.converttodatetime($os.lastbootuptime)} } 
      ELSE 
       { New-Object psobject -Property @{computer=$s; uptime = "DOWN"} } 
      } 
    } #end function Get-Uptime 
 
Function Get-DiskSpace 
{ 
 Param ([string[]]$servers) 
  Foreach ($s in $servers)  
   {  
    if(Test-Connection -cn $s -Quiet -BufferSize 16 -Count 1) 
       { 
        Get-WmiObject -Class win32_volume -cn $s | 
        Select-Object @{LABEL='Comptuer';EXPRESSION={$s}}, 
         driveletter, label,  
         @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f ($_.freespace/1GB)}} 
        } 
    } #end foreach $s 
} #end function Get-DiskSpace 
 
# Entry Point *** 
 
[array]$servers = ([adsisearcher]"(&(objectcategory=computer)(OperatingSystem=*server*))").findall() | 
      foreach-object {([adsi]$_.path).cn} 
 
$upTime = Get-UpTime -servers $servers |  
ConvertTo-Html -As Table -Fragment -PreContent " 
  <h2>Server Uptime Report</h2> 
  The following report was run on $(get-date)" | Out-String 
 
$disk = Get-DiskSpace -servers $servers |  
ConvertTo-Html -As Table -Fragment -PreContent " 
  <h2>Disk Report</h2> "| Out-String   
    
ConvertTo-Html -PreContent "<h1>Server Uptime and Disk Report</h1>" ` 
  -PostContent $upTime, $disk >> $path  
 Invoke-Item $path 