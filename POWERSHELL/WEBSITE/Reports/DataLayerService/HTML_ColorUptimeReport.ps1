# ----------------------------------------------------------------------------- 
# Script: HTML_ColorUptimeReport.ps1 
# Author: ed wilson, msft 
# Date: 08/06/2012 15:11:03 
# Keywords: Scripting Techniques, Web Pages and HTAs 
# comments: Get-Wmiobject, New-Object, Get-Date, Convertto-HTML, Invoke-Item 
# HSG-8-9-2012 
# ----------------------------------------------------------------------------- 
Param( 
  [string]$path = 'C:\inetpub\wwwroot\DLSProcess.html',
  [array]$servers = @("LASMT01","LASMT02","LASMT03","LASMT04","LASMT05","LASMT06","LASMT07","LASMT08","LASMT09","LASMT10","LASMT11","LASMT12","LASMT13","LASMT14","LASMT15","LASMT16","LASMT17","LASMT18","LASMT19","LASMT20","LASMT21","LASMT22","LASMT23","LASMT24","LASMT25)
)

function GetDLSMemUse{


Param(
    [string[]]$servers) 
    Foreach ($s in $servers){  
     if(Test-Connection -cn $s -Quiet -BufferSize 16 -Count 1) 
       { 
        $os = Get-WmiObject -class win32_OperatingSystem -cn $s New-Object psobject -Property @{computer=$s;uptime = (get-date) - $os.converttodatetime($os.lastbootuptime)
       } 
      ELSE 
       {New-Object psobject -Property @{computer=$s;uptime=}
       }

$style = @"))

