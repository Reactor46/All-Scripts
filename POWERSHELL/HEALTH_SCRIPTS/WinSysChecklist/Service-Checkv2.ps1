$today=(Get-Date -format dd-MM-yyyy) # In Date Month Year fomat
$reportpath="C:\LazyWinAdmin\WinSysChecklist\Services Check\ServicesReport.$today.html"
$ReportTitle="Contoso Server Service Status $today"
$computers= GC -Path C:\LazyWinAdmin\WinSysChecklist\Configs\W3SVC.txt
$computers2= GC -Path C:\LazyWinAdmin\WinSysChecklist\Configs\Contosodls.txt
$computers3="LASSVC03"
$computers4="LASMCE01, LASMCE02"
$computers5="LASCAPSMT01, LASCAPSMT02, LASCAPSMT05, LASCAPSMT06"
$computers6="LASCHAT01, LASCHAT02"
$computers7="LASPROCDB02"
$computers8="LASPROCAPP01"
$computers9="LASPROCAPP04"
$computers10="LASRFAX01"
$computers11="LASCODE02"
$computers12="LASITS01, LASITS02"
$computers13="LASPRINT02"
$computers14="LASINFRA02"
$computers15="DPC-16736"
$services="W3SVC"
$services2="ContosoDataLayerService"
$services3="CollectionsAgentTimeService", "Contoso*"
$services4="CreditEngine"
$services5="CreditPullService", "Contoso*"
$services6="WhosOn*"
$services7="MSSQLSERVER", "SQLSERVERAGENT"
$services8="P360*"
$services9="Service1"
$services10="RF*"
$services11="AccuRev*", "JIRA050414101333"
$services12="*Symantec*"
$services13="Spooler"
$services14="Schedule"
$services15="Readerboard_Server"

####******************* set email parameters ****************** ######

$from="winsys_service_chk@creditone.com"
$to="john.battista@creditone.com"
$smtpserver="mailgateway.fnmb.corp"
####*******************######################****************** ######

$Style = @"
<style>
BODY{font-family:Calibri;font-size:12pt;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:black;background-color:#0BC68D;text-align:center;}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align:center;}
</style>
"@

### $Computers
$array = @()            
foreach($computer in $computers) { 
Foreach($service in $services) {         
 $svc = Get-Service $service -ComputerName $computer -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append

### $Computers2
$array = @()            
foreach($computer in $computers2) { 
Foreach($service in $services2) {         
 $svc = Get-Service $service -ComputerName $computer -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}       
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append

### $Computers3
$array = @()            
foreach($computer in $computers3) { 
Foreach($service in $services3) {         
 $svc = Get-Service $service -ComputerName $computer -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}      
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append        

### $Computers4     
$array = @()            
foreach($computer in $computers4) { 
Foreach($service in $services4) {         
 $svc = Get-Service $service -ComputerName $computer -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}   
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append

### $Computers5                
$array = @()            
foreach($computer in $computers5) { 
Foreach($service in $services5) {         
 $svc = Get-Service $service -ComputerName $computer -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}        
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"}
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append      

### $Computers6     
$array = @()            
foreach($computer6 in $computers6) { 
Foreach($service6 in $services6) {         
 $svc = Get-Service $service6 -ComputerName $computer6 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer6          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}       
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"}
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append        

### $Computers7    
$array = @()            
foreach($computer7 in $computers7) { 
Foreach($service7 in $services7) {         
 $svc = Get-Service $service7 -ComputerName $computer7 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer7          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}      
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append        

### $Computers8     
$array = @()            
foreach($computer8 in $computers8) { 
Foreach($service8 in $services8) {         
 $svc = Get-Service $service8 -ComputerName $computer8 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer8          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}      
$array | Select Computername,displayname,name,status | 
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append       

### $Computers9     
$array = @()            
foreach($computer9 in $computers9) { 
Foreach($service9 in $services9) {         
 $svc = Get-Service $service9 -ComputerName $computer9 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer9          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}       
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append       

### $Computers10     
$array = @()            
foreach($computer10 in $computers10) { 
Foreach($service10 in $services10) {         
 $svc = Get-Service $service10 -ComputerName $computer10 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer10          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}    
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append           

### $Computers11    
$array = @()            
foreach($computer11 in $computers11) { 
Foreach($service11 in $services11) {         
 $svc = Get-Service $service11 -ComputerName $computer11 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer11          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}          
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append     

### $Computers12    
$array = @()            
foreach($computer12 in $computers12) { 
Foreach($service12 in $services12) {         
 $svc = Get-Service $service12 -ComputerName $computer12 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer12          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}      
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append       

### $Computers13    
$array = @()            
foreach($computer13 in $computers13) { 
Foreach($service13 in $services13) {         
 $svc = Get-Service $service13 -ComputerName $computer13 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer13          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}     

$array | Select Computername,displayname,name,status | 
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" | 
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append
         
### $Computers14    
$array = @()            
foreach($computer14 in $computers14) { 
Foreach($service14 in $services14) {         
 $svc = Get-Service $service14 -ComputerName $computer14 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer14          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}   
$array | Select Computername,displayname,name,status | 
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" | 
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"} 
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append
            
### $Computers15    
$array = @()            
foreach($computer15 in $computers15) { 
Foreach($service15 in $services15) {         
 $svc = Get-Service $service15 -ComputerName $computer15 -ea "0"            
 $obj = New-Object psobject -Property @{  
  ComputerName = $computer15          
  DisplayName=$svc.displayname   
  Name = $svc.name            
  Status = $svc.status
    }  
     $array += $obj                 
}   
}               
## Build HTML Report                       
$array | Select Computername,displayname,name,status |
    ConvertTo-Html -property 'ComputerName','Displayname','Name','Status' -head $Style -body "<h1> $ReportTitle </h1>" |
        foreach {if($_ -like "*<td>Running</td>*"){$_ -replace "<tr>", "<tr bgcolor=#089437>"}
            elseif($_ -like "*<td>Stopped</td>*" -or "*<td>Stopping</td>*" -or "*<td>Pending</td>*" -or "*<td>Starting</td>*")
                {$_ -replace "<tr>", "<tr bgcolor=#C60B1C>"}  else{$_}} | out-file $reportpath -Append


$body = Get-Content $reportpath
Send-MailMessage -To $to -From $from -Subject "Daily Service Report" -Body $body -BodyAsHtml -SmtpServer $smtpserver