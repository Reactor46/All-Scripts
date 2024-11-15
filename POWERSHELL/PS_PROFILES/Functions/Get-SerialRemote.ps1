######################################################################   
# Powershell script to get the the serial numbers on remote servers   
# It will give the serial numbers on remote servers and export to csv 
# Customized script useful to every one   
# Please contact  mllsatyanarayana@gmail.com for any suggestions#   
#########################################################################  
####################serial start################# 
 
 function Get-Serial { 
 param( 
 $computername =$env:computername 
 ) 
 
 $os = Get-WmiObject Win32_bios -ComputerName $computername -ea silentlycontinue 
 if($os){ 
 
 $SerialNumber =$os.SerialNumber 
 
 $servername=$os.PSComputerName  
  
 
  
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty SerialNumber  $SerialNumber 
 $results |Add-Member noteproperty ComputerName  $servername 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,SerialNumber 
 
 } 
 
 
 else 
 
 { 
 
 $results =New-Object psobject 
 
 $results =new-object psobject 
 $results |Add-Member noteproperty SerialNumber "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
 
 
  
 #display the results 
 
 $results | Select-Object computername,SerialNumber 
 
 
 
 
 } 
 
 
 
 } 
 
 $infserial =@() 
 
 
 foreach($allserver in $allservers){ 
 
$infserial += Get-Serial $allserver  
 } 
 
 $infserial  
 
 
 
 
 
<#   ####################serial end################# 
 
 
   #save the the servers in any location 
    $servers = Get-Content -Path "C:\LazyWinAdmin\Servers\servers2.txt" 
 
   foreach ($ser in $servers) 
 
   { 
    
   get-serial -computername $ser | Export-Csv -Path c:\LazyWinAdmin\Servers\serial.csv 
    
   }
   #>