#############################################################################
# Script monitors specified services
#Define Server & Services Variable
#$ServerListB = Get-Content "C:\Users\bradley.voris\Desktop\Checkservices\EvJserver.txt"
#$ServerListC =  Get-Content "C:\Users\bradley.voris\Desktop\Checkservices\EvDAserver.txt"

#Define other variables
If ($checkrep -like "True")
{
Remove-Item "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\servicehosts.htm"
}
New-Item "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\servicehosts.htm" -type file
#Get Services Status
foreach ($machineName in $serverlist) 
 { 
  foreach ($service in $serviceslist)
     {
      $serviceStatus = get-service -ComputerName $machineName -DisplayName $service
		 if ($serviceStatus.status -eq "Running") {

         $svcName = $serviceStatus.displayname 
         Add-Content $report "<tr>" 
    
                                                   }

	        else 
                                                   { 

         $svcName = $serviceStatus.displayname 
         Add-Content $report "<tr>" 


         
                                                  } 

             

       } 


 } 

}
#Call Function
#Close HTMl Tables