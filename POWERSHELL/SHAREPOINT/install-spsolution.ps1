Add-PSSnapin "Microsoft.SharePoint.PowerShell" 
 
 function wait4timer($solutionName) 
 {    
     $solution = Get-SPSolution | where-object {$_.Name -eq $solutionName}    
     if ($solution -ne $null)     
     {        
         Write-Host "Waiting to finish soultion timer job" -ForegroundColor Green      
         while ($solution.JobExists -eq $true )          
         {               
             Write-Host "Please wait...Either a Retraction/Deployment is happening" -ForegroundColor DarkYellow           
             sleep 2            
         }                
 
         Write-Host "Finished the solution timer job" -ForegroundColor Green  
         
     }
 }  
 
       
     try
     {
         # Get the WebApplicationURL
         $MyWebApplicationUrl = "http://WEBAPPLICATIONURL";
         
         # Get the Solution Name
         $MywspName = "MySolution.WSP"
         
         # Get the Path of the Solution
         $MywspFullPath = "D:TestMySolution.WSP"
 
         # Try to get the Installed Solutions on the Farm.
         $MyInstalledSolution = Get-SPSolution | Where-Object Name -eq $MywspName
         
         # Verify whether the Solution is installed on the Target Web Application
         if($MyInstalledSolution -ne $null)
         {
             if($MyInstalledSolution.DeployedWebApplications.Count -gt 0)
             {
                 wait4timer($MywspName)  
 
                 # Solution is installed in atleast one WebApplication.  Hence, uninstall from all the web applications.
                 # We need to unInstall from all the WebApplicaiton.  If not, it will throw error while Removing the solution
                 Uninstall-SPSolution $MywspName  -AllWebApplications:$true -confirm:$false
 
                 # Wait till the Timer jobs to Complete
                 wait4timer($MywspName)   
 
                 Write-Host "Remove the Solution from the Farm" -ForegroundColor Green 
                 # Remove the Solution from the Farm
                 Remove-SPSolution $MywspName -Confirm:$false 
 
                 sleep 3
             }
             else
             {
                 wait4timer($MywspName) 
 
                 # Solution not deployed on any of the Web Application.  Go ahead and Remove the Solution from the Farm
                 Remove-SPSolution $MywspName -Confirm:$false 
 
                 sleep 3
             }
         }
 
         wait4timer($MywspName) 
 
         # Add Solution to the Farm
         Add-SPSolution -LiteralPath "$MywspFullPath"
     
         # Install Solution to the WebApplication
         install-spsolution -Identity $MywspName -WebApplication $MyWebApplicationUrl -FullTrustBinDeployment:$true -GACDeployment:$false -Force:$true
 
         # Let the Timer Jobs get finishes       
         wait4timer($MywspName)    
 
         Write-Host "Successfully Deployed to the WebApplication" -ForegroundColor Green 
         
     }
     catch
     {
         Write-Host "Exception Occuerd on DeployWSP : " $Error[0].Exception.Message -ForegroundColor Red  
     }
 