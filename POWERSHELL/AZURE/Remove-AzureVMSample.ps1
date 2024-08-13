<#
.SYNOPSIS 
    Removes a Virtual Machine (VM) on Azure, deletes attached VHD(s) too.
    Runbook waits untill VM is removed.

.DESCRIPTION
    This runbook removes a VM on Azure
    The Connect-Azure runbook needs to be imported and published before this runbook can be sucessfully run.
       
.PARAMETER AzureConnectionName
    Name of the Azure connection asset that was created in the Automation service.
    This connection asset contains the subscription id and the name of the certificate asset that 
    holds the management certificate for this subscription.
    
.PARAMETER ServiceName
    Name of the cloud service which VM will belong to. A new cloud service will be created if cloud service by name ServiceName does not exists

.PARAMETER VMName    
    Name of the virtual machine. 

.EXAMPLE
    1) Create Certificate Asset "myCert":
     Use the certificate file (.pfx or .cer file) to create a Certificate asset for ex. "myCred" in 
     Azure -> Automation -> select automation account "MyAutomationAccount" -> Assets -> Add Setting -> Add Credential -> 
     Certificate -> Provide name "myCred" and upload the certificate file (.pfx or .cer)
     
     The same certificate must be associated with the subscription, You can verify the same for your subscription 
     at Azure -> Settings -> Management Certificates


 2) Create Azure Connection Asset "AzureConnection"
     Azure -> Automation -> select automation account "MyAutomationAccount" -> Assets -> Add Setting 
     -> Add Connection -> Select 'Azure' from dropdown -> Provide name ex. "AzureConnection"  ->
     Provide AutomationCertificateName "myCert" you created in step 1 and subscription Id
     
 3) To run runbook: Test or Start the runbook from Author tab
   
   to call from another runbook, ex:
Remove-AzureVMSample -AzureConnectionName "AzureConnection" -ServiceName "myService" -VMName "myVM" 


.NOTES
    AUTHOR: Viv Lingaiah
    LASTEDIT: Apr 16 , 2014 
#>
workflow Remove-AzureVMSample
{
    Param
    (
        [parameter(Mandatory=$true)] [String] $AzureConnectionName,
	    [parameter(Mandatory=$true)] [String] $ServiceName,
        [parameter(Mandatory=$true)] [String] $VMName
    )
     
    # Call the Connect-Azure Runbook to set up the connection to Azure using the Automation connection asset
    Connect-Azure -AzureConnectionName $AzureConnectionName 
       
    InlineScript
    {
        # Select the Azure subscription we will be working against
        Select-AzureSubscription -SubscriptionName $Using:AzureConnectionName
        $sub = Get-AzureSubscription -SubscriptionName $Using:AzureConnectionName
            
        # Check whether a VM by name $VMName exists
        Write-Output ("Checking whether VM '{0}' exists.." -f $Using:VMName)
        $AzureVM = Get-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName
        if ($AzureVM -eq $null) 
       	{
            Write-Output ("VM '{0}' does not exist to remove" -f $Using:VMName)
        }
        else
        {
            Write-Output ("VM '{0}' exists. Stopping and Removing it..." -f $Using:VMName)
            $OSDisk = $AzureVM | Get-AzureOSDisk
            $OSDiskName = $OSDisk.DiskName
            
            if ($AzureVM.PowerState -ne "Stopped") 
       	    {
                  Write-Output ("Stopping VM '{0}'.." -f $Using:VMName)
          	      $stopVM = Stop-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName -Force
                  
            }
            $GetVM = Get-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName 
            if ($GetVM -eq $null) 
       	    {
          	     throw "Could not get VM '{0}' info after stopping it " -f $VMName
            }
            if ($GetVM.Powerstate -ne "Stopped") 
       	    {
          	     throw "VM {0} was not stopped" -f $VMName
            }
            else
            {
                Write-Output ("VM '{0}' was stopped successfully" -f $Using:VMName)
            }     
            
            $RemoveVM = Remove-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName -DeleteVHD
            if ($RemoveVM -eq $null) 
       	    {
          	 throw "'Remove-AzureVM' activity: returned null. VM {0} might not have been removed " -f $VMName
            }
            
            $GetVM2 = Get-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName 
            if ($GetVM2 -ne $null) 
       	    {
          	 throw "'Get-AzureVM' activity: returned non null. VM {0} might not have been removed " -f $VMName
            }  
            
            $timeout = 420  
            $waitperiod = 120
            while(($OSDisk -ne $null) -and ($timeout -gt 0))
            {
                Start-Sleep -seconds $waitperiod
                $OSDisk = Get-AzureDisk -DiskName $OSDiskName
                $timeout -= $waitperiod
                Write-Output ("Waiting for Azure disk {0} to be deleted.." -f $OSDiskName) 
            } 
            
            $GetVM = Get-AzureVM -ServiceName $Using:ServiceName -Name $Using:VMName 
            
            if($GetVM -eq $null)
            {
                Write-Output ("VM '{0}' was removed successfully" -f $Using:VMName)
            }
            else
            {
                throw "VM {0} was not removed. Please check" -f $Using:VMName
            }        
        }        
    } 
}