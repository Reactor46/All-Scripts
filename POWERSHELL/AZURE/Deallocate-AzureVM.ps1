Workflow Deallocate-AzureVM
{

	Param
    (   
        [Parameter(Mandatory=$true)]
        [String]
        $vmName,       

		[Parameter(Mandatory=$true)]
        [String]
        $cloudServiceName 
    )

#
# Deallocate-AzureVM.ps1
#
		set-executionpolicy -scope process  -executionpolicy remotesigned -Force 

		# Specify Azure Subscription Name under which automation runbook need to be executed
		$automationSubscription = 'My subscription Connection'#this name should be same what you have provided as Connection name in Azure Automation Asset settings
    
		# Connect to Azure Subscription
		Connect-Azure -AzureConnectionName $automationSubscription
        
		Select-AzureSubscription -SubscriptionName $automationSubscription 	

		$sourceSubscription = "My subscription Connection" 
		Select-AzureSubscription -SubscriptionName $sourceSubscription
		

		#read the VM configuration and OS disk and data disks(if any)details		
		$sourceVm = Get-AzureVM -ServiceName $cloudServiceName -Name $vmName

		#IMPORTANT - I am getting local time based on India, if your location is different then please change the hours, minutes I am adding below
		$localTime = (Get-Date).ToUniversalTime().AddHours(5).AddMinutes(30).ToShortTimeString().SubString(0,1) 

		if($localTime -ge  "6")
		{
			if ($sourceVm.PowerState -eq 'Started')
			{
				$vmStarted = $true				
				#stop the VM if it is running mode, if already stopped no issue then
				Stop-AzureVM -ServiceName $cloudServiceName -Name $vmName -Force
				Write-Output "successfully stopped azure VM"
			}			
		}
		else
		{
			Write-Output "VM is already stopped"
		}
}