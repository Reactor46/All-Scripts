##=================================================================================
##       Project:  vmware-collect
##        AUTHOR:  SADDAM ZEMMALI
##         eMail:  saddam.zemmali@gmail.com
##       CREATED:  14.07.2018 02:03:01
##      REVISION:  
##       Version:  1.0  ¯\_(ツ)_/¯
##    Repository:  https://github.com/szemmali/vmware-collect
##          Task:  Listed of all VMs with all fields With vCenter Credential
##          FILE:  Collect-vm-info.ps1
##   Description:  This script will Listed of all VMs and Snapshot information
##   Requirement:  --
##          Note:  Connect With USERNAME/PASSWORD Credential 
##          BUGS:  Set-ExecutionPolicy -ExecutionPolicy Bypass
##=================================================================================
$StartTime = Get-Date
Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false
#################################
#    VMware vCenter server name # 
#################################  
$vCenterList = "192.168.1.200"
#$username = Read-Host 'Enter The vCenter Username'
#$password = Read-Host 'Enter The vCenter Password' -AsSecureString    
#$vccredential = New-Object System.Management.Automation.PSCredential ($username, $password)

#################################
# 		REPORTS INFO            # 
#################################  
$Export_VM_INFO = "_VM_INFO_" + (Get-Date -UFormat "%d-%b-%Y-%H-%M") + ".csv"
$All_VM_Export = "All_VM_Export_" + (Get-Date -UFormat "%d-%b-%Y-%H-%M") + ".csv"
$Snap_INFO = "All_Snap_Export_" + (Get-Date -UFormat "%d-%b-%Y-%H-%M") + ".csv"
$DCReport = @()   
$ReportAllSnap = @()

#################################
# 		START COLLECT vCenter   # 
################################# 
foreach ($vCenter in $vCenterList){
	
	Write-Host "Connecting to $vCenter..." -Foregroundcolor "Yellow" -NoNewLine
	$connection = Connect-VIServer -Server $vCenter  -User "administrator@vsphere.local" -Password "J@bb3rJ4w" -ErrorAction SilentlyContinue -WarningAction 0 | Out-Null
	If($? -Eq $True){
	Write-Host "Connected" -Foregroundcolor "Green" 
	############################
	# Gathering VM information #
	############################
	Write-Host "Gathering VM Information"
	$Report = @() 
	
	foreach ($vm in Get-VM) {  
			#Write-Host $vm.Name  
			#All global info here  
			$GlobalHDDinfo = $vm | Get-HardDisk  
			$vNicInfo =      $vm | Get-NetworkAdapter  
			$Snapshotinfo =  $vm | Get-Snapshot  
			$Resourceinfo =  $vm | Get-VMResourceConfiguration  

			#IPinfo  
			$IPs = $vm.Guest.IPAddress -join " , " #$vm.Guest.IPAddres[0] <#it will take first ip#>  

			#FQDN - AD domain name  
			$OriginalHostName = $($vm.ExtensionData.Guest.Hostname -split '\.')[0]  
			$Domainname = $($vm.ExtensionData.Guest.Hostname -split '\.')[1,2] -join '.'  

			#All hardisk individual capacity  
			$TotalHDDs = $vm.ProvisionedSpaceGB -as [int]  

			#All hardisk individual capacity  
			$HDDsGB = $($GlobalHDDinfo | Select-Object -ExpandProperty CapacityGB) -join " + "  

			#All HDD disk type,($vdisk.Capacity /1GB -as [int])}  
			$HDDtype = foreach ($HDDtype in $GlobalHDDinfo) {"{0}={1}GB"-f ($HDDtype.Name), ($HDDtype.StorageFormat)}  
			$HDDtypeResult = $HDDtype -join (", ")  

			#Associated Datastores  
			$datastore = $(Get-Datastore -vm $vm) -split ", " -join " , "  

			#Guest OS Internal HDD info  
			$internalHDDinfo = ($vm | Get-VMGuest).ExtensionData.disk  
			$internalHDD = foreach ($vdisk in $internalHDDinfo) {"{0}={1}GB/{2}GB"-f ($vdisk.DiskPath), ($vdisk.FreeSpace /1GB -as [int]),($vdisk.Capacity /1GB -as [int])}  
			$internalHDDResult = $internalHDD -join (", ")  

			#vCenter Server  
			$vCenter = $vm.ExtensionData.Client.ServiceUrl.Split('/')[2].trimend(":443")   

			#VM Macaddress  
			$Macaddress = $vNicInfo.MacAddress -join " , "  

			#Vmdks and its location  
			$vmdk = $GlobalHDDinfo.filename -join " , "  

			#Snapshot info  
			$snapshot = $Snapshotinfo.count  

			#Datacenter info  
			$datacenter = $vm | Get-Datacenter | Select-Object -ExpandProperty name  

			#Cluster info  
			$cluster = $vm | Get-Cluster | Select-Object -ExpandProperty name  

			#vNic Info  
			$vNics = foreach ($vNic in $VnicInfo) {"{0}={1}"-f ($vnic.Name.split("")[2]), ($vNic.Type)}  
			$vnic = $vNics -join (" , ")  

			#Virtual Port group Info  
			$portgroup = $vNicInfo.NetworkName -join " , "  

			#RDM Disk Info  
			$RDMInfo = $GlobalHDDinfo | Where-Object {$_.DiskType -eq "RawPhysical"-or $_.DiskType -eq "RawVirtual"}   
			$RDMHDDs = foreach ($RDM in $RDMInfo) {"{0}/{1}/{2}/{3}"-f ($RDM.Name), ($RDM.DiskType),($RDM.Filename), ($RDM.ScsiCanonicalName)}  
			$RDMs = $RDMHDDs -join (" , ")  

			$Vmresult = New-Object PSObject   
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Datacenter" -Value $datacenter  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "vCenter Server" -Value $vCenter  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Cluster" -Value $cluster 
			$Vmresult | Add-Member -MemberType NoteProperty -Name "EsxiHost" -Value $VM.VMHost 
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMName" -Value $vm.Name  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "IP Address" -Value $IPs  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "PowerState" -Value $vm.PowerState 
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Domain Name" -Value $Domainname  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $OriginalHostName  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "vCPU" -Value $vm.NumCpu  
			$Vmresult | Add-Member -MemberType NoteProperty -Name CPUSocket -Value $vm.ExtensionData.config.hardware.NumCPU  
			$Vmresult | Add-Member -MemberType NoteProperty -Name Corepersocket -Value $vm.ExtensionData.config.hardware.NumCoresPerSocket  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "RAM(GB)" -Value $vm.MemoryGB  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Total-HDD(GB)" -Value $TotalHDDs  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "HDDs(GB)" -Value $HDDsGB  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "HDDsType" -Value $HDDtypeResult  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $datastore  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Partition/Size" -Value $internalHDDResult  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Installed-OS" -Value $vm.guest.OSFullName  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Setting-OS" -Value $VM.ExtensionData.summary.config.guestfullname  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Hardware Version" -Value $vm.HardwareVersion  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Folder" -Value $vm.folder  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "MacAddress" -Value $macaddress  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMX" -Value $vm.ExtensionData.config.files.VMpathname  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMDK" -Value $vmdk  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMTools Status" -Value $vm.ExtensionData.Guest.ToolsStatus  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMTools Version" -Value $vm.ExtensionData.Guest.ToolsVersion  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMTools Version Status" -Value $vm.ExtensionData.Guest.ToolsVersionStatus  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "VMTools Running Status" -Value $vm.ExtensionData.Guest.ToolsRunningStatus  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "SnapShots" -Value $snapshot  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "vNic" -Value $vNic  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "Portgroup" -Value $portgroup  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "RDM" -Value $RDMs  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "NumCpuShares" -Value $Resourceinfo.NumCpuShares  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "CpuReservationMhz" -Value $Resourceinfo.CpuReservationMhz  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "CpuLimitMhz" -Value $Resourceinfo.CpuLimitMhz  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "CpuSharesLevel" -Value $Resourceinfo.CpuSharesLevel  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "NumMemShares" -Value $Resourceinfo.NumMemShares  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "MemReservationGB" -Value $Resourceinfo.MemReservationGB  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "MemLimitGB" -Value $Resourceinfo.MemLimitGB  
			$Vmresult | Add-Member -MemberType NoteProperty -Name "MemSharesLevel" -Value $Resourceinfo.MemSharesLevel 
			$Vmresult | Add-Member -MemberType NoteProperty -Name "CpuAffinityList" -Value $Resourceinfo.CpuAffinityList  

		$Report += $Vmresult     
		}  
		Write-Host "Create CSV File with VM Information from $vCenter" 
		$report | Export-Csv $vCenter$Export_VM_INFO -NoTypeInformation -UseCulture  
		#Invoke-Item    $vCenter$Export_VM_INFO
		$DCReport +=$Report

		##############################
		# Gathering Snap information #
		##############################
		Write-Host "Gathering Snapshot Information"
		$ReportSnap = @()
	
		Get-VM | Get-Snapshot | %{
			$Snap = {} | Select VM,VMId,Name,Created,Description,Id
			$Snap.VM = $_.vm.name
			$Snap.Description = $_.description
			$Snap.VMId = $_.VMId
			$Snap.Name = $_.name
			$Snap.Created = $_.created
			$Snap.Id = $_.Id
			$ReportSnap += $Snap
		}
		$ReportAllSnap += $ReportSnap
	}

		Else{
			  Write-Host "Error in Connecting to $vCenter; Try Again with correct user name & password!" -Foregroundcolor "Red" 
		}    
  
	Write-Host "Export All Snapshot Information"             
	$ReportAllSnap  | Export-Csv $Snap_INFO -NoTypeInformation -UseCulture 
	
	Write-Host "Export All VM Information" 
	$DCReport | Export-Csv $All_VM_Export -NoTypeInformation -UseCulture 
	
	##############################
	# Disconnect session from VC #
	##############################  
	Write-Host "Disconnect to $vCenter..." -Foregroundcolor "Yellow" -NoNewLine 
	$disconnection = Disconnect-VIServer -Server $vCenter  -Force -confirm:$false  -ErrorAction SilentlyContinue -WarningAction 0 | Out-Null
	
	If($? -Eq $True){
		Write-Host "Disconnected" -Foregroundcolor "Green" 
		Write-Host "#####################################" -Foregroundcolor "Blue" 
		}

	Else{
		Write-Host "Error in Disconnecting to $vCenter" -Foregroundcolor "Red" 
	   }
}    
##############################
# 	Open Collect VM Info 	 #
############################## 
Write-Host "Open Collect VM Info Report $All_VM_Export"
Invoke-Item    $All_VM_Export

###############################
# Open Collect SnapShots Info #
############################### 
Write-Host "Open Collect SnapShots Info Report $Snap_INFO"
Invoke-Item    $Snap_INFO

$EndTime = Get-Date
$duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalMinutes,2)
Write-Host "================================"
Write-Host "vSphere Collect Complete!"
Write-Host "StartTime: $StartTime"
Write-Host "  EndTime: $EndTime"
Write-Host "  Duration: $duration minutes"
Write-Host "================================"