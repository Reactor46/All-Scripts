<# 
 .Synopsis
  Update the configuration of a virtual machine to enable monitoring support for SAP.

 .Description
  Updates the configuration of a virtual machine to enable or update the support for monitoring for SAP systems that are installed on the virtual machine.
  The commandlet installs the extension that collects the performance data and makes it discoverable for the SAP system.

 .Parameter DisableWAD
  If this parameter is provided, the commandlet will not enable Windows Azure Diagnostics for this virtual machine.    

 .Example
   Update-VMConfigForSAP_GUI
#>
function Update-VMConfigForSAP_GUI
{
	param
	(
        [Switch] $DisableWAD
	)

	Select-Subscription

    $selectedVM = Select-VM
    if (-not $selectedVM)
    {
        return
    }

    Update-VMConfigForSAP -VMName $selectedVM.Name -VMServiceName $selectedVM.ServiceName @PSBoundParameters
}

<# 
 .Synopsis
  Update the configuration of a virtual machine to enable monitoring support for SAP.

 .Description
  Updates the configuration of a virtual machine to enable or update the support for monitoring for SAP systems that are installed on the virtual machine.
  The commandlet installs the extension that collects the performance data and makes it discoverable for the SAP system.

 .Parameter VMName
  The name of the virtual machine that should be enable for monitoring.

 .Parameter VMServiceName
  The name of the cloud service that the virtual machine is part of.

 .Parameter DisableWAD
  If this parameter is provided, the commandlet will not enable Windows Azure Diagnostics for this virtual machine.    

 .Example
   Update-VMConfigForSAP -VMName SAPVM -ServiceName SAPLandscape
#>
function Update-VMConfigForSAP
{
    param
    (
        [Parameter(Mandatory=$True)] $VMName,
        [Parameter(Mandatory=$True)] $VMServiceName,
        [Switch] $DisableWAD
    )

    Write-Verbose "Retrieving VM..."
    $selectedVM = Get-AzureVM -ServiceName $VMServiceName -Name $VMName
    $selectedRole = Get-AzureRole -ServiceName $VMServiceName -RoleName $VMName -Slot Production

    if (-not $selectedVM)
    {
        $subName = (Get-AzureSubscription -Current).SubscriptionName
        Write-Error "No VM with name $VMName and Service Name $VMServiceName in subscription $subName found"
        return
    }
    if (-not $selectedVM.VM.ProvisionGuestAgent)
    {        
        Write-Warning $missingGuestAgentWarning
        return
    }      
	
    $osdisk = Get-AzureOSDisk -VM $selectedVM
	$disks = Get-AzureDataDisk -VM $selectedVM
    
    $sapmonPublicConfig = @()
    $sapmonPrivateConfig = @()
    $cpuOvercommit = 0
    if ($selectedVM.InstanceSize -eq "ExtraSmall")
    {
        Write-Verbose "VM Size is ExtraSmall - setting overcommitted setting"
        $cpuOvercommit = 1
    }
    $memOvercommit = 0
    $vmsize = $selectedVM.InstanceSize
    switch ($selectedVM.InstanceSize)
    {
        "ExtraSmall" { $vmsize = "ExtraSmall (A0)" }
        "Small" { $vmsize = "Small (A1)" }
        "Medium" { $vmsize = "Medium (A2)" }
        "Large" { $vmsize = "Large (A3)" }
        "ExtraLarge" { $vmsize = "ExtraLarge (A4)" }
    }
    $sapmonPublicConfig += @{ key = "vmsize";value=$vmsize}
    $sapmonPublicConfig += @{ key = "vm.roleinstance";value=$selectedVM.VM.RoleName}
    $sapmonPublicConfig += @{ key = "vm.role";value=$ROLECONTENT}
    $sapmonPublicConfig += @{ key = "vm.deploymentid";value=$selectedRole.DeploymentID}
    $sapmonPublicConfig += @{ key = "vm.memory.isovercommitted";value=$memOvercommit}
    $sapmonPublicConfig += @{ key = "vm.cpu.isovercommitted";value=$cpuOvercommit}
    $sapmonPublicConfig += @{ key = "script.version";value=$CurrentScriptVersion}
    $sapmonPublicConfig += @{ key = "verbose";value="0"}
    

	# Get Disks
    $accounts = @()
    $accountName = Get-StorageAccountFromUri $osdisk.MediaLink
    $storageKey = (Get-AzureStorageKey  -StorageAccountName $accountName).Primary
    $accounts += @{Name=$accountName;Key=$storageKey}

    Write-Host "[INFO] Adding configuration for OS disk"
    $sapmonPublicConfig += @{ key = "osdisk.connminute";value=($accountName + ".minute")}
    $sapmonPublicConfig += @{ key = "osdisk.connhour";value=($accountName + ".hour")}
    $sapmonPublicConfig += @{ key = "osdisk.name";value=($osdisk.MediaLink.Segments[$osdisk.MediaLink.Segments.Count - 1])}
        
	# Get Storage accounts from disks
    $diskNumber = 1
    foreach ($disk in $disks)
    {
        $accountName = Get-StorageAccountFromUri $disk.MediaLink
        if (-not ($accounts | where Name -eq $accountName))
        {
            $storageKey = (Get-AzureStorageKey  -StorageAccountName $accountName).Primary
            $accounts += @{Name=$accountName;Key=$storageKey}
        }            

        Write-Host ("[INFO] Adding configuration for data disk " + $disk.DiskName)
        $sapmonPublicConfig += @{ key = "disk.lun.$diskNumber";value=$disk.Lun}
        $sapmonPublicConfig += @{ key = "disk.connminute.$diskNumber";value=($accountName + ".minute")}
        $sapmonPublicConfig += @{ key = "disk.connhour.$diskNumber";value=($accountName + ".hour")}
        $sapmonPublicConfig += @{ key = "disk.name.$diskNumber";value=($disk.MediaLink.Segments[$disk.MediaLink.Segments.Count - 1])}

        $diskNumber += 1
    }
    
	# Check storage accounts for analytics
    foreach ($account in $accounts)
    {
        Write-Verbose "Testing Storage Metrics for $account"

        $storageKey = $null
        $context    = $null
        $sas = $null

		$storage = Get-AzureStorageAccountWA -StorageAccountName $account.Name
		if ($storage.AccountType -like "Standard_*")
		{
			$currentConfig = Get-StorageAnalytics -AccountName $account.Name

			if (-not (Check-StorageAnalytics $currentConfig))
			{
				Write-Host "[INFO] Enabling Storage Account Metrics for storage account"$account.Name

				# Enable analytics on storage accounts
				Set-StorageAnalytics -AccountName $account.Name -StorageServiceProperties $DefaultStorageAnalyticsConfig
			}            
			
			$endpoint = ($storage.Endpoints | where { $_ -like "*table*" })
			$hourUri = "$endpoint$MetricsHourPrimaryTransactionsBlob"
			$minuteUri = "$endpoint$MetricsMinutePrimaryTransactionsBlob"

			Write-Host "[INFO] Adding Storage Account Metric information for storage account"($account.Name)
        
			$sapmonPrivateConfig += @{ key = (($account.Name) + ".hour.key");value=$account.Key}
			$sapmonPrivateConfig += @{ key = (($account.Name) + ".minute.key");value=$account.Key}
        
			$sapmonPublicConfig += @{ key = (($account.Name) + ".hour.uri");value=$hourUri}
			$sapmonPublicConfig += @{ key = (($account.Name) + ".minute.uri");value=$minuteUri}
			$sapmonPublicConfig += @{ key = (($account.Name) + ".hour.name");value=$account.Name}
			$sapmonPublicConfig += @{ key = (($account.Name) + ".minute.name");value=$account.Name}
		}
		else
		{
			Write-Host "[INFO]"($account.Name)"is of type"($storage.AccountType)"- Storage Account Metrics are not available for Premium Type Storage."
			$sapmonPublicConfig += @{ key = (($account.Name) + ".hour.ispremium");value="1"}
			$sapmonPublicConfig += @{ key = (($account.Name) + ".minute.ispremium");value="1"}
		}
    }

	# Enable VM Diagnostics
    if (-not $DisableWAD)
    {
        Write-Host ("[INFO] Enabling Windows Azure Diagnostics for VM " + $selectedVM.Name)
        $wadstorage = $accounts[0]
        $selectedVM = Set-AzureVMDiagnosticsExtensionC -VM $selectedVM -DiagnosticsConfiguration $wadcfg -StorageAccountName $wadstorage.Name -StorageAccountKey $wadstorage.Key;
    
        $storage = Get-AzureStorageAccountWA -StorageAccountName $wadstorage.Name
        $endpoint = ($storage.Endpoints | where { $_ -like "*table*" })
        $wadUri = "$endpoint$wadTableName"

        $sapmonPrivateConfig += @{ key = "wad.key";value=$wadstorage.Key}
        $sapmonPublicConfig += @{ key = "wad.name";value=$wadstorage.Name}
        $sapmonPublicConfig += @{ key = "wad.isenabled";value="1"}
        $sapmonPublicConfig += @{ key = "wad.uri";value=$wadUri}
    }
    else
    {
        $sapmonPublicConfig += @{ key = "wad.isenabled";value="0"}
    }
    
    $jsonPublicConfig = @{}
    $jsonPublicConfig.cfg = $sapmonPublicConfig
    $publicConfString = ConvertTo-Json $jsonPublicConfig

    $jsonPrivateConfig = @{}
    $jsonPrivateConfig.cfg = $sapmonPrivateConfig
    $privateConfString = ConvertTo-Json $jsonPrivateConfig  
    
    $selectedVM = Set-AzureVMExtension -ExtensionName $sapmonitoringextName -Publisher $sapmonitoringextPublisher -VM $selectedVM -PrivateConfiguration $privateConfString -PublicConfiguration $publicConfString -Version "2.*"
    Write-Host "[INFO] Updating Azure Enhanced Monitoring Extension for SAP configuration - Please wait..."
    $selectedVM = Update-AzureVM -Name $selectedVM.Name -VM $selectedVM.VM -ServiceName $selectedVM.ServiceName
    Write-Host "[INFO] Azure Enhanced Monitoring Extension for SAP configuration updated. It can take up to 15 Minutes for the monitoring data to appear in the SAP system."
    Write-Host "[INFO] You can check the configuration of a virtual machine by calling the Test-VMConfigForSAP_GUI commandlet."
}

<# 
 .Synopsis
  Checks the configuration of a virtual machine that should be enabled for monitoring.

 .Description  
  This commandlet will check the configuration of the extension that collects the performance data and if performance data is available. 

 .Example
  Test-VMConfigForSAP_GUI
#>
function Test-VMConfigForSAP_GUI
{
    param
    (
    )

    Select-Subscription

    $selectedVM = Select-VM
    if (-not $selectedVM)
    {
        return
    }

    Test-VMConfigForSAP -VMName $selectedVM.Name -VMServiceName $selectedVM.ServiceName
}

<# 
 .Synopsis
  Checks the configuration of a virtual machine that should be enabled for monitoring.

 .Description  
  This commandlet will check the configuration of the extension that collects the performance data and if performance data is available. 
  
 .Parameter VMName
  The name of the virtual machine that should be enable for monitoring.

 .Parameter VMServiceName
  The name of the cloud service that the virtual machine is part of.

 .Parameter ContentAgeInMinutes
  Defines how old the performance data is allowed to be.

 .Example
  Test-VMConfigForSAP -VMName SAPVM -VMServiceName SAPLandscape
#>
function Test-VMConfigForSAP
{	
    param
    (
        [Parameter(Mandatory=$True)] $VMName,
        [Parameter(Mandatory=$True)] $VMServiceName,
        $ContentAgeInMinutes = 5
    )

    $OverallResult = $true

    #################################################
    # Check if VM exists
    #################################################
    Write-Host "VM Existance check for $VMName ..." -NoNewline
    $selectedVM = Get-AzureVM -ServiceName $VMServiceName -Name $VMName
    $selectedRole = Get-AzureRole -ServiceName $VMServiceName -RoleName $VMName -Slot Production
    if (-not $selectedVM)
    {
        Write-Host "NOT OK " -ForegroundColor Red        
        return
    }
    else
    {
        Write-Host "OK " -ForegroundColor Green
    }
    #################################################    
    #################################################


    #################################################
    # Check for Guest Agent
    #################################################
    Write-Host "VM Guest Agent check..." -NoNewline
    if (-not $selectedVM.VM.ProvisionGuestAgent)
    {
        Write-Host "NOT OK " -ForegroundColor Red
        Write-Warning $missingGuestAgentWarning
        return
    }
    else
    {     
	    Write-Host "OK " -ForegroundColor Green
    }
    #################################################    
    #################################################


    #################################################
    # Check for Azure Enhanced Monitoring Extension for SAP
    #################################################
    Write-Host "Azure Enhanced Monitoring Extension for SAP Installation check..." -NoNewline
    $extensions = @(Get-AzureVMExtension -VM $selectedVM)
    $monExtension = $extensions | where { $_.ExtensionName -eq $sapmonitoringextName -and $_.Publisher -eq $sapmonitoringextPublisher }
    if (-not $monExtension -or [String]::IsNullOrEmpty($monExtension.PublicConfiguration))
    {
        Write-Host "NOT OK " -ForegroundColor Red
        $OverallResult = $false
    }
    else
    {
	    Write-Host "OK " -ForegroundColor Green
    }
    #################################################    
    #################################################


    $accounts = @()
    $osdisk = Get-AzureOSDisk -VM $selectedVM
	$disks = Get-AzureDataDisk -VM $selectedVM
    $accountName = Get-StorageAccountFromUri $osdisk.MediaLink    
    $osaccountName = $accountName
    $accounts += @{Name=$accountName}
    foreach ($disk in $disks)
    {
        $accountName = Get-StorageAccountFromUri $disk.MediaLink
        if (-not ($accounts | where Name -eq $accountName))
        {            
            $accounts += @{Name=$accountName}
        }
    }


    #################################################
    # Check storage metrics
    #################################################
    Write-Host "Storage Metrics check..."
    foreach ($account in $accounts)
    {
        Write-Host "`tStorage Metrics check for"$account.Name"..."
		$storage = Get-AzureStorageAccountWA -StorageAccountName $account.Name
		if ($storage.AccountType -like "Standard_*")
		{
			Write-Host "`t`tStorage Metrics configuration check for"$account.Name"..." -NoNewline
			$currentConfig = Get-StorageAnalytics -AccountName $account.Name

			if (-not (Check-StorageAnalytics $currentConfig))
			{            
				Write-Host "NOT OK " -ForegroundColor Red
				$OverallResult = $false
			}
			else
			{
				Write-Host "OK " -ForegroundColor Green
			}

			Write-Host "`t`tStorage Metrics data check for"$account.Name"..." -NoNewline
			$filterMinute =  [Microsoft.WindowsAzure.Storage.Table.TableQuery]::GenerateFilterConditionForDate("Timestamp", "gt", (get-date).AddMinutes($ContentAgeInMinutes * -1))
			if (Check-TableAndContent -StorageAccountName $account.Name -TableName $MetricsMinutePrimaryTransactionsBlob -FilterString $filterMinute -WaitChar ".")
			{
				Write-Host "OK " -ForegroundColor Green
			}
			else
			{            
				Write-Host "NOT OK " -ForegroundColor Red
				$OverallResult = $false
			}
		}
		else
		{
			Write-Host "`t`tStorage Metrics not available for Premium Storage account"$account.Name"..." -NoNewline
			Write-Host "OK " -ForegroundColor Green
		}
    }
    ################################################# 
    #################################################    

    
    #################################################
    # Check Azure Enhanced Monitoring Extension for SAP Configuration
    #################################################
    Write-Host "Azure Enhanced Monitoring Extension for SAP public configuration check..." -NoNewline
    if ($monExtension)
    {        
        Write-Host "" #New Line

        $sapmonPublicConfig = ConvertFrom-Json $monExtension.PublicConfiguration

        $storage = Get-AzureStorageAccountWA -StorageAccountName $osaccountName
		$osaccountIsPremium = ($storage.AccountType -notlike "Standard_*")
        $endpoint = ($storage.Endpoints | where { $_ -like "*table*" })
        $minuteUri = "$endpoint$MetricsMinutePrimaryTransactionsBlob"

        $vmSize = Get-VMSize -VM $selectedVM
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Size ..." -PropertyName "vmsize" -Properties $sapmonPublicConfig -ExpectedValue $vmSize
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Role Name ..." -PropertyName "vm.roleinstance" -Properties $sapmonPublicConfig -ExpectedValue $selectedVM.VM.RoleName
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Memory ..." -PropertyName "vm.memory.isovercommitted" -Properties $sapmonPublicConfig -ExpectedValue 0
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM CPU ..." -PropertyName "vm.cpu.isovercommitted" -Properties $sapmonPublicConfig -ExpectedValue 0
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: Deployment ID ..." -PropertyName "vm.deploymentid" -Properties $sapmonPublicConfig -ExpectedValue $selectedRole.DeploymentID
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: Script Version ..." -PropertyName "script.version" -Properties $sapmonPublicConfig
        
        $wadEnabled = Get-MonPropertyValue -PropertyName "wad.isenabled" -Properties $sapmonPublicConfig
        if ($wadEnabled.value -eq 1)
        {
            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: WAD name ..." -PropertyName "wad.name" -Properties $sapmonPublicConfig
            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: WAD URI ..." -PropertyName "wad.uri" -Properties $sapmonPublicConfig
        }
        else
        {
            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: WAD name ..." -PropertyName "wad.name" -Properties $sapmonPublicConfig -ExpectedValue $null
            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: WAD URI ..." -PropertyName "wad.uri" -Properties $sapmonPublicConfig -ExpectedValue $null
        }

        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM OS disk URI Key ..." -PropertyName "osdisk.connminute" -Properties $sapmonPublicConfig -ExpectedValue "$osaccountName.minute"
		if (-not $osaccountIsPremium)
		{
			$OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM OS disk URI Value ..." -PropertyName "$osaccountName.minute.uri" -Properties $sapmonPublicConfig -ExpectedValue $minuteUri
			$OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM OS disk URI Name ..." -PropertyName "$osaccountName.minute.name" -Properties $sapmonPublicConfig -ExpectedValue $osaccountName
		}
		else
		{
			Write-Host "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM OS disk Is Premium Storage Type ..." -NoNewLine
			Write-Host "OK " -ForegroundColor Green
		}
        $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM OS disk name ..." -PropertyName "osdisk.name" -Properties $sapmonPublicConfig -ExpectedValue ($osdisk.MediaLink.Segments[$osdisk.MediaLink.Segments.Count - 1])

        
        $diskNumber = 1
        foreach ($disk in $disks)
        {
            $accountName = Get-StorageAccountFromUri $disk.MediaLink
            $storage = Get-AzureStorageAccountWA -StorageAccountName $accountName
			$accountIsPremium = ($storage.AccountType -notlike "Standard_*")
            $endpoint = ($storage.Endpoints | where { $_ -like "*table*" })
            $minuteUri = "$endpoint$MetricsMinutePrimaryTransactionsBlob"

            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber LUN ..." -PropertyName "disk.lun.$diskNumber" -Properties $sapmonPublicConfig -ExpectedValue $disk.Lun
			$OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber URI Key ..." -PropertyName "disk.connminute.$diskNumber" -Properties $sapmonPublicConfig -ExpectedValue ($accountName + ".minute")
            if (-not $accountIsPremium)
			{				
				$OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber URI Value ..." -PropertyName ($accountName + ".minute.uri") -Properties $sapmonPublicConfig -ExpectedValue $minuteUri
				$OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber URI Name ..." -PropertyName ($accountName + ".minute.name") -Properties $sapmonPublicConfig -ExpectedValue $accountName
			}
			else
			{
				Write-Host "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber Is Premium Storage Type ..." -NoNewLine
				Write-Host "OK " -ForegroundColor Green
			}
            $OverallResult = Check-MonProp -CheckMessage "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disk $diskNumber name ..." -PropertyName "disk.name.$diskNumber" -Properties $sapmonPublicConfig -ExpectedValue ($disk.MediaLink.Segments[$disk.MediaLink.Segments.Count - 1])
            
            $diskNumber += 1
        }
        if ($disks.Count -eq 0)
        {
            Write-Host "`tAzure Enhanced Monitoring Extension for SAP public configuration check: VM Data Disks " -NoNewline
	        Write-Host "OK " -ForegroundColor Green
        }
    }
    else
    {
        Write-Host "NOT OK " -ForegroundColor Red
        $OverallResult = $false
    }
    ################################################# 
    #################################################    

    
    #################################################
    # Check WAD Configuration
    #################################################
    $wadEnabled = Get-MonPropertyValue -PropertyName "wad.isenabled" -Properties $sapmonPublicConfig      
    if ($wadEnabled -eq 1)
    {
        $extensions = @(Get-AzureVMExtension -VM $selectedVM)
        $wadExtension = $extensions | where { (($_.ExtensionName -eq $wadExtName) -and ($_.Publisher -eq  $wadPublisher))}
    
        Write-Host "Windows Azure Diagnostics check..." -NoNewline
        if ($wadExtension)
        {
            Write-Host "" #New Line
    
            Write-Host "`tWindows Azure Diagnostics configuration check..." -NoNewline

            $currentJSONConfig = ConvertFrom-Json ($wadExtension.PublicConfiguration)
            $base64 = $currentJSONConfig.xmlCfg
            [XML] $currentConfig = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))


            if (-not (Check-WADConfiguration -CurrentConfig $currentConfig))
            {
                Write-Host "NOT OK " -ForegroundColor Red            
                $OverallResult = $false
            }
            else
            {
	            Write-Host "OK " -ForegroundColor Green
            }

            Write-Host "`tWindows Azure Diagnostics performance counters check..."

            foreach ($perfCounter in $PerformanceCounters)
            {
                Write-Host "`t`tWindows Azure Diagnostics performance counters"($perfCounter.counterSpecifier)"check..." -NoNewline
		        $currentCounter = $currentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.PerformanceCounterConfiguration | where counterSpecifier -eq $perfCounter.counterSpecifier
                if ($currentCounter)
                {
	                Write-Host "OK " -ForegroundColor Green                            
                }
                else
                {
                    Write-Host "NOT OK " -ForegroundColor Red            
                    $OverallResult = $false
                }
            }

            $wadstorage = Get-MonPropertyValue -PropertyName "wad.name" -Properties $sapmonPublicConfig

            Write-Host "`tWindows Azure Diagnostics data check..." -NoNewline
            $filterMinute =  [Microsoft.WindowsAzure.Storage.Table.TableQuery]::GenerateFilterCondition("PartitionKey", "gt", "0" + [DateTime]::UtcNow.AddMinutes($ContentAgeInMinutes * -1).Ticks)
            if ((-not [String]::IsNullOrEmpty($wadstorage)) -and (Check-TableAndContent -StorageAccountName $wadstorage -TableName $wadTableName -FilterString $filterMinute -RoleInstanceName $selectedVM.VM.RoleName -DeploymentId $selectedRole.DeploymentID -WaitChar "."))
            {        
	            Write-Host "OK " -ForegroundColor Green
            }
            else
            {
                Write-Host "NOT OK " -ForegroundColor Red            
                $OverallResult = $false
            }
        }
        else
        {
            Write-Host "NOT OK " -ForegroundColor Red
            $OverallResult = $false
        }
    }
    ################################################# 
    #################################################

    if ($OverallResult -eq $false)
    {
        Write-Host "The script found some configuration issues. Please run the Update-VMConfigForSAP_GUI commandlet to update the configuration of the virtual machine!"
    }
}

function Enable-ProvisionGuestAgent_GUI
{
    param
    (
    )

    Select-Subscription

    $selectedVM = Select-VM
    if (-not $selectedVM)
    {
        return
    }

    Enable-ProvisionGuestAgent -VMName $selectedVM.Name -VMServiceName $selectedVM.ServiceName
}

function Enable-ProvisionGuestAgent
{
    param
    (
        [Parameter(Mandatory=$True)] $VMName,
        [Parameter(Mandatory=$True)] $VMServiceName
    )

    $selectedVM = Get-AzureVM -ServiceName $VMServiceName -Name $VMName

    if (-not $selectedVM)
    {
        $subName = (Get-AzureSubscription -Current).SubscriptionName
        Write-Error "No VM with name $VMName and Service Name $VMServiceName in subscription $subName found"
        return
    }
    if ($selectedVM.VM.ProvisionGuestAgent -eq $true)
    {       
        Write-Host "Guest Agent is already installed and enabled." -ForegroundColor Green
        return
    }

    Write-Host "This commandlet will enabled the Guest Agent on the Azure Virtual Machine. The Guest Agent needs to be installed on the Azure Virtual Machine. It will not be installed as part of this commandlet. Please read the documentation for more information"

    $selectedVM.VM.ProvisionGuestAgent = $TRUE
    Update-AzureVM –Name $VMName -VM $selectedVM.VM -ServiceName $VMServiceName
}


#######################################################################
## PRIVATE METHODS
#######################################################################

###
# Workaround for warnings: WARNING: GeoReplicationEnabled property will be deprecated in a future release of Azure PowerShell. The value will be merged into the AccountType property.
###
function Get-AzureStorageAccountWA
{
	param
	(
		$StorageAccountName,
		$ErrorAction = $ErrorActionPreference
	)
	
	$OldPreference = $WarningPreference
	$WarningPreference = "SilentlyContinue"
	Get-AzureStorageAccount -StorageAccountName $StorageAccountName -ErrorAction $ErrorAction
	$WarningPreference = $OldPreference
}

function Get-VMSize
{
    param
    (
        $VM
    )

    $vmsize = $VM.InstanceSize
    switch ($VM.InstanceSize)
    {
        "ExtraSmall" { $vmsize = "ExtraSmall (A0)" }
        "Small" { $vmsize = "Small (A1)" }
        "Medium" { $vmsize = "Medium (A2)" }
        "Large" { $vmsize = "Large (A3)" }
        "ExtraLarge" { $vmsize = "ExtraLarge (A4)" }
    }

    return $vmsize
}

function Get-MonPropertyValue
{
    param
    (
        $PropertyName,
        $Properties
    )

    $property = $Properties.cfg | where key -eq $PropertyName          
    return $property.value
}

function Check-MonProp
{
    param
    (
        $CheckMessage,
        $PropertyName,
        $Properties,
        $ExpectedValue
    )

    $value = Get-MonPropertyValue -PropertyName $PropertyName -Properties $Properties
    Write-Host $CheckMessage -NoNewline
    if ((-not [String]::IsNullOrEmpty($value) -and [String]::IsNullOrEmpty($ExpectedValue)) -or ($value -eq $ExpectedValue))
    {
        Write-Host "OK " -ForegroundColor Green
        return $true
    }
    else
    {
        Write-Host "NOT OK " -ForegroundColor Red
        return $false
    }
}

function Check-StorageAnalytics
{
    param
    (
        [XML] $CurrentConfig
    )    

    if (    (-not $CurrentConfig) `
        -or (-not $CurrentConfig.StorageServiceProperties) `
        -or (-not $CurrentConfig.StorageServiceProperties.Logging) `
        -or (-not [bool]::Parse($CurrentConfig.StorageServiceProperties.Logging.Read)) `
        -or (-not [bool]::Parse($CurrentConfig.StorageServiceProperties.Logging.Write)) `
        -or (-not [bool]::Parse($CurrentConfig.StorageServiceProperties.Logging.Delete)) `
        -or (-not $CurrentConfig.StorageServiceProperties.MinuteMetrics) `
        -or (-not [bool]::Parse($CurrentConfig.StorageServiceProperties.MinuteMetrics.Enabled)) `
        -or (-not $CurrentConfig.StorageServiceProperties.MinuteMetrics.RetentionPolicy) `
        -or (-not [bool]::Parse($CurrentConfig.StorageServiceProperties.MinuteMetrics.RetentionPolicy.Enabled)) `
        -or (-not $CurrentConfig.StorageServiceProperties.MinuteMetrics.RetentionPolicy.Days) `
        -or ([int]::Parse($CurrentConfig.StorageServiceProperties.MinuteMetrics.RetentionPolicy.Days) -lt 0))
        
    {
        return $false
    }

    return $true
}

function Check-TableAndContent
{
    param
    (
        $StorageAccountName,
        $TableName,
        $FilterString,
        $RoleInstanceName,
        $DeploymentId,
        $TimeoutinMinutes = 5,
        $WaitChar
    )

    $tableExists = $false

    $account = $null
    if (-not [String]::IsNullOrEmpty($StorageAccountName))
    {
        $account = Get-AzureStorageAccountWA $StorageAccountName -ErrorAction Ignore
    }
    if ($account)
    {
        $endpoint = Get-CoreEndpoint -StorageAccountName $StorageAccountName
        $keys = Get-AzureStorageKey -StorageAccountName $StorageAccountName
        $context = New-AzureStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $keys.Primary -Endpoint $endpoint

        $checkStart = Get-Date
        $wait = $true
        $table = Get-AzureStorageTable -Name $TableName -Context $context -ErrorAction SilentlyContinue
        while ($wait)
        {
            if ($table)
            {
                $query = new-object Microsoft.WindowsAzure.Storage.Table.TableQuery                
                $query.FilterString =  $FilterString
                $results = @($table.CloudTable.ExecuteQuery($query))
                if (($results.Count -gt 0) -and (-not [String]::IsNullOrEmpty($RoleInstanceName)))
                {
                    Write-Verbose "Filtering table results"
                    $results = $results | where { ($_.Properties.RoleInstance.StringValue -eq $RoleInstanceName) -and ($_.Properties.Role.StringValue -eq $ROLECONTENT) -and ($_.Properties.DeploymentId.StringValue -eq $DeploymentId) }
                }

                if ($results.Count -gt 0)
                {            
                    $tableExists = $true
                    break                
                }
                else
                {
                }
            }
            else
            {
            }

            Write-Host $WaitChar -NoNewline
            sleep 5
            $table = Get-AzureStorageTable -Name $TableName -Context $context -ErrorAction SilentlyContinue
            $wait = ((Get-Date) - $checkStart).TotalMinutes -lt $TimeoutinMinutes
        }
    }
    return $tableExists
}


function Get-StorageAccountFromUri
{
    param
    (
        $URI
    )

    if ($URI.Host -match "(.*?)\..*")
    {
        return $Matches[1]        
    }
    else
    {
        Write-Error "Could not determine storage account for OS disk. Please contact support"
        return
    }
}

function Check-WADConfiguration
{
    param
    (
        [XML] $CurrentConfig
    )

    if ( `
            (-not $CurrentConfig) `
            -or (-not $CurrentConfig.WadCfg) `
            -or (-not $CurrentConfig.WadCfg.DiagnosticMonitorConfiguration) `
            -or ([int]::Parse($CurrentConfig.WadCfg.DiagnosticMonitorConfiguration.Attributes["overallQuotaInMB"].Value) -lt 4096) `
            -or (-not $CurrentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters) `
            -or ($CurrentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.Attributes["scheduledTransferPeriod"].Value -ne "PT1M") `
            -or (-not $CurrentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.PerformanceCounterConfiguration) `
            )
    {
        return $false      
    }

    return $true
}

function Set-AzureVMDiagnosticsExtensionC
{
    param
    (
        $VM,
        $StorageAccountName,
        $StorageAccountKey
    )   

    $sWADPublicConfig = [String]::Empty
    $sWADPrivateConfig = [String]::Empty
    
    $publicConf = (Get-AzureVMExtension -ExtensionName $wadExtName -Publisher $wadPublisher -VM $VM -WarningAction SilentlyContinue).PublicConfiguration
    if (-not [String]::IsNullOrEmpty($publicConf))
    {
        $currentJSONConfig = ConvertFrom-Json ($publicConf)
        $base64 = $currentJSONConfig.xmlCfg
        [XML] $currentConfig = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
    }
    $xmlnsConfig = "http://schemas.microsoft.com/ServiceHosting/2010/10/DiagnosticsConfiguration"

    if ($currentConfig.WadCfg.DiagnosticMonitorConfiguration.DiagnosticInfrastructureLogs -and 
        $currentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters)
    {
        $currentConfig.WadCfg.DiagnosticMonitorConfiguration.overallQuotaInMB = "4096"
        $currentConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.scheduledTransferPeriod = "PT1M"           

        $publicConfig = $currentConfig

    }
    else
    {    
        $publicConfig = $WADPublicConfig
    }
    $publicConfig = $WADPublicConfig
        
    foreach ($perfCounter in $PerformanceCounters)
    {
		$currentCounter = $publicConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.PerformanceCounterConfiguration | where counterSpecifier -eq $perfCounter.counterSpecifier
        if (-not $currentCounter)
        {
            $node = $publicConfig.CreateElement("PerformanceCounterConfiguration", $xmlnsConfig)
            $nul = $publicConfig.WadCfg.DiagnosticMonitorConfiguration.PerformanceCounters.AppendChild($node)    
            $node.SetAttribute("counterSpecifier", $perfCounter.counterSpecifier)
            $node.SetAttribute("sampleRate", $perfCounter.sampleRate)
        }
    }
    
    $Endpoint = Get-CoreEndpoint $StorageAccountName
    $Endpoint = "https://$Endpoint"

    $jPublicConfig = @{}
    $jPublicConfig.xmlCfg = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($publicConfig.InnerXml))
    
    $jPrivateConfig = @{}    
    $jPrivateConfig.storageAccountName = $StorageAccountName
    $jPrivateConfig.storageAccountKey = $StorageAccountKey
    $jPrivateConfig.storageAccountEndPoint = $Endpoint

    $sWADPublicConfig = ConvertTo-Json $jPublicConfig
    $sWADPrivateConfig = ConvertTo-Json $jPrivateConfig
    $VM = Set-AzureVMExtension -ExtensionName $wadExtName -Publisher $wadPublisher -PublicConfiguration $sWADPublicConfig -VM $VM -PrivateConfiguration $sWADPrivateConfig -Version "1.*"
    return $VM
}

function Select-Subscription
{
    $selectedEnv = $null
    $envs = @(Get-AzureEnvironment)
    if ($envs.Count -gt 1)
    {
        write-host "Please select one of the following environments. Make sure to select the correct environment, especially if you want to use Microsoft Azure in China."
        $currentEnvIndex = 1
	    foreach ($currentEnv in $envs)
	    {
		    write-host ("[$currentEnvIndex] " + (Get-EnvironmentName $currentEnv))
		    $currentEnvIndex += 1
	    }

	    $selectedEnv = $null
	    while (-not $selectedEnv)
	    {
		    [int] $index = Read-Host -Prompt ("Select Environment [1-" + $envs.Count + "]")
		    if (($index -ge 1) -and ($index -le $envs.Count))
		    {
			    $selectedEnv = $envs[$index - 1]
		    }
	    }
    }
    elseif ($envs.Count -eq 1)
    {
        $selectedEnv = $envs[1]
    }

    if ($selectedEnv)
    {    
	    Add-AzureAccount -Environment (Get-EnvironmentName $selectedEnv)
    }
    else
    {
        Add-AzureAccount
    }
    
	# Select subscription
	$subscriptions = Get-AzureSubscription
    if ($subscriptions.Length -gt 1)
    {
	    $currentIndex = 1
	    foreach ($currentSub in $subscriptions)
	    {
		    write-host ("[$currentIndex] " + $currentSub.SubscriptionName)
		    $currentIndex += 1
	    }

	    $selectedSubscription = $null
	    while (-not $selectedSubscription)
	    {
		    [int] $index = Read-Host -Prompt ("Select Subscription [1-" + $subscriptions.Count + "]")
		    if (($index -ge 1) -and ($index -le $subscriptions.Count))
		    {
			    $selectedSubscription = $subscriptions[$index - 1]
		    }
	    }
    }
    elseif ($subscriptions.Length -eq 1)
    {
        $selectedSubscription = $subscriptions[0]
    }

	$selectedSubscription | Select-AzureSubscription
}

function Select-VM
{
    # Select VM
    $selectedVM = $null
    $vmFilterName = Read-Host -Prompt "Please enter the name of the VM or a filter you want to use to select the VM"

    #if (-not [String]::IsNullOrEmpty($vmFilterName))
    #{
    #    Write-Verbose "Using $vmFilterName to get a single VM"
    #    $selectedVM = Get-AzureVM -Name $vmFilterName -ServiceName $vmFilterName -ErrorAction SilentlyContinue
    #}
    #if (-not $selectedVM)
    #{

    Write-Host "`tRetrieving information about virtual machines in your subscription. Please wait..."
	$vms = Get-AzureVM | where Name -like ("*$vmFilterName*")
    
    if ($vms.Count -gt 0)
    {
	    $currentIndex = 1
	    foreach ($currentVM in $vms)
	    {
		    write-host ("[$currentIndex] " + $currentVM.Name + " (part of cloud service " + $currentVM.ServiceName + ")")
		    $currentIndex += 1
	    }

	    $selectedVM = $null
	    while (-not $selectedVM)
	    {
		    [int] $index = Read-Host -Prompt ("Select Virtual Machine [1-" + $vms.Count + "]")
		    if (($index -ge 1) -and ($index -le $vms.Count))
		    {
			    $selectedVM = $vms[$index - 1]
		    }
	    }
    }        
    else
    {
        Write-Warning "No Virtual machine found that matches $vmFilterName"
        return $null
    }

    #}

    return $selectedVM
}

function Get-StorageAnalytics
{
	param 
   	(
		[Parameter(Mandatory = $true)]
		[string] $AccountName
	)

    [XML]$resultXML = $null
    #-xmsversion "2013-08-15"
    $request = Create-StorageAccountRequest -accountName $AccountName -resourceType "blob.core" -operationString "?restype=service&comp=properties" -xmsversion "2014-02-14" -contentLength "" -restMethod "GET"
    
    try
    {
        $response = $request.GetResponse()
        if ($response.Headers.Count -gt 0)
        {
            # Parse the web response.
            $reader = new-object System.IO.StreamReader($response.GetResponseStream())
            [XML]$resultXML = $reader.ReadToEnd()
        }
         
        # Close the resources no longer needed.
        $response.Close()
        $reader.Close()
    }
    catch [System.Net.WebException]
    {
        $_.Exception.ToString()
        $data = $_.Exception.Response.GetResponseStream()
        $reader = new-object System.IO.StreamReader($data)
        $text = $reader.ReadToEnd();
        throw $_.Exception.ToString()
    }

    return $resultXML
}

function Set-StorageAnalytics
{
    [CmdletBinding()]
	param 
   	(
		[Parameter(Mandatory = $true)]
		[string] $AccountName,

		[Parameter(Mandatory = $true)]
        [XML] $StorageServiceProperties
	)

    $requestBody = $StorageServiceProperties.InnerXml
    $request = Create-StorageAccountRequest -accountName $AccountName -resourceType "blob.core" -operationString "?restype=service&comp=properties" -xmsversion "2013-08-15" -contentLength $requestBody.Length -restMethod "PUT"
    try
    {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($requestBody)
        $streamOut = $request.GetRequestStream()
	    $streamOut.Write($bytes, 0, $bytes.Length)
	    $streamOut.Flush()
	    $streamOut.Close()

        $response = $request.GetResponse()
        if ($response.Headers.Count -gt 0)
        {
            # Parse the web response.
            $reader = new-object System.IO.StreamReader($response.GetResponseStream())
            $resultXML = $reader.ReadToEnd()
        }
         
        # Close the resources no longer needed.
        $response.Close()
        $reader.Close()
    }
    catch [System.Net.WebException]
    {
        $_.Exception.ToString()
        $data = $_.Exception.Response.GetResponseStream()
        $reader = new-object System.IO.StreamReader($data)
        $text = $reader.ReadToEnd();
        $text
        throw $_.Exception.ToString()
    }
}

function Create-StorageAccountRequest
{
	[CmdletBinding()]
	param 
   	(
		[Parameter(Mandatory = $true)]
		[string] $accountName,

		[Parameter(Mandatory = $true)]
		[string] $resourceType,

		[Parameter(Mandatory = $true)]
		[string] $operationString,

		[Parameter(Mandatory = $true)]
		[string] $xmsversion,

		[string] $contentLength,

		[Parameter(Mandatory = $true)]
		[string] $restMethod
	)

	$nl   = [char]10 # newLine
    $date = (Get-Date).ToUniversalTime().ToString("R")

    
    $storage = Get-AzureStorageAccountWA -StorageAccountName $accountName
    $azureTableEndpoint = Get-Endpoint -StorageAccountName $accountName
    
    $keys = Get-AzureStorageKey -StorageAccountName $accountName

    [String] $azureHostString = [String]::Format("{0}.{1}.$azureTableEndpoint/", $accountName, $resourceType)
    [String] $azureUriString = [String]::Format("https://{0}{1}{2}", $azureHostString, $subscriptionId, $operationString)
    [Uri] $uri = New-Object System.Uri($azureUriString)	


    [String] $canonicalizedHeadersString = [String]::Format("{0}{1}{1}{1}{4}{1}{1}{1}{1}{1}{1}{1}{1}{1}x-ms-date:{2}{1}x-ms-version:{3}{1}",$restMethod, $nl, $date, $xmsversion, $contentLength)
	[String] $canonicalizedResourceString = [String]::Format("/{1}/{0}comp:properties{0}restype:service" ,$nl, $accountName)
	[String] $signatureString = $canonicalizedHeadersString + $canonicalizedResourceString

		# Encodes this string by using the HMAC-SHA256 algorithm and constructs the authorization header.
	[Byte[]] $unicodeKeyByteArray      = [System.Convert]::FromBase64String($keys.Primary)
	$hmacSha256               = new-object System.Security.Cryptography.HMACSHA256((,$unicodeKeyByteArray))
    
	# Encode the signature.
	[Byte[]] $signatureStringByteArray = [System.Text.Encoding]::UTF8.GetBytes($signatureString)
    [String] $signatureStringHash      = [System.Convert]::ToBase64String($hmacSha256.ComputeHash($signatureStringByteArray))
    
	# Build the authorization header.
    [String] $authorizationHeader = [String]::Format([CultureInfo]::InvariantCulture,"{0} {1}:{2}", "SharedKey", $accountName, $signatureStringHash)


    # Create the request and specify attributes of the request.
    [System.Net.HttpWebRequest] $request = [System.Net.HttpWebRequest]::Create($uri)
         
    # Define the requred headers to specify the API version and operation type.
    $request.Headers.Add('x-ms-version', $xmsversion)
    $request.Method            = $restMethod
    #$request.ContentType       = $contentType
    #$request.Accept            = $contentType
    $request.AllowAutoRedirect = $false
    $request.ServicePoint.Expect100Continue = $false
    $request.Headers.Add("Authorization", $authorizationHeader)
	$request.Headers.Add("x-ms-date", $date)

    return $request
}

function Get-Endpoint
{
    param
    (
        $StorageAccountName
    )

    $storage = Get-AzureStorageAccountWA -StorageAccountName $StorageAccountName
    $tableendpoint = ($storage.Endpoints | where { $_ -like "*table*" })
    $blobendpoint = ($storage.Endpoints | where { $_ -like "*blob*" })
    if ($tableendpoint -match "http://.*?\.table\.core\.(.*)/")
    {
        $azureTableEndpoint = $Matches[1]
    }
    elseif ($tableendpoint -match "https://.*?\.table\.core\.(.*)/")
    {
        $azureTableEndpoint = $Matches[1]
    }
    elseif ($blobendpoint -match "http://.*?\.blob\.core\.(.*)/")
    {
        $azureTableEndpoint = $Matches[1]
    }
    elseif ($blobendpoint -match "https://.*?\.blob\.core\.(.*)/")
    {
        $azureTableEndpoint = $Matches[1]
    }
    else
    {
        Write-Warning "Could not extract endpoint information from Azure Storage Account. Using default $AzureEndpoint"
        $azureTableEndpoint = $AzureEndpoint
    }
    return  $azureTableEndpoint
}

function Get-CoreEndpoint
{
    param
    (
        $StorageAccountName
    )

    $azureTableEndpoint = Get-Endpoint -StorageAccountName $StorageAccountName
    return ("core." + $azureTableEndpoint)
}

function Get-EnvironmentName
{
    param
    (
        $Environment
    )

    if ($Environment | Get-Member -Name EnvironmentName) 
    {
        return $Environment.EnvironmentName
    }
    else
    {
        return $Environment.Name
    }
}

$ErrorActionPreference = "Stop"
$CurrentScriptVersion = "1.2.0.1"
$missingGuestAgentWarning = "Provision Guest Agent is not installed on this Azure Virtual Machine. Please read the documentation on how to download and install the Provision Guest Agent. After you have installed the Provision Guest Agent, enable it with the Enable-ProvisionGuestAgent_GUI commandlet that is part of this Powershell Module."

$sapmonitoringextPublisher = "Microsoft.AzureCAT.AzureEnhancedMonitoring"
$sapmonitoringextName = "AzureCATExtensionHandler"
$AzureEndpoint = "windows.net"
$wadExtName = "IaaSDiagnostics"
$wadPublisher = "Microsoft.Azure.Diagnostics"
$ROLECONTENT = "IaaS"

$PerformanceCounters = @(
                 @{"counterSpecifier"="\Processor(_Total)\% Processor Time";"sampleRate" = "PT1M"}
                @{"counterSpecifier"="\Processor Information(_Total)\Processor Frequency";"sampleRate"="PT1M"}
		        @{"counterSpecifier"="\Memory\Available Bytes";"sampleRate"="PT1M"}
		        @{"counterSpecifier"="\TCPv6\Segments Retransmitted/sec";"sampleRate"="PT1M"}
		        @{"counterSpecifier"="\TCPv4\Segments Retransmitted/sec";"sampleRate"="PT1M"}		        
		        @{"counterSpecifier"="\Network Interface(*)\Bytes Sent/sec";"sampleRate"="PT1M"}
		        @{"counterSpecifier"="\Network Interface(*)\Bytes Received/sec";"sampleRate"="PT1M"}		       	       
            )
				




[XML] $WADPublicConfig = @"    
    <WadCfg>
        <DiagnosticMonitorConfiguration overallQuotaInMB="4096">
			<PerformanceCounters scheduledTransferPeriod="PT1M" >
			</PerformanceCounters>				
		</DiagnosticMonitorConfiguration>
    </WadCfg>    
"@

$wadTableName = "WADPerformanceCountersTable"
$WADTableName = $wadTableName

[xml]$DefaultStorageAnalyticsConfig = @'
<StorageServiceProperties>
  <Logging>
    <Version>1.0</Version>
    <Delete>true</Delete>
    <Read>true</Read>
    <Write>true</Write>
    <RetentionPolicy>
      <Enabled>true</Enabled>
      <Days>12</Days>
    </RetentionPolicy>
  </Logging>
  <HourMetrics>
    <Version>1.0</Version>
    <Enabled>true</Enabled>
    <IncludeAPIs>true</IncludeAPIs>
    <RetentionPolicy>
      <Enabled>true</Enabled>
      <Days>13</Days>
    </RetentionPolicy>
  </HourMetrics>
  <MinuteMetrics>
    <Version>1.0</Version>
    <Enabled>true</Enabled>
    <IncludeAPIs>true</IncludeAPIs>
    <RetentionPolicy>
      <Enabled>true</Enabled>
      <Days>13</Days>
    </RetentionPolicy>
  </MinuteMetrics>
  <Cors />
</StorageServiceProperties>
'@

$MetricsHourPrimaryTransactionsBlob = "`$MetricsHourPrimaryTransactionsBlob"
$MetricsMinutePrimaryTransactionsBlob = "`$MetricsMinutePrimaryTransactionsBlob"

Export-ModuleMember -Function Update-VMConfigForSAP
Export-ModuleMember -Function Update-VMConfigForSAP_GUI
Export-ModuleMember -Function Test-VMConfigForSAP
Export-ModuleMember -Function Test-VMConfigForSAP_GUI
Export-ModuleMember -Function Enable-ProvisionGuestAgent_GUI
Export-ModuleMember -Function Enable-ProvisionGuestAgent
