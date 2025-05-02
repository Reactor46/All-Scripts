workflow Update-AzureVM {
    $AzureConnectionName = "joeAzure"
    
    $CredentialAssetNameWithAccessToAllVMs = "joeAzureVMCred"
    $CredentialWithAccessToAllVMs = Get-AutomationPSCredential -Name $CredentialAssetNameWithAccessToAllVMs

    $WUModuleStorageAccountName = "joestorage123"
    $WUModuleContainerName = "psmodules"
    $WUModuleBlobName = "PSWindowsUpdate.zip"

    $ModuleName = "PSWindowsUpdate"

    Connect-Azure -AzureConnectionName $AzureConnectionName

    Write-Verbose "Getting all VMs in $AzureConnectionName"

    # Get all VMs in subscription
    $VMs = InlineScript {
        Select-AzureSubscription -SubscriptionName $Using:AzureConnectionName
        Get-AzureVM
    }

    # Install PSWindowsUpdate module on each VM if it is not installed already
    foreach($VM in $VMs) {
        Write-Verbose ("Installing $ModuleName module on " + $VM.Name + " if it is not installed already")
        
        Install-ModuleOnAzureVM `
            -AzureConnectionName $AzureConnectionName `
            -CredentialAssetNameWithAccessToVM $CredentialAssetNameWithAccessToAllVMs `
            -ModuleStorageAccountName $WUModuleStorageAccountName `
            -ModuleContainerName $WUModuleContainerName `
            -ModuleBlobName $WUModuleBlobName `
            -VM $VM `
            -ModuleName $ModuleName
    }

    # Install latest WU updates onto each VM
    foreach($VM in $VMs) {
        $ServiceName = $VM.ServiceName
        $VMName = $VM.Name
        
        $Uri = Connect-AzureVM -AzureConnectionName $AzureConnectionName -ServiceName $ServiceName -VMName $VMName
        
        Write-Verbose "Installing latest Windows Update updates on $VMName"

        InlineScript {
            Invoke-Command -ConnectionUri $Using:Uri -Credential $Using:CredentialWithAccessToAllVMs -ScriptBlock {
                $Updates = Get-WUList -WindowsUpdate | Select-Object Title, KB, Size, MoreInfoUrls, Categories

                foreach($Update in $Updates) {
                    $Output = @{
                        "KB" = $Update.KB
                        "Size" = $Update.Size
                        "Category1" = ($Update.Categories | Select-Object Description).Description[0]
                        "Category2" = ($Update.Categories | Select-Object Description).Description[1]
                    }

                    "Title: " + $Update.Title
                    $Output
                    "More info at: " + $Update.MoreInfoUrls[0]
                    "------------------------------------------------"
                }
            }
        }
    }
}