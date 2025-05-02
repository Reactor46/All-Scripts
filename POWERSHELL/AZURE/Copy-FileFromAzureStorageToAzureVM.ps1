workflow Copy-FileFromAzureStorageToAzureVM {
    param
    (
        [parameter(Mandatory=$true)]
        [String]
        $AzureConnectionName,

        [parameter(Mandatory=$true)]
        [String]
        $CredentialAssetNameWithAccessToVM,

        [parameter(Mandatory=$true)]
        [String]
        $StorageAccountName,

        [parameter(Mandatory=$true)]
        [String]
        $ContainerName,

        [parameter(Mandatory=$true)]
        [String]
        $BlobName,

        [parameter(Mandatory=$true)]
        [String]
        $PathToPlaceFile,

        [parameter(Mandatory=$true)]
        [object]
        $VM
    )

    $TempFileLocation = "C:\$BlobName"

    Connect-Azure -AzureConnectionName $AzureConnectionName

    Write-Verbose "Downloading $BlobName from Azure Blob Storage to $TempFileLocation"

    InlineScript {
        Select-AzureSubscription -SubscriptionName $Using:AzureConnectionName
        
        $StorageAccount = (Get-AzureStorageAccount -StorageAccountName $Using:StorageAccountName).Label
                
        Set-AzureSubscription `
            -SubscriptionName $Using:AzureConnectionName `
            -CurrentStorageAccount $StorageAccount

        $blob = 
            Get-AzureStorageBlobContent `
                -Blob $Using:BlobName `
                -Container $Using:ContainerName `
                -Destination $Using:TempFileLocation `
                -Force
    }

    Write-Verbose ("Copying $BlobName to $PathToPlaceFile on " + $VM.Name)
        
    Copy-ItemToAzureVM `
        -AzureConnectionName $AzureConnectionName `
        -ServiceName $VM.ServiceName `
        -VMName $VM.Name `
        -VMCredentialName $CredentialAssetNameWithAccessToVM `
        -LocalPath $TempFileLocation `
        -RemotePath $PathToPlaceFile
}