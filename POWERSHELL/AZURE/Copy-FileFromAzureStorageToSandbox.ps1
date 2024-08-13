<#
.SYNOPSIS 
    Copies a file from Azure storage account to Automation service sandbox where the runbook is executed. It returns
    the copied location.

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    and then downloads the file from storage account to local sandbox.
     

.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER ContainerName
    Name of the Azure storage container for build packages.

.PARAMETER BlobName
    Name of the blob file.

.EXAMPLE
    Copy-FileFromAzureStorageToSandbox -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" -ContainerName "MyContainer" `
    -BlobName "MyBlob"

.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>

workflow Copy-FileFromAzureStorageToSandbox {
    param
    (
        [parameter(Mandatory=$true)]
        [String]
        $AzureConnectionName,

        [parameter(Mandatory=$true)]
        [String]
        $StorageAccountName,

        [parameter(Mandatory=$true)]
        [String]
        $ContainerName,

        [parameter(Mandatory=$true)]
        [String]
        $BlobName

    )

    $TempFileLocation = "C:\temp\$BlobName"

    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    Write-Verbose "Downloading $BlobName from Azure Blob Storage to $TempFileLocation"

    $blob = 
          Get-AzureStorageBlobContent `
                -Blob $BlobName `
                -Container $ContainerName `
                -Destination $TempFileLocation `
                -Force

    Write-Output $TempFileLocation   
}