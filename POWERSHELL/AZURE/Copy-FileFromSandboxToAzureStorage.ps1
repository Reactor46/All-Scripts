<#
.SYNOPSIS 
    Copies a file from Automation service sandbox where the runbook is executed, to the storage blob. 

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    and then uploads the file from local sandbox to storage account.
     

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

.PARAMETER SandboxFilePath
    Absolute path to the local sandbox file.

.EXAMPLE
    Copy-FileFromSandboxToAzureStorage -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" -ContainerName "MyContainer" `
    -BlobName "MyBlob" -SandboxFilePath "C:\Temp\TemFile"

.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>

workflow Copy-FileFromSandboxToAzureStorage {
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
        $BlobName,

        [parameter(Mandatory=$true)]
        [String]
        $SandboxFilePath

    )

    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName
    
    Write-Verbose "Uploading $SandboxFilePath to Azure Blob container $ContainerName $BlobName"
    $BlobContent = Set-AzureStorageBlobContent -Container $ContainerName -File $SandboxFilePath -Blob $BlobName -Force
    $cscfgNewAbsoluteUri = "https://" + $StorageAccountName + ".blob.core.windows.net/" + $ContainerName + "/" + $BlobName
    Write-Output $cscfgNewAbsoluteUri
}