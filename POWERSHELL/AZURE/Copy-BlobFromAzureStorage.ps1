<#
.SYNOPSIS 
    Downloads an Azure blob to the host running the runbook job.

.DESCRIPTION
    This runbook downloads an Azure blob to the host running the runbook job.

    In addition to this runbook, you will need an OrgID credential with access to your Azure subscription
    (http://azure.microsoft.com/blog/2014/08/27/azure-automation-authenticating-to-azure-using-azure-active-directory/)
    stored in a credential asset.

	When using this runbook, be aware that the memory and disk space size of the processes running your
	runbooks is limited. Because of this, we recommened only using runbooks to transfer small files.
	All Automation Integration Module assets in your account are loaded into your processes,
	so be aware that the more Integration Modules you have in your system, the smaller the free space in
	your processes will be. To ensure maximum disk space in your processes, make sure to clean up any local
	files a runbook transfers or creates in the process before the runbook completes.

.PARAMETER AzureSubscriptionName 
    Name of the Azure subscription to connect to 
     
.PARAMETER AzureOrgIdCredential 
    A credential containing an Org Id username / password with access to this Azure subscription. 
 
    If invoking this runbook inline from within another runbook, pass a PSCredential for this parameter. 
 
    If starting this runbook using Start-AzureAutomationRunbook, or via the Azure portal UI, pass as a string the 
    name of an Azure Automation PSCredential asset instead. Azure Automation will automatically grab the asset with 
    that name and pass it into the runbook. 
#>
workflow Copy-BlobFromAzureStorage {
    param
    (
        [parameter(Mandatory=$True)] 
        [String] 
        $AzureSubscriptionName, 
 
        [parameter(Mandatory=$True)] 
        [PSCredential] 
        $AzureOrgIdCredential, 

        [parameter(Mandatory=$True)]
        [String]
        $StorageAccountName,

        [parameter(Mandatory=$True)]
        [String]
        $ContainerName,

        [parameter(Mandatory=$True)]
        [String]
        $BlobName,

        [parameter(Mandatory=$False)]
        [String]
        $PathToPlaceBlob = "C:\"
    )

    $Null = Add-AzureAccount -Credential $AzureOrgIdCredential 
    $Null = Select-AzureSubscription -SubscriptionName $AzureSubscriptionName

    Write-Verbose "Downloading $BlobName from Azure Blob Storage to $PathToPlaceBlob"

    Set-AzureSubscription `
        -SubscriptionName $AzureSubscriptionName `
        -CurrentStorageAccount $StorageAccountName

    $blob = 
        Get-AzureStorageBlobContent `
            -Blob $BlobName `
            -Container $ContainerName `
            -Destination $PathToPlaceBlob `
            -Force

    try {
        Get-Item -Path "$PathToPlaceBlob\$BlobName" -ErrorAction Stop
    }
    catch {
        Get-Item -Path $PathToPlaceBlob
    }
}