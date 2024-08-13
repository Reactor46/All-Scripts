<#
.SYNOPSIS 
    Deploys to an existing Azure cloud service production slot

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    and then checks the build location to see if the build is updated. If build is updated, it updates
    cscfg based on the encrypted automation variables.
    It then determines to deploy via in-place upgrade or vip swap, based on optional parameter $IsInplaceUpgrade. 
    If $IsInplaceUpgrade is not specified, by default it does in-place upgrade. It only does vip swap 
    when $IsInplaceUpgrade is set to $false.
    For in-place upgrade, it checks if there is an existing PROD deployment already. If yes, then does in-place upgrade. 
    Otherwise, deploy directly to the Production slot.
    For vip swap, it first deploys to staging slot, then performs vip swap. 
     

.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER ContainerName
    Name of the Azure storage container for build packages.

.PARAMETER CscfgBlobName
    Name of the cscfg file.

.PARAMETER CspkgBlobName
    Name of the cspkg file.

.PARAMETER CloudServiceName
    Name of the Azure cloud service to deploy.

.PARAMETER IsInplaceUpgrade
    whether to perform in-place upgrade or vip swap. If it is not specified, by default does in-place upgrade.

.PARAMETER CscfgVariableNamesToUpdate
    Array of Automation variable names which store the cscfg setting values, could be encrypted. 

.PARAMETER DeploymentLabel
    Optional. If not specified, it will be "Deployment" appended with time stamp.

.EXAMPLE
    Deploy-AzureCloudService -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" -ContainerName "MyContainer" `
    -CscfgBlobName "cloud.cscfg" -CspkgBlobName "cloud.cspkg" -CloudServiceName "MyCloudService" -CscfgVariableNamesToUpdate [ "MyCloudService#WebRole1#DBConnectionString", `
    "MyCloudService#WebRole1#MyCertThumbprint" ]


.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>

workflow Set-AzureCloudServiceDeployment
{
    param(
        [parameter(Mandatory=$true)]
        $AzureConnectionName,
        
        [parameter(Mandatory=$true)]
        [String]
        $StorageAccountName,

        [parameter(Mandatory=$true)]
        [String]
        $ContainerName,

        [parameter(Mandatory=$true)]
        [String]
        $CscfgBlobName,
        
        [parameter(Mandatory=$true)]
        [String]
        $CspkgBlobName,

        [parameter(Mandatory=$true)]
        $CloudServiceName,
                
        [parameter(Mandatory=$false)]
        [Boolean]
        $IsInplaceUpgrade = $true,

        # in the format of [$cloudservicename#rolename#variablename,...]
        [parameter(Mandatory=$false)]
        [String[]]
        $CscfgVariableNamesToUpdate,

        [parameter(Mandatory=$false)]
        [String]
        $DeploymentLabel
    )

    $ForceUpgrade = $true
    if ($IsInplaceUpgrade -eq $false)
    {
         $ForceUpgrade = $false
         Write-Verbose "$IsInplaceUpgrade is set to false, will deploy to staging first and then perform vipswap if the production deployment already exists"
    }

    if ($DeploymentLabel -eq $null)
    {
        $DeploymentLabel = "Deployment" + [System.DateTime]::Now.ToString("HH:mm:ss.ffffzzz")
    }

    $Slot="Production"

    # Call the Connect-Azure Runbook to set up the connection to Azure using the Automation connection asset
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName
    
    # check build status
    $BuildStatus = Check-AzureCloudServiceBuildStatus -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -ContainerName $ContainerName -CscfgBlobName $CscfgBlobName -CspkgBlobName $CspkgBlobName
    $BuildStatus
    $IsBuildUpdated = $BuildStatus.IsBuildUpdated
    $CspkgPath = $BuildStatus.cspkgAbsoluteUri
    $CscfgPath = $BuildStatus.cscfgAbsoluteUri

    if ($CspkgPath -eq $null) 
    {
       throw "Cspkg Path is undefined."
    }
    
    if ($CscfgPath -eq $null)
    {
       throw "Cscfg Path is undefined."
    }

    Checkpoint-Workflow
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    if($IsBuildUpdated -eq "True")
    {        
       $cscfgUpdatedPathLocal = Copy-FileFromAzureStorageToSandbox -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -ContainerName $ContainerName -BlobName $CscfgBlobName  
       $cscfgUpdatedPath = Update-Cscfg -ServiceName $CloudServiceName -CscfgSandboxPath $cscfgUpdatedPathLocal -VariableNamesToUpdate $CscfgVariableNamesToUpdate
       Write-Output ("Updated cscfg in sandbox: $cscfgUpdatedPath")
       $CscfgBlobUpdatedName = $CscfgBlobName + [System.DateTime]::Now.ToString("HH:mm:ss.ffffzzz")
       Copy-FileFromSandboxToAzureStorage -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -ContainerName $ContainerName -BlobName $CscfgBlobUpdatedName -SandboxFilePath $cscfgUpdatedPath
       Write-Output ("Updated cscfg in blob: $CscfgBlobUpdatedName")
       
       # check if the deployment exists
       $Deployment = Get-AzureDeployment -Slot $Slot -ServiceName $CloudServiceName -ErrorAction Ignore
       if (($Deployment -ne $null) -AND ($Deployment.DeploymentId  -ne $null))
       {
           if ($ForceUpgrade -eq $true)
           {
              Set-AzureCloudServiceInplaceUpgrade -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -ServiceName $CloudServiceName -Slot $Slot -NumberOfMinutesTillTimeOut 60 -PackagePath $CspkgPath -ConfigurationPath $cscfgUpdatedPath 
              Write-Output ("Upgrade successful")
           }else
           {
              New-AzureCloudServiceDeployment -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -Slot "Staging" -PackageFile $CspkgPath -ServiceConfigurationFile $cscfgUpdatedPath -ServiceName $CloudServiceName 
              Checkpoint-Workflow
              Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
              Select-AzureSubscription -SubscriptionName $AzureConnectionName

              Set-AzureCloudServiceVipSwap -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -ServiceName $CloudServiceName
              Write-Output ("vipswap successful")
           }

       }else
       {
          $Service = Get-AzureService –ServiceName $CloudServiceName
          if ($Service -ne $null)
          {
              New-AzureCloudServiceDeployment -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName -Slot "Production" -PackageFile $CspkgPath -ServiceConfigurationFile $cscfgUpdatedPath -ServiceName $CloudServiceName
              Write-Output ("Creating new deployment successful")
          }else
          {
              throw "$CloudServiceName need to exist in order to deploy."
          }
       }
    }   
}