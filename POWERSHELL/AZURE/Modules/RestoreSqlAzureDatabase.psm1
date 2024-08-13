<#
.SYNOPSIS
    Restore a SqlAzure Database from a given storage account in a given subscription.
.DESCRIPTION
    First this workflow requires SqlAuthentication and open connection to the SqlAzure server, so please add the execution host into the server's 
    firewall rule. Otherwise, restore will fail.
    
    Second, please note DBSie and DBEdition only apply when the given database doesnot exist.    
    $DBSize represents DBMaxSize in GB allowed, and if DBSize is null, it takes 1GB as default
    $DBEdition is an Enum type of {None | Business | Web | Premium }, takes Web as default unless specified otherwise. 
    DBSize and DBEdition should follow: Web - 1 or 5GB; Business - 10, 20, 30, 40, 50, 100, or 150 GB
    The workflow restores the SqlAzure Database file to given SqlAzure database in given SqlAzure server. Prints performance data of restore

.EXAMPLE
    import-module .\Restore-SqlAzureDatabase.psm1
    Restore-SqlAzureDatabase -SubscriptionId "11111111-aaaa-bbbb-cccc-222222222222" `
        -SubscriptionName "My Test Subscription" -PfxFilePath "C:\MyAbsolutePath.pfx" -PfxPassword MyPassword -SqlAzureServerName MyServerName `
        -SqlAzureDatabaseName MyDBName -SqlAzureUserName myusername -SqlAzurePassword mypassword -StorageAccountName MyStorage -StroageContainerName MyContainer`
        -StorageBlobName myblob -DBSize 5 -Edition Web
#>

workflow Restore-SqlAzureDatabase
{
    param(
       
                [parameter(Mandatory=$true)]
                [String]$SubscriptionId,

                [parameter(Mandatory=$true)]
                [String]$SubscriptionName,
	
	            [parameter(Mandatory=$true)]
                [String]$PfxFilePath, 

                [parameter(Mandatory=$true)]
                [String]$PfxPassword,    
            
                [Parameter(Mandatory = $true)] 
                [String]$SqlAzureServerName,

                [Parameter(Mandatory = $true)] 
                [String]$SqlAzureDatabaseName,

                [Parameter(Mandatory = $true)] 
                [String]$SqlAzureUserName,

                [Parameter(Mandatory = $true)] 
                [String]$SqlAzurePassword,

                [Parameter(Mandatory = $true)]
                [String]$StorageAccountName,

                [Parameter(Mandatory = $true)]
                [String]$StorageContainerName,

                [Parameter(Mandatory = $true)]
                [String]$StorageBlobName,

                [Int32]$DBSize,

                [String]$DBEdition
    )

    # Check if Windows Azure Powershell is avaiable
    if ((Get-Module -ListAvailable Azure) -eq $null)
    {
        throw "Windows Azure Powershell not found! Please install from http://www.windowsazure.com/en-us/downloads/#cmd-line-tools"
    }

    $DBMaxSize = [Int32]1
    if ($DBSize -ne $null)
    {
       $DBMaxSize = $DBSize
    }

    $Edition = "None"
    if ($DBEdition -ne $null)
    {
        $Edition = $DBEdition
    }

    $Start = [System.DateTime]::Now
    "Starting: " + $Start.ToString("HH:mm:ss.ffffzzz")

    $SecurePwd = ConvertTo-SecureString -String "$PfxPassword" -Force -AsPlainText
    $importedCert = Import-PfxCertificate -FilePath $PfxFilePath  -CertStoreLocation Cert:\CurrentUser\My  -Exportable  -Password $SecurePwd 
    $MyCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList($PfxFilePath, $SecurePwd, "Exportable")
    $SqlCredential = new-object System.Management.Automation.PSCredential($SqlAzureUserName, ($SqlAzurePassword | ConvertTo-SecureString -asPlainText -Force))
    Write-Debug $SqlCredential

    inlinescript
    {
        
        Set-AzureSubscription -SubscriptionName "$using:SubscriptionName" -SubscriptionId $using:SubscriptionId -Certificate $using:MyCert -CurrentStorageAccount $using:StorageAccountName
        Select-Azuresubscription -SubscriptionName "$using:SubscriptionName"     
     
        $StoragePrimaryKey=(Get-AzureStorageKey -StorageAccountName $using:StorageAccountName).Primary
        Write-Output ("Storage Account primary key : {0} " -f $StoragePrimaryKey)

        $Storagectx = New-AzureStorageContext -StorageAccountName $using:StorageAccountName -StorageAccountKey $StoragePrimarykey
        Write-Debug $Storagectx
                
        $SqlDBCtx=New-AzureSqlDatabaseServerContext -ServerName $using:SqlAzureServerName -Credential $using:SqlCredential
        Write-Debug $SqlDBCtx

        if (($StorageCtx -ne $null) -And ($SqlDBCtx -ne $null))
        {
           $using:DBMaxSize
           $using:Edition

           $ImportOP = Start-AzureSqlDatabaseImport -SqlConnectionContext $SqlDBCtx -StorageContext $Storagectx -StorageContainerName $using:StorageContainerName -DatabaseName $using:SqlAzureDatabaseName -blobname $using:StorageBlobName -DatabaseMaxSize $using:DBMaxSize -Edition $using:Edition
           $ImportStatus = Get-AzureSqlDatabaseImportExportStatus -username $using:SqlAzureUserName -password $using:SqlAzurePassword -servername $using:SqlAzureServerName -RequestId $ImportOP.RequestGuid
  
           while ($ImportStatus.Status -ne "Completed" -and $ImportStatus.Status -ne "Failed")
           {
                Write-OutPut ("Restoring Sql Azure Database {0} on server {1} from {2}/{3} of {4}: {5}" -f $using:SqlAzureDatabaseName, $using:SqlAzureServerName, $using:StorageContainerName, $using:StorageBlobName, $using:StorageAccountName, $ImportStatus.Status)
                Start-Sleep 20
                $ImportStatus = Get-AzureSqlDatabaseImportExportStatus -username $using:SqlAzureUserName -password $using:SqlAzurePassword -servername $using:SqlAzureServerName -RequestId $ImportOP.RequestGuid
           }

           if ($ImportStatus.Status -eq "Failed")
           { 
              Write-Output -Message $ImportStatus.ErrorMessage
           }

           Write-OutPut ("Restoring Sql Azure Database {0} on server {1} from {2}/{3} of {4}: {5}" -f $using:SqlAzureDatabaseName, $using:SqlAzureServerName, $using:StorageContainerName, $using:StorageBlobName, $using:StorageAccountName, $ImportStatus.Status)
       }

    }

    $Finish = [System.DateTime]::Now
    $TotalUsed = $Finish.Subtract($Start).TotalSeconds
   
    Write-Output ("Restored Sql Azure Database {0} on server {1} in subscription {2} in {3} seconds." -f $SqlAzureDatabaseName, $SqlAzureServerName, $SubscriptionName, $TotalUsed)
    "Finished " + $Finish.ToString("HH:mm:ss.ffffzzz")
 } 

