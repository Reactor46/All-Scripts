#--------------------------------------------------------------------------------- 
#The sample scripts are not supported under any Microsoft standard support 
#program or service. The sample scripts are provided AS IS without warranty  
#of any kind. Microsoft further disclaims all implied warranties including,  
#without limitation, any implied warranties of merchantability or of fitness for 
#a particular purpose. The entire risk arising out of the use or performance of  
#the sample scripts and documentation remains with you. In no event shall 
#Microsoft, its authors, or anyone else involved in the creation, production, or 
#delivery of the scripts be liable for any damages whatsoever (including, 
#without limitation, damages for loss of business profits, business interruption, 
#loss of business information, or other pecuniary loss) arising out of the use 
#of or inability to use the sample scripts or documentation, even if Microsoft 
#has been advised of the possibility of such damages 
#--------------------------------------------------------------------------------- 

#requires -Version 3.0

<#
 	.SYNOPSIS
       This script can be batch change the content type of blob.                                         .
    .DESCRIPTION
       This script can be batch change the content type of blob.        
    .PARAMETER  ContainerName
		Specifies the name of container.
	.PARAMETER	BlobName
		Specifies the name of blob. It is not required, if the name of blob is not specified, this will changes the content 
        type of all blob items in the sepcified container and in all child container.
	.PARAMETER  ContentType
		Specifies a content type, it can be any type. 
	.PARAMETER  StorageAccountName
		Specifies the name of the storage account to be connected.
	.EXAMPLE
        C:\PS> C:\Script\ChangeAzureBlobContentType.ps1 -StorageAccountName "AzureStorage01" -ContainerName "pics" -BlobName "doc.txt" -ContentType "image/jpg"

        Successfully changed the content type of 'doc.txt' to 'image/jpg'.
		
		This example shows how to change the content type of 'doc.txt' to 'image/jpg'.
	.EXAMPLE
        C:\PS> C:\Script\ChangeAzureBlobContentType.ps1 -StorageAccountName "AzureStorage01" -ContainerName "pics" -ContentType "image/jpg"

        Successfully changed content type of '1365687925207897_840_560.jpg' to 'image/jpg'.
        Successfully changed content type of 'Shanghai/1365687925207897_840_560.jpg' to 'image/jpg'.
        Successfully changed content type of 'doc.txt' to 'image/jpg'.
        Successfully changed content type of 'doc1.txt' to 'image/jpg'.
		
		If you don't specify a BlobName parameter, the script batch changes the content type of all blob items in the specified container and in all child container.
    .EXAMPLE
        C:\PS> C:\Script\ChangeAzureBlobContentType.ps1 -StorageAccountName "AzureStorage01" -ContainerName "pics" -ContentType "image/jpg" -WhatIf
		
        What if: Performing the operation "Change the content type" on target "1365687925207897_840_560.jpg".
        What if: Performing the operation "Change the content type" on target "Shanghai/1365687925207897_840_560.jpg".
        What if: Performing the operation "Change the content type" on target "doc.txt".
        What if: Performing the operation "Change the content type" on target "doc1.txt".

		It displays a message that describes the effect of the command, instead of executing the command.
#>
[CmdletBinding(SupportsShouldProcess = $true)]

Param
(
    [Parameter(Mandatory = $true)]
    [Alias('CN')]
    [String]$ContainerName,
    [Parameter(Mandatory = $false)]
    [Alias('BN')]
    [String]$BlobName,
    [Parameter(Mandatory = $true)]
    [Alias('CT')]
    [String]$ContentType,
    [Parameter(Mandatory = $true)]
    [Alias('SN')]
    [String]$StorageAccountName
)

#Check if Windows Azure PowerShell Module is avaliable
If((Get-Module -ListAvailable Azure) -eq $null)
{
    Write-Warning "Windows Azure PowerShell module not found! Please install from http://www.windowsazure.com/en-us/downloads/#cmd-line-tools"
}
Else
{

    Function ChangeBlobContentType
    {
        Param
        (
            [String]$ContainerName,
            [String]$BlobName,
            [String]$ContentType
        )

        Write-Verbose "Getting the container object named $ContainerName."
        $BlobContainer = $CloudBlobClient.GetContainerReference($ContainerName)

        Write-Verbose "Getting the blob object named $BlobName."
        $Blob = $BlobContainer.GetBlockBlobReference($BlobName)
                    
        Write-Verbose "Changing content type of '$BlobName' to '$ContentType'."
        Try
        {
            $Blob.Properties.ContentType = $ContentType
            $Blob.SetProperties()
            Write-Host "Successfully changed content type of '$BlobName' to '$ContentType'."
        }
        Catch
        {
            Write-Host "Failed to change content type of '$BlobName'." -ForegroundColor Red
        }
    }

    If($StorageAccountName)
    {
        Get-AzureStorageAccount -StorageAccountName $StorageAccountName -ErrorAction SilentlyContinue `
        -ErrorVariable IsExistStorageError | Out-Null
        #Check if storage account is exist
        If($IsExistStorageError.Exception -eq $null)
        {
            #Specify a Windows Azure Storage Library path
            $StorageLibraryPath = "$env:SystemDrive\Program Files\Microsoft SDKs\Windows Azure\.NET SDK\v2.2\ref\Microsoft.WindowsAzure.Storage.dll"

            #Getting Azure storage account key
            $Keys = Get-AzureStorageKey -StorageAccountName $StorageAccountName
            $StorageAccountKey = $Keys[0].Primary

            #Loading Windows Azure Storage Library for .NET.
            Write-Verbose -Message "Loading Windows Azure Storage Library from $StorageLibraryPath"
            [Reflection.Assembly]::LoadFile("$StorageLibraryPath") | Out-Null

            $Creds = New-Object Microsoft.WindowsAzure.Storage.Auth.StorageCredentials("$StorageAccountName","$StorageAccountKey")
            $CloudStorageAccount = New-Object Microsoft.WindowsAzure.Storage.CloudStorageAccount($creds, $true)
            $CloudBlobClient = $CloudStorageAccount.CreateCloudBlobClient()
        }
        Else
        {
            Write-Warning "Cannot find storage account '$StorageAccountName' because it does not exist. Please make sure thar the name of storage is correct."
        }
    }

    If($ContainerName)
    {
        Get-AzureStorageContainer -Name $ContainerName -ErrorAction SilentlyContinue `
        -ErrorVariable IsExistContainerError | Out-Null
        #Check if container is exist
        If($IsExistContainerError.Exception -eq $null)
        {
            If($BlobName)
            {
                Get-AzureStorageBlob -Container $ContainerName -Blob $BlobName -ErrorAction SilentlyContinue `
                -ErrorVariable IsExistBlobError | Out-Null
                 #Check if blob is exist
                If($IsExistBlobError.Exception -eq $null)
                {
                    #user specifiy a name of blob, the script will change the content type of specified blob only.
                    If($PSCmdlet.ShouldProcess("$BlobName","Change the content type"))
                    {
                        ChangeBlobContentType -ContainerName $ContainerName -BlobName $BlobName -ContentType $ContentType
                    }
                }
                Else
                {
                    Write-Warning "Cannot find blob '$BlobName' because it does not exist. Please make sure thar the name of blob is correct."
                }
            }
            Else
            {
                #If user does not specifiy a blob, the script will change the content type of
                #all blob items in the sepcified container and in all child container.
                $BlobItems = Get-AzureStorageBlob -Container $ContainerName | Where{$_.SnapshotTime -eq $null}
                Foreach($BlobItem in $BlobItems)
                {
                    $BlobN = $BlobItem.Name
                    If($PSCmdlet.ShouldProcess("$BlobN","Change the content type"))
                    {
                        ChangeBlobContentType -ContainerName $ContainerName -BlobName $BlobN -ContentType $ContentType
                    }
                }  
            }
        }
        Else
        {
            Write-Warning "Cannot find container '$ContainerName' because it does not exist. Please make sure thar the name of container is correct."
        }
    }
}