<#
.SYNOPSIS

Publishes a web application (and any other resources specified in the template file) to a resource group.

.EXAMPLE

.\Publish-AzureResourceGroupPreview.ps1 -ResourceGroupName MyResourceGroup -DeploymentName MyDeploymentName `
	-Location '<Enter Your Location>' -TemplateFile .\TemplatePreview-WebSite.json -ParameterFile .\TemplatePreview-WebSite.dev.json `
    -WebDeployPackage .\WebApplication1\Package\WebApplication1.zip
#>

[CmdletBinding()]
param
(
    [parameter(Mandatory = $true)]
    [String]
    $ResourceGroupName,

    [parameter(Mandatory = $true)]
    [String]
    $Location,

    [Parameter(Mandatory = $true)]
    [String]
    $DeploymentName,

    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]
    $TemplateFile,

    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]
    $ParameterFile,

    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]
    $WebDeployPackage
)

Set-StrictMode -Version 3

# Make sure we're in service management mode
$wasServiceManagementMode = Get-Module -Name Azure -ListAvailable
Switch-AzureMode AzureServiceManagement

# Upload the deploy package to a blob in our storage account, so that
#   the MSDeploy extension can access it.  Create the container if it
#   doesn't already exist.
$containerName = 'msdeploypackages'
$blobName = (Get-Date -Format 'ssmmhhddMMyyyy') + '-' + $ResourceGroupName + '-' + $DeploymentName + '-WebDeployPackage.zip'

# Use the CurrentStorageAccount which is set by Set-AzureSubscription
if (!(Get-AzureStorageContainer $containerName -ErrorAction SilentlyContinue)) 
{
    New-AzureStorageContainer -Name $containerName -Permission Off
}              
Set-AzureStorageBlobContent -Blob $blobName -Container $containerName -File $WebDeployPackage

# Create a SAS token to add to the blob's URI that we give to MSDeploy. This gives it temporary read-only
#   access to the package.
$webDeployPackageUri = New-AzureStorageBlobSASToken -Container $containerName -Blob $blobName -Permission r -FullUri

# Switch to the resource manager APIs
Switch-AzureMode AzureResourceManager

# Read the values from the parameters file and create a hashtable. We do this because we need to modify one 
#   of the parameters, otherwise we could pass the file directly to New-AzureResourceGroupDeployment.
$parameters = New-Object -TypeName hashtable
$jsonContent = Get-Content $ParameterFile -Raw | ConvertFrom-Json
$jsonContent.parameterValues | Get-Member -Type NoteProperty | ForEach-Object {
    $parameters.Add($_.Name, $jsonContent.parameterValues.($_.Name))
}

# Set the msdeployPackageUri parameter to the URL of the package. This is referenced in the template file.
$parameters.msdeployPackageUri = $WebDeployPackageUri

# Create a new resource group (if it doesn't already exist) using our template file and template parameters
New-AzureResourceGroup -Name $ResourceGroupName -DeploymentName $DeploymentName `
    -Location $Location -TemplateFile $TemplateFile -TemplateParameterObject $parameters -Force

# Switch back to original mode before exiting
if ($wasServiceManagementMode) { 
    Switch-AzureMode AzureServiceManagement
}
