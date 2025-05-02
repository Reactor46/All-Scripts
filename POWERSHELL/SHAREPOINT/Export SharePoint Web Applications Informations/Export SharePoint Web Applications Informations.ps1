Param(
 [Parameter(Mandatory=$False)]
 [string]$fileName,
 [Parameter(Mandatory=$False)]
 [string]$exportPath,
 [Parameter(Mandatory=$False)]
 [string]$delimiter,
 [Parameter(Mandatory=$False)]
 [string]$siteUrl,
 [Parameter(Mandatory=$False)]
 [string]$listName,
 [Parameter(Mandatory=$False)]
 [string]$listDescription
)

Add-PSSnapin Microsoft.SharePoint.Powershell

function AddColumn($fieldType, $fieldLabel, $fieldRequired)
{
    # Add column
    $SPFieldType = [Microsoft.SharePoint.SPFieldType]::$fieldType 
    $exportList.Fields.Add($fieldLabel,$SPFieldType,$fieldRequired)
}

function CreateSharePointlist
{   
try
{
    if (!$listDescription) {$listDescription = "Contains all web applications within the farm"}

    # Get SPWeb
    try
    {
        $currentWeb = Get-SPWeb $siteUrl
        $global:exportToSharePoint = $True
    }
    catch
    {
        Write-Warning "The site '$siteUrl' cannot be found; the data will only be exported to the csv file."
        $global:exportToSharePoint = $False
        return
    }

    # Get list template (Generic template) 
    $listTemplate = [Microsoft.SharePoint.SPListTemplateType]::GenericList

    # Get Lists Collection
    $listsCollection = $currentWeb.Lists  

    # Create list if necessary
    $targetList = $listsCollection.TryGetList($listName)
    if($targetList -ne $null) {$targetList.Delete()}
    $newList = $listsCollection.Add($listName,$listDescription,$listTemplate)

    # Get list
    $global:exportList = $currentWeb.Lists[$listName]
    
    # Add columns
    $dump = AddColumn -fieldType "Note" -fieldLabel "URL" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "AlternateUrls" -fieldRequired $false
    $dump = AddColumn -fieldType "Boolean" -fieldLabel "UseClaimsAuthentication" -fieldRequired $false
    $dump = AddColumn -fieldType "Text" -fieldLabel "ApplicationPool" -fieldRequired $false
    $dump = AddColumn -fieldType "Text" -fieldLabel "ApplicationPoolUserName" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "ContentDatabases" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "Sites" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "Solutions" -fieldRequired $false
    
    # Update list
    $global:exportList.Update()

    # Get default view
    $view = $global:exportList.Views[0]
    
    # Add / delete columns
    try{$view.ViewFields.delete("Attachments")} catch{}
    $view.ViewFields.add("URL")
	$view.ViewFields.add("AlternateUrls")
	$view.ViewFields.add("UseClaimsAuthentication")
	$view.ViewFields.add("ApplicationPool")
	$view.ViewFields.add("ApplicationPoolUserName")
	$view.ViewFields.add("ContentDatabases")
    $view.ViewFields.add("Sites")
    $view.ViewFields.add("Solutions")
    
    # Update default view
	$view.Update()
    
    write-host "The list '$listName' has been created." -foreground "green"
    }
    catch
    { 
        Write-Warning "The list '$listName' could not be created; the data will only be exported to the csv file."
        $global:exportToSharePoint = $False
        return
    }
}

# Global variables
$global:exportList = $null
$global:exportToSharePoint = $False

# Test parameters
if (($exportPath -ne "") -and (!(Test-Path -Path $exportPath)))
{
    Write-Warning "The export path is invalid."
    exit 0
}

if (($listName -ne "") -and ($siteUrl -eq ""))
{Write-Warning "You must provide a site URL associated to the list; the data will only be exported to the csv file."}

# Local variables
$webApplicationsList = $null
$webApplicationsList = @()

# Create SP list if necessary
if ($siteUrl -ne "")
{
    if ($listName -eq "") {$listName = "Farm Web Applications Export"}
    CreateSharePointlist
}

# Build structure
$itemStructure = New-Object psobject 
$itemStructure | Add-Member -MemberType NoteProperty -Name "DisplayName" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "URL" -value ""
$itemStructure | Add-Member -MemberType NoteProperty -Name "AlternateUrls" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "UseClaimsAuthentication" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "ApplicationPool" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "ApplicationPoolUserName" -value ""
$itemStructure | Add-Member -MemberType NoteProperty -Name "ContentDatabases" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "Sites" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "Solutions" -value "" 

# Get solutions
foreach ($webApplication in Get-SPWebApplication)
{
    if ($exportToSharePoint)
    {
        # Create SPItem
        $newItem = $global:exportList.Items.Add()
    }

    $webApplicationInfos = $itemStructure | Select-Object *; 
    $webApplicationInfos.DisplayName = $webApplication.DisplayName;
    if ($exportToSharePoint)
    {
        $titleField = $global:exportList.Fields | where {$_.internalname -eq "LinkTitle"}
        $newItem[$titleField] = $webApplication.DisplayName
    }
    
    $webApplicationInfos.URL = $webApplication.Url;
    if ($exportToSharePoint) {$newItem["URL"] = $webApplication.Url}
    
    foreach ($alternateURL in $webApplication.AlternateUrls)
    {
       $webApplicationInfos.AlternateUrls += $alternateURL.PublicUrl + "|" +$alternateURL.Zone +"`n";
       $alternateURLIndex++;
    }
    if ($exportToSharePoint) {$newItem["AlternateUrls"] = $webApplicationInfos.AlternateUrls}

    $webApplicationInfos.UseClaimsAuthentication = $webApplication.UseClaimsAuthentication;
    if ($exportToSharePoint) {$newItem["UseClaimsAuthentication"] = $webApplication.UseClaimsAuthentication}
    
    $webApplicationInfos.ApplicationPool = $webApplication.ApplicationPool.DisplayName;
    if ($exportToSharePoint) {$newItem["ApplicationPool"] = $webApplication.ApplicationPool.DisplayName}

    $webApplicationInfos.ApplicationPoolUserName = $webApplication.ApplicationPool.Username;
    if ($exportToSharePoint) {$newItem["ApplicationPoolUserName"] = $webApplication.ApplicationPool.Username}

    foreach ($contentDatabase in $webApplication.ContentDatabases)
    {$webApplicationInfos.ContentDatabases += $contentDatabase.Name +"`n";}
    if ($exportToSharePoint) {$newItem["ContentDatabases"] = $webApplicationInfos.ContentDatabases}
    
    foreach ($site in $webApplication.Sites)
    {$webApplicationInfos.Sites += $site.URL +"`n";}
    if ($exportToSharePoint) {$newItem["Sites"] = $webApplicationInfos.Sites}
    
    Get-SPSolution | ForEach-Object {
    if ($_.LastOperationDetails.IndexOf($webApplication.Url) -gt 0) 
    { $webApplicationInfos.Solutions += $_.DisplayName +"`n";}
    if ($_.DeploymentState -eq "GlobalDeployed")
    { $webApplicationInfos.Solutions += $_.DisplayName +" (Globally Deployed)`n";}    
    }

    if ($exportToSharePoint) {$newItem["Solutions"] = $webApplicationInfos.Solutions}

    $webApplicationsList += $webApplicationInfos;
    
    if ($exportToSharePoint)
    {
        # Update SPItem
        $newItem.Update()
    }
}

write-host "The data have been collected." -foreground "green"

if ($exportToSharePoint)
{
    write-host "The data have been copied in the SharePoint list." -foreground "green"
    
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Navigate("$siteUrl/Lists/$listName/AllItems.aspx")
    try{$ie.Visible = $true} catch{}
}
 
# Set export variables
if (!$fileName) {$fileName = "SharePoint - Farm Web Applications Export"}
if (!$exportPath) {$exportPath = "."}
if (!$delimiter) {$delimiter = ","}

# Export
$webApplicationsList | Where-Object {$_} | Export-Csv -Delimiter "$delimiter" -Path "$exportPath\$fileName.csv" -notype;

write-host "The data have been exported to the csv file $exportPath\$fileName.csv." -foreground "green"

