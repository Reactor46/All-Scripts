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

function AddColumn($fieldType, $fieldLabel, $fieldRequired)
{
    # Add column
    $SPFieldType = [Microsoft.SharePoint.SPFieldType]::$fieldType 
    $exportList.Fields.Add($fieldLabel,$SPFieldType,$fieldRequired)
}

function CreateSharePointlist
{
    # Get SPWeb
    try
    {
        $currentWeb = (New-Object Microsoft.SharePoint.SPSite($siteUrl)).OpenWeb()
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
    
    # Try to get list
    $global:exportList = $currentWeb.Lists[$listName]
 
    # Create list if necessary
    if($global:exportList -ne $null)
    {$global:exportList.Delete()}

    # Add list
    $listId = $currentWeb.Lists.Add($listName,$listDescription,$listTemplate)
  
    # Get list
    $global:exportList = $currentWeb.Lists[$listName]
    
    # Add columns
    $dump = AddColumn -fieldType "Boolean" -fieldLabel "Deployed" -fieldRequired $false
    $dump = AddColumn -fieldType "Boolean" -fieldLabel "ContainsCasPolicy" -fieldRequired $false
    $dump = AddColumn -fieldType "Boolean" -fieldLabel "ContainsGlobalAssembly" -fieldRequired $false
    $dump = AddColumn -fieldType "Boolean" -fieldLabel "ContainsWebApplicationResource" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "DeployedServers" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "DeployedWebApplications" -fieldRequired $false
    $dump = AddColumn -fieldType "Text" -fieldLabel "DeploymentState" -fieldRequired $false
    $dump = AddColumn -fieldType "Note" -fieldLabel "LastOperationDetails" -fieldRequired $false
    $dump = AddColumn -fieldType "Text" -fieldLabel "Status" -fieldRequired $false
    
    # Update list
    $global:exportList.Update()

    # Get default view
    $view = $global:exportList.Views[0]
    
    # Add / delete columns
    try{$view.ViewFields.delete("Attachments")} catch{}
    $view.ViewFields.add("Deployed")
	$view.ViewFields.add("ContainsCasPolicy")
	$view.ViewFields.add("ContainsGlobalAssembly")
	$view.ViewFields.add("ContainsWebApplicationResource")
	$view.ViewFields.add("DeployedServers")
	$view.ViewFields.add("DeployedWebApplications")
    $view.ViewFields.add("DeploymentState")
    $view.ViewFields.add("LastOperationDetails")
    $view.ViewFields.add("Status")
    
    # Update default view
	$view.Update()
    
    write-host "The list '$listName' has been created." -foreground "green"
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

# Get farm
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
$farm=[Microsoft.SharePoint.Administration.SPFarm]::Local

# Local variables
$solutionsList = $null
$solutionsList = @()

# Create SP list if necessary
if ($siteUrl -ne "")
{
    if ($listName -eq "") {$listName = "Farm Solutions Export"}
    CreateSharePointlist
}

# Build structure
$itemStructure = New-Object psobject 
$itemStructure | Add-Member -MemberType NoteProperty -Name "DisplayName" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "Deployed" -value ""
$itemStructure | Add-Member -MemberType NoteProperty -Name "ContainsCasPolicy" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "ContainsGlobalAssembly" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "ContainsWebApplicationResource" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "DeployedServers" -value ""
$itemStructure | Add-Member -MemberType NoteProperty -Name "DeployedWebApplications" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "DeploymentState" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "LastOperationDetails" -value "" 
$itemStructure | Add-Member -MemberType NoteProperty -Name "Status" -value ""

# Get solutions
foreach ($solution in $farm.Solutions)
{
    if ($exportToSharePoint)
    {
        # Create SPItem
        $newItem = $global:exportList.Items.Add()
    }

    $solutionInfos = $itemStructure | Select-Object *; 
    $solutionInfos.DisplayName = $solution.DisplayName;
    if ($exportToSharePoint)
    {
        $titleField = $global:exportList.Fields | where {$_.internalname -eq "LinkTitle"}
        $newItem[$titleField] = $solution.DisplayName
    }
    
    $solutionInfos.Deployed = $solution.Deployed;
    if ($exportToSharePoint) {$newItem["Deployed"] = $solution.Deployed}
    
    $solutionInfos.ContainsCasPolicy = $solution.ContainsCasPolicy;
    if ($exportToSharePoint) {$newItem["ContainsCasPolicy"] = $solution.ContainsCasPolicy}
    
    $solutionInfos.ContainsGlobalAssembly = $solution.ContainsGlobalAssembly;
    if ($exportToSharePoint) {$newItem["ContainsGlobalAssembly"] = $solution.ContainsGlobalAssembly}
    
    $solutionInfos.ContainsWebApplicationResource = $solution.ContainsWebApplicationResource;
    if ($exportToSharePoint) {$newItem["ContainsWebApplicationResource"] = $solution.ContainsWebApplicationResource}
    
    $serverIndex = 0
    foreach ($server in $solution.DeployedServers)
    {
       $solutionInfos.DeployedServers += $solution.DeployedServers[$serverIndex].Name +"`n";
       $serverIndex++;
    }
    if ($exportToSharePoint) {$newItem["DeployedServers"] = $solutionInfos.DeployedServers}
    
    $webApplicationIndex = 0
    if ($solution.DeployedWebApplications -ne $null)
    {
        foreach ($webApplication in $solution.DeployedWebApplications)
        {
           $solutionInfos.DeployedWebApplications += $solution.DeployedWebApplications[$webApplicationIndex].Name +"`n";
           $webApplicationIndex++;
        }
    }
    if ($exportToSharePoint) {$newItem["DeployedWebApplications"] =$solutionInfos.DeployedWebApplications}
    
    $solutionInfos.DeploymentState = $solution.DeploymentState;
    if ($exportToSharePoint) {$newItem["DeploymentState"] = $solution.DeploymentState}
    
    $solutionInfos.LastOperationDetails = $solution.LastOperationDetails;
    if ($exportToSharePoint) {$newItem["LastOperationDetails"] = $solution.LastOperationDetails}
    
    $solutionInfos.Status = $solution.Status;
    if ($exportToSharePoint) {$newItem["Status"] = $solution.Status}

    $solutionsList += $solutionInfos;
    
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
if (!$fileName) {$fileName = "SharePoint - Farm Solutions Export"}
if (!$exportPath) {$exportPath = "."}
if (!$delimiter) {$delimiter = ","}

# Export
$solutionsList | Where-Object {$_} | Export-Csv -Delimiter "$delimiter" -Path "$exportPath\$fileName.csv" -notype;

write-host "The data have been exported to the csv file $exportPath\$fileName.csv." -foreground "green"

