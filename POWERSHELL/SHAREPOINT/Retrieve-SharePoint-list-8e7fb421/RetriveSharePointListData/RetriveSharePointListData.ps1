<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages 
#> 

#requires -Version 2

#Import Localized Data
Import-LocalizedData -BindingVariable Messages
#Load Microsoft SharePoint Snapin
if ((Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {Add-PSSnapin Microsoft.SharePoint.PowerShell}

Function New-OSCPSCustomErrorRecord
{
	#This function is used to create a PowerShell ErrorRecord
    [CmdletBinding()]
    Param
    (
       [Parameter(Mandatory=$true,Position=1)][String]$ExceptionString,
       [Parameter(Mandatory=$true,Position=2)][String]$ErrorID,
       [Parameter(Mandatory=$true,Position=3)][System.Management.Automation.ErrorCategory]$ErrorCategory,
       [Parameter(Mandatory=$true,Position=4)][PSObject]$TargetObject
    )
    Process
    {
       $exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
       $customError = New-Object System.Management.Automation.ErrorRecord($exception,$ErrorID,$ErrorCategory,$TargetObject)
       return $customError
    }
}

Function Get-OSCSPListItem
{
	<#
		.SYNOPSIS
		Get-OSCSPListItem is an advanced function which can be used to retrieve the list data in Microsoft SharePoint 2010.
		.PARAMETER SiteURL
		Indicates the URL of a SharePoint site, in the form http://server_Name/sites/sitename.
		.PARAMETER ListName
		Indicates the Name of a SharePoint site list, for example, "Shared Documents", "Calendar", "Tasks", "Site Pages".
		.PARAMETER Property
		Indicates the properties of a SharePoint site list item which will be returned by Get-OSCSPListItem.
		For example, "ID", "Name", "ParentList". Wildcard is accepted if you want get all properties.
		.PARAMETER ListAllItems
		Indicates Get-OSCSPListItem will return all items in a specified list.
		.PARAMETER ItemID
		Indicates Get-OSCSPListItem will return a item with the specified ID.
		.PARAMETER StartID
		Indicates the start ID of an ID range.
		.PARAMETER EndID
		Indicates the end ID of an ID range
		.PARAMETER ItemURL
		Indicates a URL of an Item which will be downloaded.
		.PARAMETER DownloadPath
		Indicates a local path which will be used to save the item.
		.EXAMPLE
		#Get all items from a list
		Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName "Shared Documents" -ListAllItems -Property "ID","Name","ParentList"
		.EXAMPLE
		#Get a single item with all properties from a list
		Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName "Shared Documents" -ItemID 1 -Property *
		.EXAMPLE
		#Get multiple items from a list
		Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName "Shared Documents" -StartID 4 -EndID 5 -Property "ID","Name"
		.EXAMPLE
		#Download a file from a document library list
		Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName "Shared Documents" -ItemURL "Shared Documents/NewWordDoc.docx" -DownloadPath "C:\Data" -Verbose
		.EXAMPLE
		#Download multiple files from a document library list
		Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName "Shared Documents" -ListAllItems -Property "Name","URL","ParentList" | ForEach-Object {
			Get-OSCSPListItem -SiteURL "http://server_name/sites/sitename" -ListName $_.ParentList -ItemURL $_.Url -DownloadPath "C:\Data" -Verbose
		}
		.LINK
		Windows PowerShell Advanced Function
		http://technet.microsoft.com/en-us/library/dd315326.aspx
		.LINK
		Microsoft.SharePoint.SPWeb class
		http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.spweb.aspx
		.LINK
		Microsoft.SharePoint.SPList class
		http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splist.aspx
	#>
	
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="SingleID")]
    Param
    (
		#Define parameters and parameter sets
        [Parameter(Mandatory=$true,Position=1)]
        [string]$SiteURL,
		[Parameter(Mandatory=$true,Position=2)]
        [string]$ListName,
		[Parameter(Mandatory=$false,Position=3)]
        [string[]]$Property="*",
		[Parameter(Mandatory=$true,Position=4,ParameterSetName="ListAllItems")]
        [switch]$ListAllItems=$true,
        [Parameter(Mandatory=$true,Position=4,ParameterSetName="SingleID")]
        [int]$ItemID,
        [Parameter(Mandatory=$true,Position=4,ParameterSetName="MultipleID")]
        [int]$StartID,       
        [Parameter(Mandatory=$true,Position=5,ParameterSetName="MultipleID")]
        [int]$EndID,       
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=4,ParameterSetName="SaveItem")]
        [string]$ItemURL,
        [Parameter(Mandatory=$true,Position=5,ParameterSetName="SaveItem")]
        [string]$DownloadPath
    )
	Process
	{
		Try
		{
			#Use Get-SPSite to get a Microsoft SharePoint site object
			$spSite = Get-SPSite -Identity $SiteURL -ErrorAction Stop -Verbose:$false
		}
		Catch 
		{
			#If Get-SPSite failed for any reason, this function will be terminated.
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString $Error[0] `
			-ErrorCategory ResourceUnavailable -ErrorID 1 -TargetObject $pscmdlet
			$pscmdlet.WriteError($customError)
			return $null
		}
		#Get the specified web site,$spWeb is a instance of Microsoft.SharePoint.SPWeb class
		$spWeb = $spSite.OpenWeb()
		#Get the specified list, $spList is a instance of Microsoft.SharePoint.SPList class
		$spList = $spWeb.Lists[$ListName]
		#If the specified list exists, keep on processing, otherwise display an error message
		if ($spList -ne $null) {
			#Check the parameter set name
			Switch ($pscmdlet.ParameterSetName) {
				"ListAllItems" {
					#Get all items from a list by default
					$spList.Items | Select-Object -Property $Property
				}
				"SingleID" {
					#Get a single item from a list
					if (-not (($ItemID -lt 0) -or ($ItemID -gt $spList.Items.Count))) {
						$spList.GetItemByID($ItemID) | Select-Object -Property $Property
					} else {
						$customError = New-OSCPSCustomErrorRecord `
						-ExceptionString $Messages.CannotFindItem `
						-ErrorCategory ObjectNotFound -ErrorID 1 -TargetObject $pscmdlet
						$pscmdlet.WriteError($customError)								
					}
				}
				"MultipleID" {
					#Get multiple items from a list
					if (-not (($StartID -lt 0) -or ($EndID -gt $spList.Items.Count))) {
						for ($id = $StartID;$id -le $EndID;$id++) {
							$spList.GetItemByID($id) | Select-Object -Property $Property
						}
					} else {
						$customError = New-OSCPSCustomErrorRecord `
						-ExceptionString $Messages.CannotFindItemRange `
						-ErrorCategory ObjectNotFound -ErrorID 1 -TargetObject $pscmdlet
						$pscmdlet.WriteError($customError)					
					}
				}
				"SaveItem" {
					#Download a file from a document library list
					#User cannot download file from a non-dcoument library list
					if ($spList.BaseType -ne "DocumentLibrary") {
						$customError = New-OSCPSCustomErrorRecord `
						-ExceptionString $Messages.CannotDLFromNonDocList `
						-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
						$pscmdlet.WriteError($customError)						
					} else {
						#spFile is a instance of Microsoft.SharePoint.SPFile class
						$spFile = $spWeb.GetFile($ItemURL)
						if (-not $spFile.Exists) {
							#If file not exists, an error message will be displayed.
							$errMsg = $Messages.CannotFindSpeciedItem -replace "Placeholder01",$ItemURL
							$customError = New-OSCPSCustomErrorRecord `
							-ExceptionString $errMsg `
							-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
							$pscmdlet.WriteError($customError)
						} else {
							#If file exists, use static method WriteAllBytes of System.IO.Path class 
							#for saving the file to a local path.
							$spContent = $spFile.OpenBinary()
							if (($DownloadPath.SubString($DownloadPath.length - 1,1)) -ne "\") {
								$spTempFile = $DownloadPath + "\" + $spFile.Name
							} else {
								$spTempFile = $DownloadPath + $spFile.Name
							}
							$verboseMessage = $Messages.SaveFilePrompt
							$verboseMessage = $verboseMessage -replace "Placeholder01",$($spFile.Name)
							$verboseMessage = $verboseMessage -replace "Placeholder02",$spTempFile
							$pscmdlet.WriteVerbose($verboseMessage)
							[System.IO.File]::WriteAllBytes($spTempFile,$spContent)
						}
					}
				}
			}
		} else {
			#If list not exists, an error message will be displayed.
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString "Cannot find the specified list: `"$ListName`"." `
			-ErrorCategory ResourceUnavailable -ErrorID 1 -TargetObject $pscmdlet
			$pscmdlet.WriteError($customError)
		}
		#Dispose SPSite and SPWeb object according to the article:
		#Microsoft Press: Using Windows PowerShell to Perform and Automate Farm Administrative Tasks
		#Memory Considerations When Using Windows PowerShell 
		$spWeb.Dispose()
		$spSite.Dispose()
	}
}

