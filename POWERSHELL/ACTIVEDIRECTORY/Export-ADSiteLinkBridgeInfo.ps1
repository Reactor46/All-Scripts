<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Link Bridge Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory site link bridges
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site link bridge information

.EXAMPLE 
	.\Export-ADSiteLinkBridgeInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 2.0
# 
###########################################################################


#Region Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
Try
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";
	exit
}
#EndRegion

#Region Global Variables
$ADRootDSE = Get-ADRootDSE
[PSObject[]]$ADSiteLinkBridgeObject = @()
$reportsFolder = 'C:\LazyWinAdmin\Reports'
#EndRegion

#Region Functions

Function fnGet-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
} #End function fnGet-LongDate

Function fnGet-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function fnGet-ReportDate

#EndRegion

#Region Script
#Begin Script

#Region SiteLinkBridgeConfig
#Begin collecting AD Site Link Bridge Configuration info.
$SiteLinkBridges = Get-ADReplicationSiteLinkBridge -Filter * | Sort-Object -Property Name

ForEach ($slb in $SiteLinkBridges)
{
	$slbName = [String]$slb.Name
	$slbDN = [String]$slb.distinguishedName
	$slbLinksIncluded = [String]($slb.SiteLinksIncluded -join "`n")
	
	$ADSiteLinkBridgeObject += New-Object -TypeName PSCustomObject -Property ([Ordered] @{
		"Site Link Bridge Name" = $slbName
		"Site Link Bridge DN" = $slbDN
		"Site Links In Bridge" = $slbLinksIncluded
	})
	$slbName = $slbDN = $slbLinksIncluded = $null
}
#EndRegion

#Save output
If ( ( Test-Path -Path $reportsFolder -PathType Container ) -eq $true )
{
	Write-Host "Folder: $reportsFolder already exists..." -BackgroundColor White -ForegroundColor DarkGreen
}
Else
{
	New-Item -Path $reportsFolder -ItemType Directory
	Write-Host "Folder: $reportsFolder not present, creating new folder..." -BackgroundColor White -ForegroundColor Red
}

$wsName = "AD Site-Link Bridge Config"
$OutputFileName = "\Active_Directory_Site_Link_Bridge_Info_as_of_$(fnGet-ReportDate).xlsx"
$OutputFile = Join-Path -Path $reportsFolder -ChildPath $OutputFileName
    $ExcelParams = @{
	    Path = $OutputFile
	    StartRow = 2
	    StartColumn = 1
	    AutoSize = $true
	    AutoFilter = $true
	    BoldTopRow = $true
	    FreezeTopRow = $true
    }
$Excel = $ADSiteLinkBridgeObject | Sort-Object -Property "Site Link Bridge Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site-Link Bridge Config"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
#$Excel.Workbook.Worksheets[$wsName].Cells["A2:Z1048576"].Style.WrapText = $true
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "Active Directory Site-Link Bridge Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid
#EndRegion