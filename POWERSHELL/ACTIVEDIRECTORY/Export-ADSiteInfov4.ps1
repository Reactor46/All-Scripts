<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory sites
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site information

.EXAMPLE 
	.\Export-ADSiteInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 3.0 - Initial Release
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

Try
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "PowerShell Group Policy module could not be loaded. $($_.Exception.Message)";
	exit
}
#EndRegion

#Region Global Variables
$ADRootDSE = Get-ADRootDSE
$forestName = (Get-ADForest).Name
$domDNS = (Get-ADDomain).DNSRoot
$NameSpace = "root\CIMV2"
$RootRNC = ($ADRootDSE).rootDomainNamingContext
$RootNC = ($ADRootDSE).ConfigurationNamingContext
$RootSC = ($ADRootDSE).defaultNamingContext
$reportsFolder = 'C:\LazyWinAdmin\Reports'
[PSObject[]]$ADSitesObject = @()
#EndRegion

#Region Functions

Function fnGet-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
} #End function fnGet-LongDate

Function fnGet-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function fnGet-ReportDate

Function fnGet-GPSiteLink {

Param 
(
    [Parameter(Position=0,ValueFromPipeline=$True)]
    [string]$SiteName="Default-First-Site-Name",
    [Parameter(Position=1)]
    [string]$Domain="USON",
    [Parameter(Position=2)]
    [string]$Forest="USON.LOCAL"
)

	Begin
	{
	    Write-Verbose "Starting Function"
	    #define the permission constants hash table
	    $GPPerms=@{
	        "permGPOApply"                 = 65536;
	        "permGPORead"                  = 65792;
	        "permGPOEdit"                  = 65793;
	        "permGPOEditSecurityAndDelete" = 65794;
	        "permGPOCustom"                = 65795;
	        "permWMIFilterEdit"            = 131072;
	        "permWMIFilterFullControl"     = 131073;
	        "permWMIFilterCustom"          = 131074;
	        "permSOMLink"                  = 1835008;
	        "permSOMLogging"               = 1573120;
	        "permSOMPlanning"              = 1573376;
	        "permSOMGPOCreate"             = 1049600;
	        "permSOMWMICreate"             = 1049344;
	        "permSOMWMIFullControl"        = 1049345;
	        "permStarterGPORead"           = 197888;
	        "permStarterGPOEdit"           = 197889;
	        "permStarterGPOFullControl"    = 197890;
	        "permStarterGPOCustom"         = 197891;
	        }
	    
	    #define the GPMC COM Objects
	    $gpm = New-Object -ComObject "GPMGMT.GPM"
	    $gpmConstants = $gpm.GetConstants()
	    $gpmDomain = $gpm.GetDomain($domain,"",$gpmConstants.UseAnyDC)
	} #Begin
	Process 
	{
	 	ForEach ($item in $siteName) 
		{
	   		#connect to site container
	   		$SiteContainer = $gpm.GetSitesContainer($forest,$domain,$null,$gpmConstants.UseAnyDC)
	   		Write-Verbose "Connected to site container on $($SiteContainer.domainController)"
	   		#get sites
	   		Write-Verbose "Getting $item"
	   		$site = $SiteContainer.GetSite($item)
	   		Write-Verbose ("Found {0} sites" -f ($sites | measure-object).count )
	   		if ($site) 
			{
	       		Write-Verbose "Getting site GPO links"
	       		$links = $Site.GetGPOLinks()
	       		if ($links) 
				{
	          		#add the GPO name
	          		Write-Verbose ("Found {0} GPO links" -f ($links | measure-object).count)
	          		$links | Select @{Name = "Name";Expression={($gpmDomain.GetGPO($_.GPOID)).DisplayName}},
	          		@{Name = "Description";Expression = {($gpmDomain.GetGPO($_.GPOID)).Description}},GPOID,Enabled,Enforced,GPODomain,SOMLinkOrder,@{Name = "SOM";Expression = {$_.SOM.Path}}
	        	} #if $links
	   		} #if $site
	 	} #foreach site  
	   
	} #process
	End 
	{
		Write-Verbose "Finished"
	} #end
} #End function fnGet-GPSiteLink

#EndRegion




#Region Script
#Begin Script

#Region SiteConfig
#Begin collecting AD Site Configuration info.
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites | Sort-Object -Property Name

ForEach ($Site in $Sites)
{
	$SiteLocation = [String]($Site).Location
	$SiteName = [String]$Site.Name
	$SCSubnets = [String]($Site.Subnets -join ("`n"))
	$SiteLinks = [String]($Site.SiteLinks -join ("`n"))
	$AdjacentSites = [String]($Site.AdjacentSites -join ("`n"))
	$SiteDomains = [String]($Site.Domains -join ("`n"))
	$SiteServers = [String]($Site.Servers -join ("`n"))
	$BridgeHeads = [String]($Site.BridgeHeadServers -join ("`n"))
	
	$adSite += Get-ADObject -LDAPFilter '(objectClass=site)' -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions -ErrorAction SilentlyContinue | Where { $_.Name -match $SiteName }
    $gpoCount = ($adSite).gpLink.count
	
	If ( ( $adSite ).gpLink -eq $null )
	{
		$gpoNames = "None."
	}
    Else
    {
        $gpoNames = @()
        $siteGPOS = @()
		ForEach ( $siteDomain in ($site).Domains )
		{
			$siteGPOS += fnGet-GPSiteLink -SiteName $SiteName -Domain $siteDomain -Forest $forestName
		}
		
        ForEach ($siteGPO in $siteGPOS)
        {
            $id = ($siteGPO).GPOID
            $gpoDom = ($siteGPO).GPODomain
            $gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction SilentlyContinue
            #$gpoGUID = $gpoInfo.Id.ToString()
            $gpoName = $gpoInfo.DisplayName.ToString()
			
			$gpoNames += $gpoName
			
			$id = $gpoDom = $gpoInfo = $gpoGUID = $gpoName = $null
        }
    }
		
		
	$ADSitesObject += New-Object -TypeName PSCustomObject -Property ([Ordered] @{
		"Site Name" = $SiteName
		"Site Location" = $SiteLocation
		"Site Links" = $SiteLinks
		"Adjacent Sites" = $AdjacentSites
		"Subnets in Site" = $SCSubnets
		"Domains in Site" = $SiteDomains
		"Servers in Site" = $SiteServers
		"Bridgehead Servers" = $BridgeHeads
		"GPOs linked to Site" = $gpoNames -join ("`n")
		"Notes" = $null
	})
	
$adSite = $gpoNames = $SiteName = $SiteLocation = $SiteLinks = $AdjacentSites = $SCSubnets = $SiteDomains = $SiteServers = $BridgeHeads = $null
}

#EndRegion

#Save output
If ( ( Test-Path -Path $reportsFolder -PathType Container ) -eq $false )
{
	New-Item -Path $reportsFolder -ItemType Directory
	Write-Host "Folder: $reportsFolder not present, creating new folder..." -BackgroundColor White -ForegroundColor Red
}

$wsName = "AD Site Configuration"
$OutputFileName = "Active_Directory_Site_Info_as_of_$(fnGet-ReportDate).xlsx"
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

$Excel = $ADSitesObject | Sort-Object -Property "Site Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
#$Excel.Workbook.Worksheets[$wsName].Cells["A2:Z1048576"].Style.WrapText = $true
$Sheet = $Excel.Workbook.Worksheets["AD Site Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "Active Directory Site Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid

#EndRegion