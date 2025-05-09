<#

.SYNOPSIS 
Get and Send Firmware Packages to UCS

.DESCRIPTION
Downloads firmware packages from support.cisco.com and uploads images to UCSM

.PARAMETER UCSMIP
IP Address or comma separated set if IP Addresses of the UCSM instance/s.  If an IP is entered here the UCSM .csv list will be ignored.

.PARAMETER GetLatest
Get the latest version of UCSM Firmware

.PARAMETER UCSMCSV
Full path to a .csv formatted list of UCSM IP's to interact with.  Must be 2 rows with a "UCSM" header in row 1, and a "IP" header in row 2.

.PARAMETER VERSION
Target version you want to deploy.  Valid formats are x.x.xx or x.x(xx).  Example 2.0.2q | 2.0(2q).

.PARAMETER OUTPUTDIR
Directory you want the images downloaded too.  Defaults to "C:\UCSMImages"
If the directory does not exist, it will be created.

.PARAMETER Drivers
Include Drivers .iso's (Always Latest)

.PARAMETER IncludeCSeries
Include C-Series server package.  (By default C-Series packages will not be downloaded)

.EXAMPLE
Update-UcsFirmwarePackages -UCSIP "10.10.10.1,10.10.10.2,10.10.10.3" -GetLatest -Drivers

.EXAMPLE
Update-UcsFirmwarePackages -UCSIP "10.10.10.1,10.10.10.2,10.10.10.3" -version "2.0(2Q)"

.EXAMPLE
Update-UcsFirmwarePackages -UCSCSV "C:\zInput\ucsmList.csv" -version "2.0(2q)"

.Example
Update-UcsFirmwarePackages -UCSCSV "C:\zInput\ucsmList.csv" -version "2.0(2q)" -OUTPUTDIR "\\MyComputer\MyUcsImages"

.Example
Update-UcsFirmwarePackages -UCSCSV "C:\zInput\ucsmList.csv" -version "2.0(2q)" -IncludeCSeries

.NOTES
Author: Chris Shockey, Architect, Cisco Advanced Services
Email: chris.shockey@cisco.com
Version: 1_0

.LINK
http://developer.cisco.com
#>

param(
      [string]$UCSIP,
      [string]$UCSMCSV,
	  [string]$VERSION,
	  [switch]$IncludeCSeries,
	  [switch]$GetLatest,
	  [switch]$Drivers,
	  [string]$OUTPUTDIR
	  
)

#_______________________________________________________________________________
#__________________ GLOBALS_____________________________________________________
#_______________________________________________________________________________

$ReportErrorShowExceptionClass = $true
$Global:cred = $null

#_______________________________________________________________________________
#__________________ FUNCTIONS __________________________________________________
#_______________________________________________________________________________
function Get-UCSMList 
{
	if (!($UCSIP) -and !($UCSMCSV))
	{
		Write-Host "Error: No UCSM instance specified or CSV of UCSM instances specified."
		Write-Host 'Use "-ucsmip YourIP,YourIP,YourIP", or "-ucsmcsv {csvPath}" to specify a UCSM instance or instances'
		Write-Host "Exiting..."
		exit 1
	}
	if ($UCSIP)
	{
		foreach ($IP in $UCSIP)
		{
	    	[array]$mgrList += $IP
		}
	}
    if ($UCSMCSV)
    {
        if (Get-Item $UCSMCSV -ErrorAction SilentlyContinue)
        {
            $global:csvImport = Import-Csv $UCSMCSV
			$global:csvin = $true
			foreach ($ucsm in $csvImport)
			{
		    	[array]$mgrList += $ucsm.ip
			}
        }
        else
        {
              Write-Host "Error: CSV File not found at $UCSMCSV. Please check the path"
              Write-Host "Exiting...."
              exit 1
        }
    }
    return [array]$mgrList
}
function Validate-UcsPSLibrary
{
	if (!(Get-Module CiscoUCSPS -ErrorAction SilentlyContinue))
	{	
		Write-Host "UCS Powershell Library not Loaded, attempting to load..."
		if (Get-Item "c:\Program Files (x86)\Cisco\Cisco UCS PowerTool\" -ErrorAction SilentlyContinue)
		{
			Import-Module "c:\Program Files (x86)\Cisco\Cisco UCS PowerTool\CiscoUCSPs.psd1"
			Import-Module "c:\Program Files (x86)\Cisco\Cisco UCS PowerTool\UcsAlias.psm1"
		}
		else
		{
			Write-Host "UCS Powershell Library not found at c:\Program Files (x86)\Cisco\Cisco UCS PowerTool\"
			Write-Host "Please verify library is installed.  You can download the UCS Powershell Library at: http://developer.cisco.com/web/unifiedcomputing/microsoft"
			Write-Host "Exiting...."
			exit 1
		}
		if (!(Get-Module CiscoUCSPS -ErrorAction SilentlyContinue))
		{
			Write-Host "Failed to load the UCS Powershell Library, please download and install from http://developer.cisco.com/web/unifiedcomputing/microsoft"
			Write-Host "Exiting..."
			exit 1
		}
		else
		{
			Write-Host "Successfully imported the UCS Powershell Library.  Continuing..."
		}
	}
}
function Validate-UcsFIConnect ([string]$ucsm)
{
	$ucsSession = $null
	if (!($cred))
	{
		Write-Host "Enter your UCS Credentials"
		$global:cred = Get-Credential -ErrorAction SilentlyContinue
	}
	$ucsSession = Connect-Ucs -Name $ucsm -Credential $cred -NoSsl -NotDefault -ErrorAction SilentlyContinue
	if (!($ucsSession))
	{
		$ucsSession = Connect-Ucs -Name $ucsm -Credential $cred -NotDefault -ErrorAction SilentlyContinue
	}
	if (!($ucsSession))
	{
		Write-Host "$ucsm: Error Connecting to $ucsm, Most likely causes:"
#		write-host "	- Bad Password"
#		write-host "	- Invalid VIP to the target Fabric Interconnects"
#		write-host "	- Bad Proxy, check your browser proxy settings"
#	    Write-Host "Exiting..."
#		Exit 1
		continue
	}		
	return $ucsSession
}
Function Get-UcsFirmwarePackages ($rootDir)
{
	Write-Host "Enter your CCO (support.cisco.com) Credentials:"
	$ccocred = Get-Credential
	if (($GetLatest) -or ($VERSION -eq $null))
	{
		Try
		{
			$images = Get-UcsCcoImageList -Credential $ccocred | where {$_.ImageName -imatch "ucs-k9-bundle"} | sort version -Descending
		}
		catch
		{
			Write-Host -ForegroundColor Red	"Unable to connect to cisco.com for images.  Most likely bad username/password, exiting..."
			exit 1
		}
		$versionOut = "ucs-k9-bundle*" + $images[0].Version + "*"
		$VERSION = $images[0].Version
	}
	elseif ($VERSION)
	{
		$versionOut = "ucs-k9-bundle*" + ($version).ToLower() + "*"
	}
	else
	{
		Write-Host "No Images found, no version identified"
		exit 1
	}
	if ($IncludeCSeries -eq $true)
	{
		try
		{
			$ccoimages = Get-UcsCcoImageList -Credential $ccocred -ErrorAction SilentlyContinue| where { ($_.ImageName -ilike $versionOut) -or ($_.Version -ilike ($version).ToLower()) }
		}
		catch
		{
			Write-Host -ForegroundColor Red	"Unable to connect to cisco.com for images.  Most likely bad username/password, exiting..."
			exit 1
		}
	}
	else
	{
		try
		{
			$ccoimages = Get-UcsCcoImageList -Credential $ccocred -ErrorAction SilentlyContinue | where { (($_.ImageName -like $versionOut) -or ($_.Version -like ($version).ToLower())) -and ($_.ImageName -notmatch "c-series") }
		}
		catch
		{
			Write-Host -ForegroundColor Red	"Unable to connect to cisco.com for images.  Most likely bad username/password, exiting..."
			exit 1
		}
	}
	if ($ccoImages)
	{
		foreach ($ccoImage in $ccoImages)
		{
			if (!(Get-Item $($rootDir + $ccoImage.ImageName) -ErrorAction SilentlyContinue))
			{
				Write-Host "Image $($ccoImage.ImageName): Downloading..." -NoNewline
				$ccoImage | Get-UcsCcoImage -Path $rootDir
				Write-Host "...Complete."
			}
			else
			{
				Write-Host "Image $($ccoImage.ImageName): Already Downloaded.  Skipping."
			}
		}
	}
	else
	{
		Write-Host "Error: Failed to get images from CCO."
		Write-Host "Possible Causes:"
		Write-Host "	- The CCO ID you used is invalid"
		Write-Host "	- The version you have used is invalid.  See Help for proper formatting. Version Format = x.x(xx)"
		Write-Host "Exiting..."
		exit 1
	}
	if ($Drivers)
	{
		
		$ISOPackages = Get-UcsCcoImageList -Credential $ccocred | where {$_.ImageName -match "bxxx-drivers."} | sort Version -Descending
		if (!(Get-Item $($rootDir + $ISOPackages[0].ImageName) -ErrorAction SilentlyContinue))
		{
			Write-Host "B-Series Drivers ISO: $($ISOPackages[0].ImageName): Downloading..." -NoNewline
			$ISOPackages[0] | Get-UcsCcoImage -Path $rootDir
			Write-Host "...Complete."
		}
		else
		{
			Write-Host "B-Series Drivers ISO: $($ISOPackages[0].ImageName): Already Downloaded.  Skipping."
		}
		if ($IncludeCSeries)
		{
			$ISOPackages = Get-UcsCcoImageList -Credential $ccocred | where {$_.ImageName -match "cxxx-drivers."} | sort Version -Descending
			if (!(Get-Item $($rootDir + $ISOPackages[0].ImageName) -ErrorAction SilentlyContinue))
			{
				Write-Host "C-Series Drivers ISO: $($ISOPackages[0].ImageName): Downloading..." -NoNewline
				$ISOPackages[0] | Get-UcsCcoImage -Path $rootDir
				Write-Host "...Complete."
			}
			else
			{
				Write-Host "C-Series Drivers ISO: $($ISOPackages[0].ImageName): Already Downloaded.  Skipping."
			}
		}
	}
	return $ccoimages
}
Function Set-UcsFirmwarePackages ($ucsHandleList, $images, $rootDir)
{
	if ($ucsHandleList -ne $null)
	{
		$Total = 0
		foreach ($image in $images)
		{
			$Total += $image.Size / 1000000
		}
		foreach ($handle in $ucsHandleList)
		{
			Write-Host "Checking for space:" -NoNewline
			[array]$ucsmStorage = Get-UcsNetworkElement -Ucs $handle | Get-UcsStorageItem | where {$_.Name -eq "bootflash"}
			[int]$freespace = ([int]$ucsmStorage[0].Size * ((100 - [int]$ucsmStorage[0].Used)/100)) + ([int]$ucsmStorage[1].Size * ((100 - [int]$ucsmStorage[1].Used)/100))
			if (($Total / $freespace) -ge .8) # Do not take greater then 80% of total available space.
			{ 
				Write-Host "$($handle.Name): Not enough free space to apply image, please free up bootflash by removing older firmware images." -ForegroundColor Red
				$ucsHandleList = $ucsHandleList | where {$_.Name -ne $handle.Name}
				Disconnect-Ucs -Ucs $handle
			}
			else 
			{
				$avail = (100 - (($Total / $freespace) * 100))
				Write-Host "$($handle.Name): Free Space ok:" + $avail + "%"
			}
		}
		if ($ucsHandleList -ne $null)
		{
			foreach ($image in $images)
			{
				
				Write-Host "Image $($image.ImageName): Uploading to UCSM/s..." -NoNewline
				$outputName = $rootDir + "$($image.ImageName)"
				Send-UcsFirmware -ucs $ucsHandleList -LiteralPath $outputName -ErrorAction SilentlyContinue | Out-Null
				Write-Host "..uploaded."
			}
		}
		else
		{
			Write-Host -ForegroundColor Red "No UCSM Domains passed storage availability checks.  Please clear storage space."
			exit 1
		}
	}
	else
	{
		Write-Host "Firwmware was downloaded, but no valid UCS domains were passed into the script.  Exiting..."
		exit 1
	}
}
function Get-OutputDirectory ($OUTPUTDIR)
{
	if (!($OUTPUTDIR))
	{
		$Dir = "C:\UCSMImages\"
	}
	if (!(Get-Item $Dir -ErrorAction SilentlyContinue))
	{
		if (!($Dir.EndsWith("\")))
		{
			$dir = $dir + "\"
		}
		md -Path $Dir -Force
		if (!(Get-Item $Dir))
		{
			Write-Host "Could not create destination directory.  Most likely permissions."
			Write-Host "Exiting..."
			exit 1
		}
	}
	return $Dir
}
#_______________________________________________________________________________
#__________________MAIN PROGRAM ________________________________________________
#_______________________________________________________________________________
Validate-UcsPSLibrary # Check if you have the UCS Powershell Library Installed
$UCSMList = Get-UCSMList

$ucsHandleList = @()
foreach ($ucsm in $ucsmList)
{
	if ($ucsm)
	{
		$ucsHandleList += Validate-UcsFIConnect $ucsm
	}
	else
	{
		Write-Host "No UCSM instances defined, exiting."
	}
}

$rootDir = Get-OutputDirectory $OUTPUTDIR
$imagesOut = Get-UcsFirmwarePackages $rootDir
Set-UcsFirmwarePackages $ucsHandleList $imagesOut $rootDir	

