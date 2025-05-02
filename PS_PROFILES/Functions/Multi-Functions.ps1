####################
# Static Functions #
####################

# Clean Mac address
function Clean-MacAddress
{
<#
	.SYNOPSIS
		Function to cleanup a MACAddress string
	
	.DESCRIPTION
		Function to cleanup a MACAddress string
	
	.PARAMETER MacAddress
		Specifies the MacAddress
	
	.PARAMETER Separator
		Specifies the separator every two characters
	
	.PARAMETER Uppercase
		Specifies the output must be Uppercase
	
	.PARAMETER Lowercase
		Specifies the output must be LowerCase
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:33:44:55'
	
		001122334455
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Uppercase
	
		001122DDEEFF
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase
	
		001122ddeeff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator '-'
	
		00-11-22-dd-ee-ff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator '.'
	
		00.11.22.dd.ee.ff
	
	.EXAMPLE
		Clean-MacAddress -MacAddress '00:11:22:dD:ee:FF' -Lowercase -Separator :
	
		00:11:22:dd:ee:ff
	
	.OUTPUTS
		System.String
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	[OutputType([String], ParameterSetName = "Upper")]
	[OutputType([String], ParameterSetName = "Lower")]
	[CmdletBinding(DefaultParameterSetName = 'Upper')]
	param
	(
		[Parameter(ParameterSetName = 'Lower')]
		[Parameter(ParameterSetName = 'Upper')]
		[String]$MacAddress,
		
		[Parameter(ParameterSetName = 'Lower')]
		[Parameter(ParameterSetName = 'Upper')]
		[ValidateSet(':', 'None', '.', "-")]
		$Separator,
		
		[Parameter(ParameterSetName = 'Upper')]
		[Switch]$Uppercase,
		
		[Parameter(ParameterSetName = 'Lower')]
		[Switch]$Lowercase
	)
	
	BEGIN
	{
		# Initial Cleanup
		$MacAddress = $MacAddress -replace "-", "" #Replace Dash
		$MacAddress = $MacAddress -replace ":", "" #Replace Colon
		$MacAddress = $MacAddress -replace "/s", "" #Remove whitespace
		$MacAddress = $MacAddress -replace " ", "" #Remove whitespace
		$MacAddress = $MacAddress -replace "\.", "" #Remove dots
		$MacAddress = $MacAddress.trim() #Remove space at the beginning
		$MacAddress = $MacAddress.trimend() #Remove space at the end
	}
	PROCESS
	{
		IF ($PSBoundParameters['Uppercase'])
		{
			$MacAddress = $macaddress.toupper()
		}
		IF ($PSBoundParameters['Lowercase'])
		{
			$MacAddress = $macaddress.tolower()
		}	
		IF ($PSBoundParameters['Separator'])
		{
			IF ($Separator -ne "None")
			{
				$MacAddress = $MacAddress -replace '(..(?!$))', "`$1$Separator"
			}
		}
	}
	END
	{
		Write-Output $MacAddress
	}
}
# End Clean Mac Address
# Get-IPAddress
function Get-IPAddress
{
	Get-NetIPAddress | ?{($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*vmware*") -and ($_.interfacealias -notlike "*loopback*") -and ($_.interfacealias -notlike "*bluetooth*") -and ($_.interfacealias -notlike "*isatap*")} | ft
}
# Reload Profile
function Reload-Profile {
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost
    ) | % {
        if(Test-Path $_){
            Write-Verbose "Running $_"
            . $_
        }
    }    
}
# End Get-IPAddress


# DelProf2 User profiles
function Remove-UserProfiles {

<#
.SYNOPSIS
    Written by: JBear 1/31/2017
	
    Remove user profiles from a specified system.

.DESCRIPTION
    Remove user profiles from a specified system with the use of DelProf2.exe.

.EXAMPLE
    Remove-UserProfiles Computer123456

        Note: Follow instructions and prompts to completetion.

#>

    param(
        [parameter(mandatory=$true)]
        [string[]]$computername
    )

　
    function UseDelProf2 { 
               
        #Set parameters for remote computer and -WhatIf (/l)
        $WhatIf = @(

            "/l",
            "/c:$computer" 
        )
           
        #Runs DelProf2.exe with the /l parameter (or -WhatIf) to list potential User Profiles tagged for potential deletion
        & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $WhatIf

        #Display instructions on console
        Write-Host "`n`nPLEASE ENSURE YOU FULLY UNDERSTAND THIS COMMAND BEFORE USE `nTHIS WILL DELETE ALL USER PROFILE INFORMATION FOR SPECIFIED USER(S) ON THE SPECIFIED WORKSTATION!`n"

        #Prompt User for input
        $DeleteUsers = Read-Host -Prompt "To delete User Profiles, please use the following syntax ; Wildcards (*) are accepted. `nExample: /id:user1 /id:smith* /id:*john*`n `nEnter proper syntax to remove specific users" 

        #If only whitespace or a $null entry is entered, command is not run
        if([string]::IsNullOrWhiteSpace($DeleteUsers)) {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"

        }

        #If Read-Host contains proper syntax (Starts with /id:) run command to delete specified user; DelProf will give a confirmation prompt
        elseif($DeleteUsers -like "/id:*") {

            #Set parameters for remote computer
            $UserArgs = @(

                "/c:$computer"
            )

            #Split $DeleteUsers entries and add to $UserArgs array
            $UserArgs += $DeleteUsers.Split("")

            #Runs DelProf2.exe with $UserArgs parameters (i.e. & "C:\DelProf2.exe" /c:Computer1 /id:User1* /id:User7)
            & "C:\LazyWinAdmin\Win32-Tools\DelProf2.exe" $UserArgs
        }

        #If Read-Host doesn't begin with the input /id:, command is not run
        else {

            Write-Host "`nImproper value entered, excluding all users from deletion. You will need to re-run the command on $computer, if you wish to try again...`n"
        }
    }

    foreach($computer in $computername) {
        if(Test-Connection -Quiet -Count 1 -Computer $Computer) { 

            UseDelProf2 
        }

        else {
            
            Write-Host "`nUnable to connect to $computer. Please try again..." -ForegroundColor Red
        }

    }
}#End Remove-UserProfiles

# Remove-RemotePrintDrivers
function Remove-RemotePrintDrivers {
  <# 
  .SYNOPSIS 
  Remove printer drivers from registry of specified workstation(s) 

  .EXAMPLE 
  Remove-RemotePrintDrivers Computer123456 

  .EXAMPLE 
  Remove-RemotePrintDrivers 123456 
  #> 
	param([Parameter(Mandatory=$true)]
	[string[]]$computername)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

function RmPrintDrivers {

$i=0
$j=0
 	
foreach ($Computer in $ComputerName) { 

    Write-Progress -Activity "Clearing printer drivers..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

　
	Try {

		$RemoteSession = New-PSSession -ComputerName $Computer
}
	Catch {

		"Something went wrong. Unable to connect to $Computer"
		Break
}
	Invoke-Command -Session $RemoteSession -ScriptBlock {
    # Removes print drivers, other than default image drivers
		if ((Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\') -eq $true) {
			Remove-Item -PATH 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3\*' -EXCLUDE "*ADOBE*", "*MICROSOFT*", "*XPS*", "*REMOTE*", "*FAX*", "*ONENOTE*" -recurse
			Remove-Item -PATH 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\*' -EXCLUDE "*ADOBE*", "*MICROSOFT*", "*XPS*", "*REMOTE*", "*FAX*", "*ONENOTE*" -recurse
		Set-Service Spooler -startuptype manual
		Restart-Service Spooler
		Set-Service Spooler -startuptype automatic
			}
		} -AsJob -JobName "ClearPrintDrivers"
	} 
} RmPrintDrivers | Wait-Job | Remove-Job

Remove-PSSession *

[Void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$RMprintConfirmation = [Microsoft.VisualBasic.Interaction]::MsgBox("Printer driver removal triggered on workstation(s)!", "OKOnly,SystemModal,Information", "Success")

}#End Remove-RemotePrintDrivers

# Remote Desktop Protocol
function RDP {
  <# 
  .SYNOPSIS 
  Remote Desktop Protocol to specified workstation(s) 

  .EXAMPLE 
  RDP Computer123456 

  .EXAMPLE 
  RDP 123456 
  #> 
	param(
	[Parameter(Mandatory=$true)]
	[string]$computername)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

	#Start Remote Desktop Protocol on specifed workstation
	& "C:\windows\system32\mstsc.exe" /v:$computername /fullscreen
}# End RDP

# Get-LastReboot
function Get-LastBoot {
  <# 
  .SYNOPSIS 
  Retrieve last restart time for specified workstation(s) 

  .EXAMPLE 
  Get-LastBoot Computer123456 

  .EXAMPLE 
  Get-LastBoot 123456 
  #> 
    param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$i=0
$j=0

    foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

　
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{
            "Computer Name" = $Computer
            "Last Reboot"= $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}# End Get-LastBoot


# Get-LoggedOnUser
function Get-LoggedOnUser{
  <# 
  .SYNOPSIS 
  Retrieve current user logged into specified workstations(s) 

  .EXAMPLE 
  Get-LoggedOnUser Computer123456 

  .EXAMPLE 
  Get-LoggedOnUser 123456 
  #> 
	Param([Parameter(Mandatory=$true)]
	[string[]] $ComputerName)
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}
	write-host("")
	write-host("Gathering resources. Please wait...")
	write-host("")

    $i=0
    $j=0

    foreach($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Logged On User..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer

        Write-Host "User Logged In: " $computerSystem.UserName "`n"
    }
}#End Get-LoggedOnUser

# Get-HotFixes
function Get-HotFixes {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  Get-HotFixes Computer123456 

  .EXAMPLE 
  Get-HotFixes 123456 
  #> 
param (
    [Parameter(ValueFromPipeline=$true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [string]$NameRegex = '')

if(($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        $computername = "Computer" + $computername.Replace("Computer","")
    }	
}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

function HotFix {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving HotFix Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        Get-HotFix -Computername $computer 
    }    
}

foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}

$HotFix = HotFix
$DocPath = [environment]::getfolderpath("mydocuments") + "\HotFix-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $HotFix | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $HotFix | Out-GridView -Title "HotFix Report"; }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp HotFixes output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
}# End Get-HotFixes

# Get-RemoteGroup Policies

function Get-GPRemote {
  <# 
  .SYNOPSIS 
  Open Group Policy for specified workstation(s) 

  .EXAMPLE 
  Get-GPRemote Computer123456 

  .EXAMPLE 
  Get-GPRemote 123456 
  #> 
param(
[Parameter(Mandatory=$true)]
[string[]] $ComputerName)

if (($computername.length -eq 6)) {
    [int32] $dummy_output = $null;

    if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
       	$computername = "Computer" + $computername.Replace("Computer","")}	
}

$i=0
$j=0

foreach ($Computer in $ComputerName) {

    Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	#Opens (Remote) Group Policy for specified workstation
	gpedit.msc /gpcomputer: $Computer
    
	}
}#End Get-GPRemote

# Get Remote Processes

function CheckProcess {
  <# 
  .SYNOPSIS 
  Grabs all processes on specified workstation(s).

  .EXAMPLE 
  CheckProcess Computer123456 

  .EXAMPLE 
  CheckProcess 123456 
  #> 
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = '')
	if (($computername.length -eq 6)) {
    		[int32] $dummy_output = $null;

    	if ([int32]::TryParse($computername , [ref] $dummy_output) -eq $true) {
        	$computername = "Computer" + $computername.Replace("Computer","")}	
	}

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

function ChkProcess {

$i=0
$j=0

    foreach ($computer in $ComputerArray) {

        Write-Progress -Activity "Retrieving System Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        $getProcess = Get-Process -ComputerName $computer

        foreach ($Process in $getProcess) {
                
             [pscustomobject]@{
		"Computer Name" = $computer
                "Process Name" = $Process.ProcessName
                PID = '{0:f0}' -f $Process.ID
                Company = $Process.Company
                "CPU(s)" = $Process.CPU
                Description = $Process.Description
             }           
         }
     } 
}
	
foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}
	$chkProcess = ChkProcess | Sort "Computer Name" | Select "Computer Name","Process Name", PID, Company, "CPU(s)", Description
    	$DocPath = [environment]::getfolderpath("mydocuments") + "\Process-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $chkProcess | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $chkProcess | Out-GridView -Title "Processes";  }
    		}

	if($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp Check Process output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
    
}#End CheckProcess

# Whois Check
<#
.SYNOPSIS
Domain name WhoIs
.DESCRIPTION
Performs a domain name lookup and returns information such as
domain availability (creation and expiration date),
domain ownership, name servers, etc..

.PARAMETER domain
Specifies the domain name (enter the domain name without http:// and www (e.g. power-shell.com))

.EXAMPLE
WhoIs -domain power-shell.com 
whois power-shell.com

.NOTES
File Name: whois.ps1
Author: Nikolay Petkov
Blog: http://power-shell.com
Last Edit: 12/20/2014

.LINK
http://power-shell.com
#>
Function WhoIs {
param (
                [Parameter(Mandatory=$True,
                           HelpMessage='Please enter domain name (e.g. microsoft.com)')]
                           [string]$domain
        )
Write-Host "Connecting to Web Services URL..." -ForegroundColor Green
try {
#Retrieve the data from web service WSDL
If ($whois = New-WebServiceProxy -uri "http://www.webservicex.net/whois.asmx?WSDL") {Write-Host "Ok" -ForegroundColor Green}
else {Write-Host "Error" -ForegroundColor Red}
Write-Host "Gathering $domain data..." -ForegroundColor Green
#Return the data
(($whois.getwhois("=$domain")).Split("<<<")[0])
} catch {
Write-Host "Please enter valid domain name (e.g. microsoft.com)." -ForegroundColor Red}
} #end function WhoIs

###########################
#
# netstat
# http://blogs.microsoft.co.il/blogs/scriptfanatic/archive/2011/02/10/How-to-find-running-processes-and-their-port-number.aspx
# Get-NetworkStatistics | where-object {$_.State -eq "LISTENING"} | Format-Table
###########################

function Get-NetworkStatistics 
{ 
    $properties = 'Protocol','LocalAddress','LocalPort' 
    $properties += 'RemoteAddress','RemotePort','State','ProcessName','PID' 

    netstat -ano | Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object { 

        $item = $_.line.split(" ",[System.StringSplitOptions]::RemoveEmptyEntries) 

        if($item[1] -notmatch '^\[::') 
        {            
            if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $localAddress = $la.IPAddressToString 
               $localPort = $item[1].split('\]:')[-1] 
            } 
            else 
            { 
                $localAddress = $item[1].split(':')[0] 
                $localPort = $item[1].split(':')[-1] 
            }  

            if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6') 
            { 
               $remoteAddress = $ra.IPAddressToString 
               $remotePort = $item[2].split('\]:')[-1] 
            } 
            else 
            { 
               $remoteAddress = $item[2].split(':')[0] 
               $remotePort = $item[2].split(':')[-1] 
            }  

            New-Object PSObject -Property @{ 
                PID = $item[-1] 
                ProcessName = (Get-Process -Id $item[-1] -ErrorAction SilentlyContinue).Name 
                Protocol = $item[0] 
                LocalAddress = $localAddress 
                LocalPort = $localPort 
                RemoteAddress =$remoteAddress 
                RemotePort = $remotePort 
                State = if($item[0] -eq 'tcp') {$item[3]} else {$null} 
            } | Select-Object -Property $properties 
        } 
    } 
}
# End Get-NetworkStatistics

# Update SysInternals
<#
.Synopsis
   Download the latest sysinternals tools
.DESCRIPTION
   Downloads the latest sysinternals tools from https://live.sysinternals.com/ to a specified directory
   The function downloads all .exe and .chm files available
.EXAMPLE
   Update-Sysinternals -Path C:\sysinternals
   Downloads the sysinternals tools to the directory C:\sysinternals
.EXAMPLE
   Update-Sysinternals -Path C:\Users\Matt\OneDrive\Tools\sysinternals
   Downloads the sysinternals tools to a user's OneDrive
#>
function Update-Sysinternals {
    [CmdletBinding()]
    param (
        # Path to the directory were sysinternals tools will be downloaded to 
        [Parameter(Mandatory=$true)]      
        [string]
        $Path 
    )
    
    begin {
            if (-not (Test-Path -Path $Path)){
            Throw "The Path $_ does not exist"
        } else {
            $true
        }
        
            $uri = 'https://live.sysinternals.com/'
            $sysToolsPage = Invoke-WebRequest -Uri $uri
            
    }
    
    process {
        # create dir if it doesn't exist    
       
        Set-Location -Path $Path

        $sysTools = $sysToolsPage.Links.innerHTML | Where-Object -FilterScript {$_ -like "*.exe" -or $_ -like "*.chm"} 

        foreach ($sysTool in $sysTools){
            Invoke-WebRequest -Uri "$uri/$sysTool" -OutFile $sysTool
        }
    } #process
}
# End Update SysInternals

<#
#Function Get-UpTime
#{
#    param([string] $LastBootTime)
#    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
#    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
#}
function Get-Uptime {
# Accept input from the pipeline
Param([Parameter(mandatory=$true,ValueFromPipeline=$true)] [string[]]$ComputerName = @("."))

# Process the piped input (one computer at a time)
process { 

    # See if it responds to a ping, otherwise the WMI queries will fail
    $query = "select * from win32_pingstatus where address = '$ComputerName'"
    $ping = Get-WmiObject -query $query
if ($ping.protocoladdress) {
    # Ping responded, so connect to the computer via WMI
    $os = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ev myError -ea SilentlyContinue 

if ($myError -ne $null)
 {
  # Error: WMI did not respond
  "$ComputerName did not respond"
 } 
else
 { 
   $LastBootUpTime = $os.ConvertToDateTime($os.LastBootUpTime)
   $LocalDateTime = $os.ConvertToDateTime($os.LocalDateTime)
   
   # Calculate uptime - this is automatically a timespan
   $up = $LocalDateTime - $LastBootUpTime

   # Split into Days/Hours/Mins
   $uptime = "$($up.Days) days, $($up.Hours)h, $($up.Minutes)mins" 

   # Save the results for this computer in an object
   $results = new-object psobject
   $results | add-member noteproperty LastBootUpTime $LastBootUpTime
   $results | add-member noteproperty ComputerName $os.csname
   $results | add-member noteproperty uptime $uptime

   # Display the results
   $results | Select-Object ComputerName,LastBootUpTime, uptime
 }

# Next Ping result
}

# End of the process block
}}
#>

# End Get-Update
# Get AD GPO Replication
function Get-ADGPOReplication
{
	<#
	.SYNOPSIS
		This function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.DESCRIPTION
		This function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.PARAMETER GPOName
		Specify the name of the GPO
	.PARAMETER All
		Specify that you want to retrieve all the GPO (slow if you have a lot of Domain Controllers)
	.EXAMPLE
		Get-ADGPOReplication -GPOName "Default Domain Policy"
	.EXAMPLE
		Get-ADGPOReplication -All
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		lazywinadmin.com
	
		VERSION HISTORY
		1.0 2014.09.22 	Initial version
						Adding some more Error Handling
						Fix some typo
	#>
	#requires -version 3

	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $True, ParameterSetName = "One")]
		[String[]]$GPOName,
		[parameter(Mandatory = $True, ParameterSetName = "All")]
		[Switch]$All
	)
Remove-Module Carbon
	BEGIN
	{
		TRY
		{
			if (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginIpmoAD }
			if (-not (Get-Module -Name GroupPolicy)) { Import-Module -Name GroupPolicy -ErrorAction Stop -ErrorVariable ErrorBeginIpmoGP }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			IF ($ErrorBeginIpmoAD) { Write-Warning -Message "[BEGIN] Error while Importing the module Active Directory" }
			IF ($ErrorBeginIpmoGP) { Write-Warning -Message "[BEGIN] Error while Importing the module Group Policy" }
			Write-Warning -Message "[BEGIN] $($Error[0].exception.message)"
		}
	}
	PROCESS
	{
		FOREACH ($DomainController in ((Get-ADDomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetDC -filter *).hostname))
		{
			TRY
			{
				IF ($psBoundParameters['GPOName'])
				{
					Foreach ($GPOItem in $GPOName)
					{
						$GPO = Get-GPO -Name $GPOItem -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPO
						
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPOItem
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}#Foreach ($GPOItem in $GPOName)
				}#IF ($psBoundParameters['GPOName'])
				IF ($psBoundParameters['All'])
				{
					$GPOList = Get-GPO -All -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPOAll
					
					foreach ($GPO in $GPOList)
					{
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPO.DisplayName
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}
				}#IF ($psBoundParameters['All'])
			}#TRY
			CATCH
			{
				Write-Warning -Message "[PROCESS] Something wrong happened"
				IF ($ErrorProcessGetDC) { Write-Warning -Message "[PROCESS] Error while running retrieving Domain Controllers with Get-ADDomainController" }
				IF ($ErrorProcessGetGPO) { Write-Warning -Message "[PROCESS] Error while running Get-GPO" }
				IF ($ErrorProcessGetGPOAll) { Write-Warning -Message "[PROCESS] Error while running Get-GPO -All" }
				Write-Warning -Message "[PROCESS] $($Error[0].exception.message)"
			}
		}#FOREACH
	}#PROCESS
}
# End Get-ADGPOReplication
# Test Registry Value
function Test-RegistryValue {

param (

 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,

[parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

try {

Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
 return $true
 }

catch {

return $false

}

}
### End Test-RegistryValue
# Get Local Admins 
function Get-LocalAdmin { 
param ($ComputerName) 
 
$admins = Gwmi win32_groupuser –computer $ComputerName  
$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'} 
 
$admins |% { 
$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul 
$matches[1].trim('"') + “\” + $matches[2].trim('"') 
} 
}
### End Get-LocalAdmin
####################
# Static Functions #
####################
# Touch
function touch { $args | foreach-object {write-host > $_} }
# Notepad++
function NPP { Start-Process -FilePath "${Env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" }#-ArgumentList $args }
# Find File
function findfile($name) {
	ls -recurse -filter "*${name}*" -ErrorAction SilentlyContinue | foreach {
		$place_path = $_.directory
		echo "${place_path}\${_}"
	}
}
# RM -RF
function rm-rf($item) { Remove-Item $item -Recurse -Force }
# SUDO
function sudo(){
	Invoke-Elevated @args
}
# SED
function PSsed($file, $find, $replace){
	(Get-Content $file).replace("$find", $replace) | Set-Content $file
}
# SED-Recursive
function PSsed-recursive($filePattern, $find, $replace) {
	$files = ls . "$filePattern" -rec # -Exclude
	foreach ($file in $files) {
		(Get-Content $file.PSPath) |
		Foreach-Object { $_ -replace "$find", "$replace" } |
		Set-Content $file.PSPath
	}
}
# PSGrep
function PSgrep {

    [CmdletBinding()]
    Param(
    
        # source file to grep
        [Parameter(Mandatory=$true)]
        [string]$SourceFileName, 

        # string to search for
        [Parameter(Mandatory=$true)]
        [string]$SearchStrings,

        # do we write to file
        [Parameter()]
        [string]$OutputFile
    )

        # break the comma separated strings up
        $Strings = @()
        $Strings = $SearchStrings.split(',')
        $count = 0

        # write-host $Strings

        $Content = Get-Content $SourceFileName
        
        $Content | ForEach-Object { 
            foreach ($String in $Strings) {
                # $String
                if($_ -match $String){
                    $count ++
                    if (!($OutputFile)) {
                        write-host $_
                    } else {
                        $_ | Out-File -FilePath ".\$($OutputFile)" -Append -Force
                }

            }

        }

    }

    Write-Host "$($Count) matches found"
}
# End PSgrep
# Which
function which($name) {
	Get-Command $name | Select-Object -ExpandProperty Definition
}
# Cut
function cut(){
	foreach ($part in $input) {
		$line = $part.ToString();
		$MaxLength = [System.Math]::Min(200, $line.Length)
		$line.subString(0, $MaxLength)
	}
}
# Search Text Files
function Search-AllTextFiles {
    param(
        [parameter(Mandatory=$true,position=0)]$Pattern, 
        [switch]$CaseSensitive,
        [switch]$SimpleMatch
    );

    Get-ChildItem . * -Recurse -Exclude ('*.dll','*.pdf','*.pdb','*.zip','*.exe','*.jpg','*.gif','*.png','*.ico','*.svg','*.bmp','*.psd','*.cache','*.doc','*.docx','*.xls','*.xlsx','*.dat','*.mdf','*.nupkg','*.snk','*.ttf','*.eot','*.woff','*.tdf','*.gen','*.cfs','*.map','*.min.js','*.data') | Select-String -Pattern:$pattern -SimpleMatch:$SimpleMatch -CaseSensitive:$CaseSensitive
}
# Add to Zip
function AddTo-7zip($zipFileName) {
    BEGIN {
        #$7zip = "$($env:ProgramFiles)\7-zip\7z.exe"
        $7zip = Find-Program "\7-zip\7z.exe"
		if(!([System.IO.File]::Exists($7zip))){
			throw "7zip not found";
		}
    }
    PROCESS {
        & $7zip a -tzip $zipFileName $_
    }
    END {
    }
}
## End Add to Zip

# Connect to Exchange
Function GoGo-PSExch {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="USONVSVREX01"
    )
    
    #$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos #-Credential #$Credentials

    Import-PSSession $ExOPSession -AllowClobber

}
## End Connect to Exchange
## Out-File in UTF8 NonBom
function Out-FileUtf8NoBom {

  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position=0)] [string] $LiteralPath,
    [switch] $Append,
    [switch] $NoClobber,
    [AllowNull()] [int] $Width,
    [Parameter(ValueFromPipeline)] $InputObject
  )

  #requires -version 3

  # Make sure that the .NET framework sees the same working dir. as PS
  # and resolve the input path to a full path.
  [System.IO.Directory]::SetCurrentDirectory($PWD) # Caveat: .NET Core doesn't support [Environment]::CurrentDirectory
  $LiteralPath = [IO.Path]::GetFullPath($LiteralPath)

  # If -NoClobber was specified, throw an exception if the target file already
  # exists.
  if ($NoClobber -and (Test-Path $LiteralPath)) {
    Throw [IO.IOException] "The file '$LiteralPath' already exists."
  }

  # Create a StreamWriter object.
  # Note that we take advantage of the fact that the StreamWriter class by default:
  # - uses UTF-8 encoding
  # - without a BOM.
  $sw = New-Object IO.StreamWriter $LiteralPath, $Append

  $htOutStringArgs = @{}
  if ($Width) {
    $htOutStringArgs += @{ Width = $Width }
  }

  # Note: By not using begin / process / end blocks, we're effectively running
  #       in the end block, which means that all pipeline input has already
  #       been collected in automatic variable $Input.
  #       We must use this approach, because using | Out-String individually
  #       in each iteration of a process block would format each input object
  #       with an indvidual header.
  try {
    $Input | Out-String -Stream @htOutStringArgs | % { $sw.WriteLine($_) }
  } finally {
    $sw.Dispose()
  }

}
## End Out-File in UTF8 NonBom

## Invoke VBScript
Function Invoke-VBScript {
    <#
    .Synopsis
       Run VBScript from PowerShell
    .DESCRIPTION
       Used to invoke VBScript from PowerShell

       Will run the VBScript in a separate job using cscript.exe
    .PARAMETER Path
       Path to VBScript.
       Accepts relative or absolute path.
    .PARAMETER Argument
       Arguments to pass to VBScript
    .PARAMETER Wait
       Wait for VBScript to finish   
    .EXAMPLE
       Invoke-VBScript -Path '.\VBScript1.vbs' -Arguments '"MyFirstArgument"', '"MySecondArgument"' -Wait
       Run VBScript1.vbs using cscript and wait for the script to complete.
       Displays progressbar while waiting.
       Returns script output as single string.
    .EXAMPLE
       '.\VBScript1.vbs', '.\VBScript2.vbs' | Invoke-VBScript -Arguments '"MyArgument"'
       Starts both VBScript1.vbs and VBScript2.vbs in separate jobs simultaneously.
       Both scripts will be run using the same arguments.
       Returns job items.
    .EXAMPLE
       [PSCustomObject]@{Path='.\VBScript1.vbs';Arguments='"Script1"'},[PSCustomObject]@{Path='.\VBScript2.vbs';Arguments='"Script2"'} | Invoke-VBScript -Wait -Verbose
       Runs two scripts after each other, waiting to one to complete
       before starting next.
       Each script will run with different parameters.
       Displays progressbar while waiting.
       Returns script output in one single string per script.
    .NOTES
       Written by Simon Wåhlin
       http://blog.simonw.se
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='None',PositionalBinding=$false)]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateScript({if(Test-Path $_){$true}else{Throw "Could not find script: [$_]"}})]
        [String]
        $Path,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Args')]
        [String[]]
        $Argument,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]
        $Wait
    )
    Begin
    {
        Write-Verbose -Message 'Locating cscript.exe'
        $cscriptpath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\cscript.exe'
        if(-Not(Test-Path -Path $cscriptpath))
        {
            Throw 'cscript.exe not found.'
        }
        Write-Verbose -Message ('cscript.exe found in: {0}' -f $cscriptpath)
    }
    Process
    {
        Try
        {
            $ResolvedPath = Resolve-Path -Path $Path
            Write-Verbose -Message ('Processing script: {0}' -f $ResolvedPath)
            if($PSBoundParameters.ContainsKey('Argument'))
            {
                $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}" "{2}"' -f $cscriptpath, $ResolvedPath,($Argument -join '" "')))
            }
            else
            {
                $ScriptBlock = $ScriptBlock = [scriptblock]::Create(('& "{0}" "{1}"' -f $cscriptpath, $ResolvedPath))
            }
            Write-Verbose -Message 'Starting script'
            if($PSCmdlet.ShouldProcess($ResolvedPath,'Invoke script'))
            {
                $Job = Start-Job -ScriptBlock $ScriptBlock
                if($Wait)
                {
                    $Activity = 'Waiting for script to complete: {0}' -f $ResolvedPath
                    Write-Progress -Activity $Activity -Id 1
                    $i = 1
                    While($Job.State -eq 'Running')
                    {
                        $WaitTime = (Get-Date) - $Job.PSBeginTime
                        Write-Progress -Activity $Activity -Status "Waited for $($WaitTime.TotalSeconds -as [int]) seconds." -Id 1 -PercentComplete ($i%100)
                        Start-Sleep -Seconds 1
                        $i++
                    }
                    Write-Progress -Activity $Activity -Status 'Waiting' -Id 1 -Completed
                    $Result = Foreach($JobInstance in ($Job,$Job.ChildJobs))
                    {
                        if($JobInstance.Error -ne $null)
                        {
                            Throw $JobInstance.Error.Exception.Message
                        }
                        else
                        {
                            $JobInstance.Output
                        }
                    }
                    Write-Output -InputObject ($Result -join "`n")
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
                }
                else
                {
                    Write-Output -InputObject $Job
                }
            }
            Write-Verbose -Message 'Finished processing script'
        }
        Catch
        {
            Throw
        }
    }
}
## End Invoke VBScript
## Function Get-MOTD
Function Get-MOTD {

<#
.NAME
    Get-MOTD
.SYNOPSIS
    Displays system information to a host.
.DESCRIPTION
    The Get-MOTD cmdlet is a system information tool written in PowerShell. 
.EXAMPLE
#>


  [CmdletBinding()]
	
  Param(
    [Parameter(Position=0,Mandatory=$false)]
	[ValidateNotNullOrEmpty()]
    [string[]]$ComputerName
    ,
    [Parameter(Position=1,Mandatory=$false)]
    [PSCredential]
    [System.Management.Automation.CredentialAttribute()]$Credential
  )

  Begin {
	
        If (-Not $ComputerName) {
            $RemoteSession = $null
        }
        #Define ScriptBlock for data collection
        $ScriptBlock = {
            $Operating_System = Get-CimInstance -ClassName Win32_OperatingSystem
            $Logical_Disk = Get-CimInstance -ClassName Win32_LogicalDisk |
            Where-Object -Property DeviceID -eq $Operating_System.SystemDrive
			Try {
				$PCLi = Get-PowerCLIVersion
				$PCLiVer = ' | PowerCLi ' + [string]$PCLi.Major + '.' + [string]$PCLi.Minor + '.' + [string]$PCLi.Revision + '.' + [string]$PCLi.Build
			} Catch {$PCLiVer = ''}
			If ($DomainName = ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()).DomainName) {$DomainName = '.' + $DomainName}
			
            [pscustomobject]@{
                Operating_System = $Operating_System
                Processor = Get-CimInstance -ClassName Win32_Processor
                Process_Count = (Get-Process).Count
                Shell_Info = ("{0}.{1}" -f $PSVersionTable.PSVersion.Major,$PSVersionTable.PSVersion.Minor) + $PCLiVer
                Logical_Disk = $Logical_Disk
            }
        }
  } #End Begin

  Process {
	
        If ($ComputerName) {
            If ("$ComputerName" -ne "$env:ComputerName") {
                # Build Hash to be used for passing parameters to 
                # New-PSSession commandlet
                $PSSessionParams = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }

                # Add optional parameters to hash
                If ($Credential) {
                    $PSSessionParams.Add('Credential', $Credential)
                }

                # Create remote powershell session   
                Try {
                    $RemoteSession = New-PSSession @PSSessionParams
                }
                Catch {
                    Throw $_.Exception.Message
                }
            } Else { 
                $RemoteSession = $null
            }
        }
        
        # Build Hash to be used for passing parameters to 
        # Invoke-Command commandlet
        $CommandParams = @{
            ScriptBlock = $ScriptBlock
            ErrorAction = 'Stop'
        }
        
        # Add optional parameters to hash
        If ($RemoteSession) {
            $CommandParams.Add('Session', $RemoteSession)
        }
               
        # Run ScriptBlock    
        Try {
            $ReturnedValues = Invoke-Command @CommandParams
        }
        Catch {
            If ($RemoteSession) {
            	Remove-PSSession $RemoteSession
            }
            Throw $_.Exception.Message
        }

        # Assign variables
        Import-Module MS-Module
        $Date = Get-Date
        $OS_Name = $ReturnedValues.Operating_System.Caption + ' [Installed: ' + ([datetime]$ReturnedValues.Operating_System.InstallDate).ToString('dd-MMM-yyyy') + ']'
        $Computer_Name = $ReturnedValues.Operating_System.CSName
		If ($DomainName) {$Computer_Name = $Computer_Name + $DomainName.ToUpper()}
        $Kernel_Info = $ReturnedValues.Operating_System.Version + ' [' + $ReturnedValues.Operating_System.OSArchitecture + ']'
        $Process_Count = $ReturnedValues.Process_Count
        $Uptime = "$(($Uptime = $Date - $($ReturnedValues.Operating_System.LastBootUpTime)).Days) days, $($Uptime.Hours) hours, $($Uptime.Minutes) minutes"
        $Shell_Info = $ReturnedValues.Shell_Info
        $CPU_Info = $ReturnedValues.Processor.Name -replace '\(C\)', '' -replace '\(R\)', '' -replace '\(TM\)', '' -replace 'CPU', '' -replace '\s+', ' '
        $Current_Load = $ReturnedValues.Processor.LoadPercentage    
        $Memory_Size = "{0} MB/{1} MB " -f (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-
        ([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))),([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))
		$Disk_Size = "{0} GB/{1} GB" -f (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-
        [math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))),([math]::round($ReturnedValues.Logical_Disk.Size/1GB))

        # Write to the Console
        Write-Host -Object ("")
        Write-Host -Object ("")
        Write-Host -Object ("         ,.=:^!^!t3Z3z.,                  ") -ForegroundColor Red
        Write-Host -Object ("        :tt:::tt333EE3                    ") -ForegroundColor Red
        Write-Host -Object ("        Et:::ztt33EEE ") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @Ee.,      ..,     $($Date.ToString('dd-MMM-yyyy HH:mm:ss'))") -ForegroundColor Green
        Write-Host -Object ("       ;tt:::tt333EE7") -NoNewline -ForegroundColor Red
        Write-Host -Object (" ;EEEEEEttttt33#     ") -ForegroundColor Green
        Write-Host -Object ("      :Et:::zt333EEQ.") -NoNewline -ForegroundColor Red
        Write-Host -Object (" SEEEEEttttt33QL     ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("User: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$env:USERDOMAIN\$env:UserName") -ForegroundColor Cyan
        Write-Host -Object ("      it::::tt333EEF") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEttttt33F      ") -NoNewline -ForeGroundColor Green
        Write-Host -Object ("Hostname: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Computer_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ;3=*^``````'*4EEV") -NoNewline -ForegroundColor Red
        Write-Host -Object (" :EEEEEEttttt33@.      ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("OS: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$OS_Name") -ForegroundColor Cyan
        Write-Host -Object ("     ,.=::::it=., ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("``") -NoNewline -ForegroundColor Red
        Write-Host -Object (" @EEEEEEtttz33QF       ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Kernel: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("NT ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("$Kernel_Info") -ForegroundColor Cyan
        Write-Host -Object ("    ;::::::::zt33) ") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("  '4EEEtttji3P*        ") -NoNewline -ForegroundColor Green
        Write-Host -Object ("Uptime: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Uptime") -ForegroundColor Cyan
        Write-Host -Object ("   :t::::::::tt33.") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (":Z3z.. ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object (" ````") -NoNewline -ForegroundColor Green
        Write-Host -Object (" ,..g.        ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Shell: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("PowerShell $Shell_Info") -ForegroundColor Cyan
        Write-Host -Object ("   i::::::::zt33F") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" AEEEtttt::::ztF         ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("CPU: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$CPU_Info") -ForegroundColor Cyan
        Write-Host -Object ("  ;:::::::::t33V") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEttttt::::t3          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Processes: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Process_Count") -ForegroundColor Cyan
        Write-Host -Object ("  E::::::::zt33L") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" @EEEtttt::::z3F          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Current Load: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Current_Load") -NoNewline -ForegroundColor Cyan
        Write-Host -Object ("%") -ForegroundColor Cyan
        Write-Host -Object (" {3=*^``````'*4E3)") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" ;EEEtttt:::::tZ``          ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("Memory: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Memory_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB))-([math]::round($ReturnedValues.Operating_System.FreePhysicalMemory/1KB))) -MaxValue ([math]::round($ReturnedValues.Operating_System.TotalVisibleMemorySize/1KB)); "`r"
        Write-Host -Object ("             ``") -NoNewline -ForegroundColor Cyan
        Write-Host -Object (" :EEEEtttt::::z7            ") -NoNewline -ForegroundColor Yellow
        Write-Host -Object ("System Volume: ") -NoNewline -ForegroundColor Red
        Write-Host -Object ("$Disk_Size`t") -ForegroundColor Cyan -NoNewline
		New-PercentageBar -DrawBar -Value (([math]::round($ReturnedValues.Logical_Disk.Size/1GB)-[math]::round($ReturnedValues.Logical_Disk.FreeSpace/1GB))) -MaxValue ([math]::round($ReturnedValues.Logical_Disk.Size/1GB)); "`r"
        Write-Host -Object ("                 'VEzjt:;;z>*``           ") -ForegroundColor Yellow
        Write-Host -Object ("                      ````                  ") -ForegroundColor Yellow
        Write-Host -Object ("")
  } #End Process

  End {
        If ($RemoteSession) {
            Remove-PSSession $RemoteSession
        }
  }
} #End Function Get-MOTD

## Change Attributes
function Get-FileAttribute{
    param($file,$attribute)
    $val = [System.IO.FileAttributes]$attribute;
    if((gci $file -force).Attributes -band $val -eq $val){$true;} else { $false; }
} 


function Set-FileAttribute{
    param($file,$attribute)
    $file =(gci $file -force);
    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]$attribute).value__;
    if($?){$true;} else {$false;}
} 

## End Change Attributes

## Remote Group Policy
function GPR {
<# 
.SYNOPSIS 
    Open Group Policy for specified workstation(s) 

.EXAMPLE 
    GPR Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach ($Computer in $ComputerName) {

        Write-Progress -Activity "Opening Remote Group Policy..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

        #Opens (Remote) Group Policy for specified workstation
        GPedit.msc /gpcomputer: $Computer
    }
}#End GPR

## Begin Lastboot

function LastBoot {
<# 
.SYNOPSIS 
    Retrieve last restart time for specified workstation(s) 

.EXAMPLE 
    LastBoot Computer123456 

.EXAMPLE 
    LastBoot 123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [String[]]$ComputerName,

    $i=0,
    $j=0
)

    foreach($Computer in $ComputerName) {

        Write-Progress -Activity "Retrieving Last Reboot Time..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)
　
        $computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer

        [pscustomobject] @{

            ComputerName = $Computer
            LastReboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        }
    }
}#End LastBoot

#Begin SYSinfo
function SYSinfo {
<# 
.SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

.EXAMPLE 
  SYS Computer123456 
#> 

param(

    [Parameter(Mandatory=$true)]
    [string[]] $ComputerName,
    
    $i=0,
    $j=0
)

$Stamp = (Get-Date -Format G) + ":"

    function Systeminformation {
	
        foreach ($Computer in $ComputerName) {

            if(!([String]::IsNullOrWhiteSpace($Computer))) {

                if(Test-Connection -Quiet -Count 1 -Computer $Computer) {

                    Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	                Start-Job -ScriptBlock { param($Computer) 

	                    #Gather specified workstation information; CimInstance only works on 64-bit
	                    $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
	                    $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
	                    $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
	                    $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
	                    $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [PSCustomObject] @{

                            ComputerName = $computerSystem.Name
                            LastReboot = $computerOS.LastBootUpTime
                            OperatingSystem = $computerOS.OSArchitecture + " " + $computerOS.caption
                            Model = $computerSystem.Model
                            RAM = "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory/1GB) + "GB"
                            DiskCapacity = "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
                            TotalDiskSpace = "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
                            CurrentUser = $computerSystem.UserName
                        }
                    } -ArgumentList $Computer
                }

                else {

                    Start-Job -ScriptBlock { param($Computer)  
                     
                        [PSCustomObject] @{

                            ComputerName=$Computer
                            LastReboot="Unable to PING."
                            OperatingSystem="$Null"
                            Model="$Null"
                            RAM="$Null"
                            DiskCapacity="$Null"
                            TotalDiskSpace="$Null"
                            CurrentUser="$Null"
                        }
                    } -ArgumentList $Computer                       
                }
            }

            else {
                 
                Start-Job -ScriptBlock { param($Computer)  
                     
                    [PSCustomObject] @{

                        ComputerName = "Value is null."
                        LastReboot = "$Null"
                        OperatingSystem = "$Null"
                        Model = "$Null"
                        RAM = "$Null"
                        DiskCapacity = "$Null"
                        TotalDiskSpace = "$Null"
                        CurrentUser = "$Null"
                    }
                } -ArgumentList $Computer
            }
        } 
    }

    $SystemInformation = SystemInformation | Receive-Job -Wait | Select ComputerName, CurrentUser, OperatingSystem, Model, RAM, DiskCapacity, TotalDiskSpace, LastReboot
    $DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

	Switch($CheckBox.IsChecked) {

		$true { 
            
            $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force 
        }

		default { 
            
            $SystemInformation | Out-GridView -Title "System Information"
        }
    }

	if($CheckBox.IsChecked -eq $true) {

	    Try { 

		    $listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {

		    #Do Nothing 
	    }
	}
	
	else {

	    Try {

	        $listBox.Items.Add("$stamp System Information output processed!`n")
	    } 

	    Catch {

	        #Do Nothing 
	    }
	}
}#End SYSinfo

#Begin NetMessage

function NetMSG {
<# 
.SYNOPSIS 
    Generate a pop-up window on specified workstation(s) with desired message 

.EXAMPLE 
    NetMSG Computer123456 
#> 
	
param(

    [Parameter(Mandatory=$true)]
    [String[]] $ComputerName,

    [Parameter(Mandatory=$true,HelpMessage='Enter desired message')]
    [String]$MyMessage,

    [String]$User = [Environment]::UserName,

    [String]$UserJob = (Get-ADUser $User -Property Title).Title,
    
    [String]$CallBack = "$User | 5-2444 | $UserJob",

    $i=0,
    $j=0
)

    function SendMessage {

        foreach($Computer in $ComputerName) {

            Write-Progress -Activity "Sending messages..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)         

            #Invoke local MSG command on specified workstation - will generate pop-up message for any user logged onto that workstation - *Also shows on Login screen, stays there for 100,000 seconds or until interacted with
            Invoke-Command -ComputerName $Computer { param($MyMessage, $CallBack, $User, $UserJob)
 
                MSG /time:100000 * /v "$MyMessage {$CallBack}"
            } -ArgumentList $MyMessage, $CallBack, $User, $UserJob -AsJob
        }
    }

    SendMessage | Wait-Job | Remove-Job

}#End NetMSG

function InstallApplication {

<#     
.SYNOPSIS     
  
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP 

.DESCRIPTION     
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP

.EXAMPLE    
    .\InstallAsJob (Get-Content C:\ComputerList.txt)

.EXAMPLE    
    .\InstallAsJob Computer1, Computer2, Computer3 
    
.NOTES   
    Author: JBear 
    Date: 2/9/2017 
    
    Edit: JBear
    Date: 10/13/2017 
#> 

param(

    [Parameter(Mandatory=$true,HelpMessage="Enter Computername(s)")]
    [String[]]$Computername,

    [Parameter(ValueFromPipeline=$true,HelpMessage="Enter installer path(s)")]
    [String[]]$Path = $null,

    [Parameter(ValueFromPipeline=$true,HelpMessage='Enter remote destination: C$\Directory')]
    $Destination = "C$\TempApplications"
)

    if($Path -eq $null) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\lasfs03\Software\Current Version\Deploy"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect=$true
        $Result = $Dialog.ShowDialog()

        if($Result -eq 'OK') {

            Try {
        
                $Path = $Dialog.FileNames
            }

            Catch {

                $Path = $null
	            Break
            }
        }

        else {

            #Shows upon cancellation of Save Menu
            Write-Host -ForegroundColor Yellow "Notice: No file(s) selected."
            Break
        }
    }

    #Create function    
    function InstallAsJob {

        #Each item in $Computernam variable        
        foreach($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if(!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if(Test-Connection -Quiet -Count 1 $Computer) {                                               
                     
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach($P in $Path) {
                            
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if(!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                     
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if($FileName -like "*.exe") {

                                function InvokeEXE {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
                                    
                                        Try {

                                            #Start EXE file
                                            Start-Process $Executable -ArgumentList "/s" -Wait -NoNewWindow
                                            
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }
                                       
                                    } -AsJob -JobName "Silent EXE Install" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeEXE | Receive-Job -Wait
                            }
                               
                            #Install .MSI                                        
                            elseif($FileName -like "*.msi") {

                                function InvokeMSI {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSIArguments = @(
						
						                    "/i"
						                    $Executable
						                    "/qn"
					                    )

                                        Try {
                                        
                                            #Start MSI file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSIArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                              
                                    } -AsJob -JobName "Silent MSI Install" -ArgumentList $TempDir, $FileName, $Executable                            
                                }

                                InvokeMSI | Receive-Job -Wait
                            }

                            #Install .MSP
                            elseif($FileName -like "*.msp") { 
                                                                       
                                function InvokeMSP {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
				    
				    	                $MSPArguments = @(
						
						                    "/p"
						                    $Executable
						                    "/qn"
					                    )				    

                                        Try {
                                                                                
                                            #Start MSP file                                    
                                            Start-Process 'msiexec.exe' -ArgumentList $MSPArguments -Wait -ErrorAction Stop

                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                    
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force

                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                             
                                    } -AsJob -JobName "Silent MSP Installer" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeMSP | Receive-Job -Wait
                            }

                            else {

                                Write-Host "$Destination has an unsupported file extension. Please try again."                        
                            }
                        }                      
                    } -Name "Application Install" -Argumentlist $Computer, $Path, $Destination            
                }
                                            
                else {                                
                    
                    Write-Host "Unable to connect to $Computer."                
                }            
            }        
        }   
    }

    #Call main function
    InstallAsJob
    Write-Host "`nJob creation complete. Please use the Get-Job cmdlet to check progress.`n"
    Write-Host "Once all jobs are complete, use Get-Job | Receive-Job to retrieve any output or, Get-Job | Remove-Job to clear jobs from the session cache."
}#End InstallApplication

# Begin Get-Icon

Function Get-Icon {
    <#
        .SYNOPSIS
            Gets the icon from a file

        .DESCRIPTION
            Gets the icon from a file and displays it in a variety formats.

        .PARAMETER Path
            The path to a file to get the icon

        .PARAMETER ToBytes
            Displays outputs as a byte array

        .PARAMETER ToBitmap
            Display the icon as a bitmap object

        .PARAMETER ToBase64
            Displays the icon in Base64 encoded format

        .NOTES
            Name: Get-Icon
            Author: Boe Prox
            Version History:
                1.0 //Boe Prox - 11JAN2016
                    - Initial version

        .OUTPUT
            System.Drawing.Icon
            System.Drawing.Bitmap
            System.String
            System.Byte[]

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe'

            FullName : C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe
            Handle   : 164169893
            Height   : 32
            Size     : {Width=32, Height=32}
            Width    : 32

            Description
            -----------
            Returns the System.Drawing.Icon representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap

            Tag                  : 
            PhysicalDimension    : {Width=32, Height=32}
            Size                 : {Width=32, Height=32}
            Width                : 32
            Height               : 32
            HorizontalResolution : 96
            VerticalResolution   : 96
            Flags                : 2
            RawFormat            : [ImageFormat: b96b3caa-0728-11d3-9d7b-0000f81ef32e]
            PixelFormat          : Format32bppArgb
            Palette              : System.Drawing.Imaging.ColorPalette
            FrameDimensionsList  : {7462dc86-6180-4c7e-8e3f-ee7333a7a483}
            PropertyIdList       : {}
            PropertyItems        : {}

            Description
            -----------
            Returns the System.Drawing.Bitmap representation of the icon

        .EXAMPLE
            $FileName = 'C:\Temp\PowerShellIcon.png'
            $Format = [System.Drawing.Imaging.ImageFormat]::Png
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBitmap).Save($FileName,$Format)

            Description
            -----------
            Saves the icon as a file.

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64

            AAABAAEAICAQHQAAAADoAgAAFgAAACgAAAAgAAAAQAAAAAEABAAAAAAAgAIAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAMDAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP
            //AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmZmZmZmZmZmZmZgAAAAAAaId3d3d3d4iIiIdgAA
            AHdmhmZmZmZmZmZmZoZAAAB2ZnZmZmZmZmZmZmZ3YAAAdmZ3ZmiHZniIiHZmaGAAAHZmd2Zv/4eIiIi
            GZmhgAAB2ZmdmZ4/4eIh3ZmZnYAAAd2ZnZmZo//h2ZmZmZ3YAAHZmaGZmZo//h2ZmZmd2AAB3Zmd2Zm
            Znj/h2ZmZmhgAAd3dndmZmZuj/+GZmZoYAAHd3dod3dmZuj/9mZmZ2AACHd3aHd3eIiP/4ZmZmd2AAi
            Hd2iIiIiI//iId2ZndgAIiIhoiIiIj//4iIiIiIYACIiId4iIiP//iIiIiIiGAAiIiIaIiI//+IiIiI
            iIhkAIiIiGiIiP/4iIiIiIiIdgCIiIhoiIj/iIiIiIiIiIYAiIiIeIiIiIiIiIiIiIiGAAiIiIaP///
            ////////4hgAAAAAGZmZmZmZmZmZmZmYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
            AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD////////////////gA
            AAf4AAAD+AAAAfgAAAHAAAABwAAAAcAAAAHAAAAAwAAAAMAAAADAAAAAwAAAAMAAAABAAAAAQAAAAEA
            AAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAP4AAAH//////////////////////////w==

            Description
            -----------
            Returns the Base64 encoded representation of the icon

        .EXAMPLE
            Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBase64 | Clip

            Description
            -----------
            Returns the Base64 encoded representation of the icon and saves it to the clipboard.

        .EXAMPLE
            (Get-Icon -Path 'C:\windows\system32\WindowsPowerShell\v1.0\PowerShell.exe' -ToBytes) -Join ''

            0010103232162900002322002200040000320006400010400000128200000000000000000000000
            0128001280001281280128000128012801281280012812812801921921920002550025500025525
            5025500025502550255255002552552550000000000000000000000000000000000000000000000
            0000000000000000000000000000000000006102102102102102102102102102102960000613611
            9119119119119120136136136118000119102134102102102102102102102102102134640011810
            2118102102102102102102102102102119960011810211910210413510212013613611810210496
            0011810211910211125513513613613613410210496001181021031021031432481201361191021
            0210396001191021031021021042552481181021021021031180011810210410210210214325513
            5102102102103118001191021031181021021031432481181021021021340011911910311810210
            2102232255248102102102134001191191181351191181021101432551021021021180013511911
            8135119119136136255248102102102119960136119118136136136136143255136135118102119
            9601361361341361361361362552551361361361361369601361361351201361361432552481361
            3613613613696013613613610413613625525513613613613613613610001361361361041361362
            5524813613613613613613611801361361361041361362551361361361361361361361340136136
            1361201361361361361361361361361361361340813613613414325525525525525525525525524
            8134000061021021021021021021021021021021020000000000000000000000000000000000000
            0000000000000000000000000000000000000000000025525525525525525525525525525525525
            5224003122400152240072240070007000700070003000300030003000300010001000100010000
            0000000000000000000012800025400125525525525525525525525525525525525525525525525
            5255255255255

            Description
            -----------
            Returns the bytes representation of the icon. -Join was used in this for the sake
            of displaying all of the data.

    #>
    [cmdletbinding(
        DefaultParameterSetName = '__DefaultParameterSetName'
    )]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullorEmpty()]
        [string]$Path,
        [parameter(ParameterSetName = 'Bytes')]
        [switch]$ToBytes,
        [parameter(ParameterSetName = 'Bitmap')]
        [switch]$ToBitmap,
        [parameter(ParameterSetName = 'Base64')]
        [switch]$ToBase64
    )
    Begin {
        If ($PSBoundParameters.ContainsKey('Debug')) {
            $DebugPreference = 'Continue'
        }
        Add-Type -AssemblyName System.Drawing
    }
    Process {
        $Path = Convert-Path -Path $Path
        Write-Debug $Path
        If (Test-Path -Path $Path) {
            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)| 
            Add-Member -MemberType NoteProperty -Name FullName -Value $Path -PassThru
            If ($PSBoundParameters.ContainsKey('ToBytes')) {
                Write-Verbose "Retrieving bytes"
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $MemoryStream.ToArray()   
                $MemoryStream.Flush()  
                $MemoryStream.Dispose()           
            } ElseIf ($PSBoundParameters.ContainsKey('ToBitmap')) {
                $Icon.ToBitMap()
            } ElseIf ($PSBoundParameters.ContainsKey('ToBase64')) {
                $MemoryStream = New-Object System.IO.MemoryStream
                $Icon.save($MemoryStream)
                Write-Debug ($MemoryStream | Out-String)
                $Bytes = $MemoryStream.ToArray()   
                $MemoryStream.Flush() 
                $MemoryStream.Dispose()
                [convert]::ToBase64String($Bytes)
            }  Else {
                $Icon
            }
        } Else {
            Write-Warning "$Path does not exist!"
            Continue
        }
    }
}

# End Get-Icon

# Get Mapped Drive
function Get-MappedDrive {
	param (
	    [string]$computername = "localhost"
	)
	    Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername | 
	    Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
	}
# End Get Mapped Drive

# LayZ LazyWinAdmin GUI Tool
function LayZ {
    C:\LazyWinAdmin\LazyWinAdmin\LazyWinAdmin.ps1
    }
# End LayZ LazyWinAdmin GUI Tool

# User Last Login
Function Get-UserLastLogonTime{

<#
.SYNOPSIS
Gets the last logon time of users on a Computer.

.DESCRIPTION
Pulls information from the wmi object Win32_UserProfile and outputs an array of objects with properties Name and LastUseTime.
If a date that is year 1 is outputted, then an error occured.

.PARAMETER ComputerName
[object] Specify which computer to target when finding logged on Users.
Default is the host computer

.PARAMETER User
[string] Specify a user to find on the computer.

.PARAMETER ListAllUsers
[switch] Specify the function to list all users that logged into the computer.

.PARAMETER GetLastUsers
[switch] Specify the function to get the last user to log onto the computer.

.PARAMETER ListCommonUsers
[switch] Specify to the function to list common user.

.INPUTS
You may pipe objects into the ComputerName parameter.

.OUTPUTS
outputs an object array with a size dependant on the number of users that logged in with propeties Name and LastUseTime.


#>

    [cmdletBinding()]
    param(
        #computer Name
        [parameter(Mandatory = $False, ValueFromPipeline = $True)]
        [object]$ComputerName = $env:COMPUTERNAME,

        #parameter set, can only choose one from this group
        [parameter(Mandatory = $False, parameterSetName = 'user')]
        [string] $User,
        [parameter(ParameterSetName = 'all users')]
        [switch] $ListAllUsers,
        [parameter(ParameterSetName = 'Last user')]
        [switch] $GetLastUser,

        #Whether or not you want the function to list Common users
        [switch] $ListCommonUsers
    )

    #Begin Pipeline
    Begin{
        #List of users that are present on all PCs, these won't output unless specified to do so.
        $CommonUsers = ("NetworkService", "LocalService", "systemprofile")
    }
    
    #Process Pipeline
    Process{
        #ping the machine before trying to do anything
        if(Test-Connection $ComputerName -Count 2 -Quiet){
            #try to get the OS version of the computer
            try{$OS = (gwmi  -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction stop).caption}
            catch{
                #had an error getting the WMI-object
                return New-Object psobject -Property @{
                            User = "Error getting WMIObject Win32_OperatingSystem"
                            LastUseTime = get-date 0
                            }
              }
            #make sure the OS retrieved is either Windows 7 or Windows 10 as this function has not been set to work on other operating systems
            if($OS.contains("Windows 10") -or $OS.Contains("Windows 7")){
                try{
                    #try to get the WMiObject win32_UserProfile
                    $UserObjects = Get-WmiObject -ComputerName $ComputerName -Class Win32_userProfile -Property LocalPath,LastUseTime -ErrorAction Stop
                    $users = @()
                    #loop that handles all the data that came back from the WMI objects
                    forEach($UserObject in $UserObjects){
                        #extract the username from the local path, first find where the last slash is after, get everything after that into its own string
                        $i = $UserObject.localPath.LastIndexOf("\") + 1
                        $tempUserString = ""
                        while($UserObject.localPath.toCharArray()[$i] -ne $null){
                            $tempUserString += $UserObject.localPath.toCharArray()[$i]
                            $i++
                        }
                        #if list common users is turned on, skip this block
                        if(!$listCommonUsers){
                            #check if the user extracted from the local path is a common user
                            $isCommonUser = $false
                            forEach($userName in $CommonUsers){ 
                                if($userName -eq $tempUserString){
                                    $isCommonUser = $true
                                    break
                                }
                            }
                        }
                        #if the user is one of the users specified in the common users, skip it unless otherwise specified
                        if($isCommonUser){continue}
                        #check to see if the user has a timestamp for there last logon 
                        if($UserObject.LastUseTime -ne $null){
                            #This converts the string timestamp to a DateTime object
                            $TempUserLastUseTime = ([WMI] '').ConvertToDateTime($userObject.LastUseTime)
                        }
                        #the user had a local path but no timestamp was found for last logon
                        else{$TempUserLastUseTime = Get-Date 0}
                        #add user to array
                        $users += New-Object -TypeName psobject -Property @{
                            User = $tempUserString
                            LastUseTime = $TempUserLastUseTime
                            }
                    }
                }
                catch{
                    #error trying to retrieve WMI obeject win32_userProfile
                    return New-Object psobject -Property @{
                        User = "Error getting WMIObject Win32_userProfile"
                        LastUseTime = get-date 0
                        }
                }
            }
            else{
                #OS version was not compatible
                return New-Object psobject -Property @{
                    User = "Operating system $OS is not compatible with this function."
                    LastUseTime = get-date 0
                    }
            }
        }
        else{
            #Computer was not pingable
            return New-Object psobject -Property @{
                User = "Can't Ping"
                LastUseTime = get-date 0
                }
        }

        #check to see if any users came out of the main function
        if($users.count -eq 0){
            $users += New-Object -TypeName psobject -Property @{
                User = "NoLoggedOnUsers"
                LastUseTime = get-date 0
            }
        }
        #sort the user array by the last time they logged in
        else{$users = $users | Sort-Object -Property LastUseTime -Descending}
        #main output block
        #if List all users was chosen, output the full list of users found
        if($ListAllUsers){return $users}
        #if get last user was chosen, output the last user to log on the computer
        elseif($GetLastUser){return ($users[0])}
        else{
            #see if the user specified ever logged on
            ForEach($Username in $users){
                if($Username.User -eq $user) {return ($Username)}            
            }

            #user did not log on
            return New-Object psobject -Property @{
                User = "$user"
                LastUseTime = get-date 0
                }
        }
    }
    #End Pipeline
    End{Write-Verbose "Function get-UserLastLogonTime is complete"}
}
# End User Last Login

# Begin Unblock
Function Unblock ($path) { 

Get-ChildItem "$path" -Recurse | Unblock-File

}
# End Unblock 
Function Add-Help
{
 $helpText = @"
<#
.SYNOPSIS
    What does this do? 

.PARAMETER param1 
    What is param1?

.PARAMETER param2 
    What is param2?

.NOTES
    NAME: $($psISE.CurrentFile.DisplayName)
    AUTHOR: $env:username
    LASTEDIT: $(Get-Date)
    KEYWORDS:

.LINK
    http://julianscorner.com

.EXAMPLE     
    '12345' | THIS-FUNCTION -param1 180   
    Describe what this example accomplishes 
       
.EXAMPLE     
    THIS-FUNCTION -param2 @("text1","text2") -param1 180   
    Describe what this example accomplishes 

#Requires -Version 2.0
#>

"@
 $psise.CurrentFile.Editor.InsertText($helpText)
}

Function Add-FunctionTemplate
{
  $text1 = @"
Function THIS-FUNCTION
{
"@
  $text2 = @"
	[CmdletBinding()]
  param
  (
    [Parameter(Mandatory = `$true,
               ValueFromPipeline = `$true)]   
    [array]`$param1,                                   
    [Parameter(Mandatory = `$true)]   
    [int]`$param2
  )   
  BEGIN
  {
    # This block is used to provide optional one-time pre-processing for the function.
    # PowerShell uses the code in this block one time for each instance of the function in the pipeline.
  }
  PROCESS
  {
    # This block is used to provide record-by-record processing for the function.
    # This block might be used any number of times, or not at all, depending on the input to the function.
    # For example, if the function is the first command in the pipeline, the Process block will be used one time.
    # If the function is not the first command in the pipeline, the Process block is used one time for every
    # input that the function receives from the pipeline.
    # If there is no pipeline input, the Process block is not used.
  }
  END
  {
    # This block is used to provide optional one-time post-processing for the function.
  } 
}
"@
 $psise.CurrentFile.Editor.InsertText($text1)
 Add-Help
 $psise.CurrentFile.Editor.InsertText($text2)
}

Function Remove-AliasFromScript
{
  Get-Alias | 
    Select-Object Name, Definition | 
    ForEach-Object -Begin { $a = @{} } -Process {$a.Add($_.Name, $_.Definition)} -End {}

  $b = $errors = $null
  $b = $psISE.CurrentFile.Editor.Text

  [system.management.automation.psparser]::Tokenize($b,[ref]$errors) |
    Where-Object { $_.Type -eq "command" } |
      ForEach-Object `
      {
        if ($a.($_.Content))
        {
          $b = $b -replace
            ('(?<=(\W|\b|^))' + [regex]::Escape($_.content) + '(?=(\W|\b|$))'),
              $a.($_.content)
        }
      }

  $ScriptWithoutAliases = $psISE.CurrentPowerShellTab.Files.Add()
  $ScriptWithoutAliases.Editor.Text = $b
  $ScriptWithoutAliases.Editor.SetCaretPosition(1,1)
  $ScriptWithoutAliases.Editor.EnsureVisible(1)  
}

Function Replace-SpacesWithTabs
{
  param
  (
    [int]$spaces = 2
  ) 
  
  $tab = "`t"
  $space = " " * $spaces
  $text = $psISE.CurrentFile.Editor.Text

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    if ($line -match "\S")
    {
      $pos = $line.IndexOf($Matches[0])
      $indentation = $line.SubString(0, $pos)
      $remainder = $line.SubString($pos)
      
      $replaced = $indentation -replace $space, $tab
      
      $newText += $replaced + $remainder + [Environment]::NewLine
    }
    else
    {
      $newText += $line + [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.Text  = $newText
  }
}

Function Replace-TabsWithSpaces
{
  param
  (
    [int]$spaces = 2
  )   
  
  $tab = "`t"
  $space = " " * $spaces
  $text = $psISE.CurrentFile.Editor.Text

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    if ($line -match "\S")
    {
      $pos = $line.IndexOf($Matches[0])
      $indentation = $line.SubString(0, $pos)
      $remainder = $line.SubString($pos)
      
      $replaced = $indentation -replace $tab, $space
      
      $newText += $replaced + $remainder + [Environment]::NewLine
    }
    else
    {
      $newText += $line + [Environment]::NewLine
    }

    $psISE.CurrentFile.Editor.Text  = $newText
  }
}

Function Indent-SelectedText
{
  param
  (
    [int]$spaces = 2
  )
  
  $tab = " " * $space
  $text = $psISE.CurrentFile.Editor.SelectedText

  $newText = ""
  
  foreach ($line in $text -split [Environment]::NewLine)
  {
    $newText += $tab + $line + [Environment]::NewLine
  }

   $psISE.CurrentFile.Editor.InsertText($newText)
}

Function Add-RemarkedText
{
<#
.SYNOPSIS
    This function will add a remark character # to selected text in the ISE.
    These are comment characters, and is great when you want to comment out
    a section of PowerShell code.

.NOTES
    NAME:  Add-RemarkedText
    AUTHOR: ed wilson, msft
    LASTEDIT: 05/16/2013
    KEYWORDS: Windows PowerShell ISE, Scripting Techniques

.LINK
     http://www.ScriptingGuys.com

#Requires -Version 2.0
#>
  $text = $psISE.CurrentFile.Editor.SelectedText

  foreach ($l in $text -Split [Environment]::NewLine)
  {
   $newText += "{0}{1}" -f ("#" + $l),[Environment]::NewLine
  }

  $psISE.CurrentFile.Editor.InsertText($newText)
}

Function Remove-RemarkedText
{
<#
.SYNOPSIS
    This function will remove a remark character # to selected text in the ISE.
    These are comment characters, and is great when you want to clean up a
    previously commentted out section of PowerShell code.

.NOTES
    NAME:  Add-RemarkedText
    AUTHOR: ed wilson, msft
    LASTEDIT: 05/16/2013
    KEYWORDS: Windows PowerShell ISE, Scripting Techniques

.LINK
     http://www.ScriptingGuys.com

#Requires -Version 2.0
#>
  $text = $psISE.CurrentFile.Editor.SelectedText

  foreach ($l in $text -Split [Environment]::NewLine)
  {
    $newText += "{0}{1}" -f ($l -Replace '#',''),[Environment]::NewLine
  }

  $psISE.CurrentFile.Editor.InsertText($newText)
}
Function AbortScript
{
	$Word.Quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject( $Word ) | Out-Null
	If( Get-Variable -Name Word -Scope Global )
	{
		Remove-Variable -Name word -Scope Global
	}
	[GC]::Collect() 
	[GC]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
Function Add-ADSubnet{
<#
	.SYNOPSIS
		This function allow you to add a subnet object in your active directory using ADSI

	.DESCRIPTION
		This function allow you to add a subnet object in your active directory using ADSI
	
	.PARAMETER  Subnet
		Specifies the Name of the subnet to add

	.PARAMETER  SiteName
		Specifies the Name of the Site where the subnet will be created
	
	.PARAMETER  Description
		Specifies the Description of the subnet

	.PARAMETER  Location
		Specifies the Location of the subnet

	.EXAMPLE
		Add-ADSubnet -Subnet "192.168.10.0/24" -SiteName MTL1
	
	This will create the subnet "192.168.10.0/24" and assign it to the site "MTL1".

	.EXAMPLE
		Add-ADSubnet -Subnet "192.168.10.0/24" -SiteName MTL1 -Description "Workstations VLAN 110" -Location "Montreal, Canada" -verbose
	
	This will create the subnet "192.168.10.0/24" and assign it to the site "MTL1" with the description "Workstations VLAN 110" and the location "Montreal, Canada"
	Using the parameter -Verbose, the script will show the progression of the subnet creation.
	

	.NOTES
		NAME:	FUNCT-AD-SITE-Add-ADSubnet_using_ADSI.ps1
		AUTHOR:	Francois-Xavier CAT 
		DATE:	2013/11/07
		EMAIL:	info@lazywinadmin.com
		WWW:	www.lazywinadmin.com
		TWITTER:@lazywinadm
	
		http://www.lazywinadmin.com/2013/11/powershell-add-ad-site-subnet.html

		VERSION HISTORY:
		1.0 2013.11.07
			Initial Version

#>
	[CmdletBinding()]
	PARAM(
		[Parameter(
			Mandatory=$true,
			Position=1,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Subnet name to create")]
		[Alias("Name")]
		[String]$Subnet,
		[Parameter(
			Mandatory=$true,
			Position=2,
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Site to which the subnet will be applied")]
		[Alias("Site")]
		[String]$SiteName,
		[Parameter(
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Description of the Subnet")]
		[String]$Description,
		[Parameter(
			ValueFromPipelineByPropertyName=$true,
			HelpMessage="Location of the Subnet")]
		[String]$location
	)
	PROCESS{
			TRY{
				$ErrorActionPreference = 'Stop'
				
				# Distinguished Name of the Configuration Partition
				$Configuration = ([ADSI]"LDAP://RootDSE").configurationNamingContext

				# Get the Subnet Container
				$SubnetsContainer = [ADSI]"LDAP://CN=Subnets,CN=Sites,$Configuration"
				
				# Create the Subnet object
				Write-Verbose -Message "$subnet - Creating the subnet object..."
				$SubnetObject = $SubnetsContainer.Create('subnet', "cn=$Subnet")
			
				# Assign the subnet to a site
				$SubnetObject.put("siteObject","cn=$SiteName,CN=Sites,$Configuration")
	
				# Adding the Description information if specified by the user
				IF ($PSBoundParameters['Description']){
					$SubnetObject.Put("description",$Description)
				}
				
				# Adding the Location information if specified by the user
				IF ($PSBoundParameters['Location']){
					$SubnetObject.Put("location",$Location)
				}
				$SubnetObject.setinfo()
				Write-Verbose -Message "$subnet - Subnet added."
			}#TRY
			CATCH{
				Write-Warning -Message "An error happened while creating the subnet: $subnet"
				$error[0].Exception
			}#CATCH
	}#PROCESS Block
	END{
		Write-Verbose -Message "Script Completed"
	}#END Block
}#Function Add-ADSubnet

########################
# Office Word Functions#
########################
Function Add-OSCPicture
{
<#
.SYNOPSIS
Add-OSCPicture is an advanced function which can be used to insert many pictures into a word document.
.DESCRIPTION
Add-OSCPicture is an advanced function which can be used to insert many pictures into a word document.
.PARAMETER  <Path>
Specifies the path of slide.
.EXAMPLE
C:\PS> Add-OSCPicture -WordDocumentPath D:\Word\Document.docx -ImageFolderPath "C:\Users\Public\Pictures\Sample Pictures"
Action(Insert) ImageName
-------------- ---------
Finished   Chrysanthemum.jpg
Finished   Desert.jpg
Finished   Hydrangeas.jpg
Finished   Jellyfish.jpg
Finished   Koala.jpg
Finished   Lighthouse.jpg
Finished   Penguins.jpg
Finished   Tulips.jpg

This command shows how to insert many pictures to word document.
#>
[CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true,Position=0)]
    [Alias('wordpath')]
    [String]$WordDocumentPath,
    [Parameter(Mandatory=$true,Position=1)]
    [Alias('imgpath')]
    [String]$ImageFolderPath
    )

If(Test-Path -Path $WordDocumentPath)
{
    If(Test-Path -Path $ImageFolderPath)
    {
    $WordExtension = (Get-Item -Path $WordDocumentPath).Extension
    If($WordExtension -like ".doc" -or $WordExtension -like ".docx")
        {
    $ImageFiles = Get-ChildItem -Path $ImageFolderPath -Recurse -Include *.emf,*.wmf,*.jpg,*.jpeg,*.jfif,*.png,*.jpe,*.bmp,*.dib,*.rle,*.gif,*.emz,*.wmz,*.pcz,*.tif,*.tiff,*.eps,*.pct,*.pict,*.wpg

    If($ImageFiles)
    {
    #Create the Word application object
    $WordAPP = New-Object -ComObject Word.Application
    $WordDoc = $WordAPP.Documents.Open("$WordDocumentPath")

    Foreach($ImageFile in $ImageFiles)
    {
    $ImageFilePath = $ImageFile.FullName

    $Properties = @{'ImageName' = $ImageFile.Name
    'Action(Insert)' = Try
    {
    $WordAPP.Selection.EndKey(6)|Out-Null
    $WordApp.Selection.InlineShapes.AddPicture("$ImageFilePath")|Out-Null
    $WordApp.Selection.InsertNewPage() #insert new page to word
    "Finished"
    }
    Catch
    {
    "Unfinished"
    }
    }

    $objWord = New-Object -TypeName PSObject -Property $Properties
    $objWord
    }

    $WordDoc.Save()
    $WordDoc.Close()
    $WordAPP.Quit()#release the object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordAPP)|Out-Null
    Remove-Variable WordAPP
    }
    Else
    {
    Write-Warning "There is no image in this '$ImageFolderPath' folder."
    }
    }
    Else
    {
    Write-Warning "There is no word document file in this '$WordDocumentPath' folder."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$ImageFolderPath' because it does not exist."
    }
    }
    Else
    {
    Write-Warning "Cannot find path '$WordDocumentPath' because it does not exist."
    }
    }


Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines=$false,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = '-231'
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $null) -and ($Headers -ne $null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($Columns -ne $null) -and ($Headers -ne $null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $null) 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting)
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

########################
# AD Functions##########
###############

function Get-ADFSMORole
{
	<#
	.SYNOPSIS
		Retrieve the FSMO Role in the Forest/Domain.
	.DESCRIPTION
		Retrieve the FSMO Role in the Forest/Domain.
	.EXAMPLE
		Get-ADFSMORole
    .EXAMPLE
		Get-ADFSMORole -Credential (Get-Credential -Credential "CONTOSO\SuperAdmin")
    .NOTES
        Francois-Xavier Cat
        www.lazywinadmin.com
        @lazywinadm
		github.com/lazywinadmin
	#>
	[CmdletBinding()]
	PARAM (
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#PARAM
	BEGIN
	{
		TRY
		{
			# Load ActiveDirectory Module if not already loaded.
			IF (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction 'Stop' -Verbose:$false }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			Write-Warning -Message $Error[0]
		}
	}
	PROCESS
	{
		TRY
		{
            
			IF ($PSBoundParameters['Credential'])
			{
                # Query with the credentials specified
				$ForestRoles = Get-ADForest -Credential $Credential -ErrorAction 'Stop' -ErrorVariable ErrorGetADForest
				$DomainRoles = Get-ADDomain -Credential $Credential -ErrorAction 'Stop' -ErrorVariable ErrorGetADDomain
			}
			ELSE
			{
                # Query with the current credentials
				$ForestRoles = Get-ADForest
				$DomainRoles = Get-ADDomain
			}
			
            # Define Properties
			$Properties = @{
				SchemaMaster = $ForestRoles.SchemaMaster
				DomainNamingMaster = $ForestRoles.DomainNamingMaster
				InfraStructureMaster = $DomainRoles.InfraStructureMaster
				RIDMaster = $DomainRoles.RIDMaster
				PDCEmulator = $DomainRoles.PDCEmulator
			}
			
			New-Object -TypeName PSObject -Property $Properties
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Something wrong happened"
			IF ($ErrorGetADForest) { Write-Warning -Message "[PROCESS] Error While retrieving Forest information"}
			IF ($ErrorGetADDomain) { Write-Warning -Message "[PROCESS] Error While retrieving Domain information"}
			Write-Warning -Message $Error[0]
		}
	}#PROCESS
}

Function Get-AccountLockedOut
{
	
<#
.SYNOPSIS
	This function will find the device where the account get lockedout
.DESCRIPTION
	This function will find the device where the account get lockedout.
	It will query directly the PDC for this information
	
.PARAMETER DomainName
	Specifies the DomainName to query, by default it takes the current domain ($env:USERDOMAIN)
.PARAMETER UserName
	Specifies the DomainName to query, by default it takes the current domain ($env:USERDOMAIN)
.EXAMPLE
	Get-AccountLockedOut -UserName * -StartTime (Get-Date).AddDays(-5) -Credential (Get-Credential)
	
	This will retrieve the all the users lockedout in the last 5 days using the credential specify by the user.
	It might not retrieve the information very far in the past if the PDC logs are filling up very fast.
	
.EXAMPLE
	Get-AccountLockedOut -UserName "Francois-Xavier.cat" -StartTime (Get-Date).AddDays(-2)
#>
	
	#Requires -Version 3.0
	[CmdletBinding()]
	param (
		[string]$DomainName = $env:USERDOMAIN,
		[Parameter()]
		[ValidateNotNullorEmpty()]
		[string]$UserName = '*',
		[datetime]$StartTime = (Get-Date).AddDays(-1),
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	BEGIN
	{
		TRY
		{
            #Variables
            $TimeDifference = (Get-Date) - $StartTime

			Write-Verbose -Message "[BEGIN] Looking for PDC..."
			
			function Get-PDCServer
			{
	<#
	.SYNOPSIS
		Retrieve the Domain Controller with the PDC Role in the domain
	#>
				PARAM (
					$Domain = $env:USERDOMAIN,
					$Credential = [System.Management.Automation.PSCredential]::Empty
				)
				
				IF ($PSBoundParameters['Credential'])
				{
					
					[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(
					(New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext -ArgumentList 'Domain', $Domain, $($Credential.UserName), $($Credential.GetNetworkCredential().password))
					).PdcRoleOwner.name
				}#Credentials
				ELSE
				{
					[System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(
					(New-Object -TypeName System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain))
					).PdcRoleOwner.name
				}
			}#function Get-PDCServer
			
			Write-Verbose -Message "[BEGIN] PDC is $(Get-PDCServer)"
		}#TRY
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			Write-Warning -Message $Error[0]
		}
		
	}#BEGIN
	PROCESS
	{
		TRY
		{
			# Define the parameters
			$Splatting = @{ }
			
			# Add the credential to the splatting if specified
			IF ($PSBoundParameters['Credential'])
			{
                Write-Verbose -Message "[PROCESS] Credential Specified"
				$Splatting.Credential = $Credential
				$Splatting.ComputerName = $(Get-PDCServer -Domain $DomainName -Credential $Credential)
			}
			ELSE
			{
				$Splatting.ComputerName =$(Get-PDCServer -Domain $DomainName)
			}
			
			# Query the PDC
            Write-Verbose -Message "[PROCESS] Querying PDC for LockedOut Account in the last Days:$($TimeDifference.days) Hours: $($TimeDifference.Hours) Minutes: $($TimeDifference.Minutes) Seconds: $($TimeDifference.seconds)"
			Invoke-Command @Splatting -ScriptBlock {
				
				# Query Security Logs
				Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $Using:StartTime } |
				Where-Object { $_.Properties[0].Value -like "$Using:UserName" } |
				Select-Object -Property TimeCreated,
							  @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
							  @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
			} | Select-Object -Property TimeCreated, UserName, ClientName
		}#TRY
		CATCH
		{
				
		}
	}#PROCESS
}

function Get-ADGPOReplication
{
	<#
	.SYNOPSIS
		This function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.DESCRIPTION
		This function retrieve one or all the GPO and report their DSVersions and SysVolVersions (Users and Computers)
	.PARAMETER GPOName
		Specify the name of the GPO
	.PARAMETER All
		Specify that you want to retrieve all the GPO (slow if you have a lot of Domain Controllers)
	.EXAMPLE
		Get-ADGPOReplication -GPOName "Default Domain Policy"
	.EXAMPLE
		Get-ADGPOReplication -All
	.NOTES
		Francois-Xavier Cat
		@lazywinadm
		lazywinadmin.com
	
		VERSION HISTORY
		1.0 2014.09.22 	Initial version
						Adding some more Error Handling
						Fix some typo
	#>
	#requires -version 3
	[CmdletBinding()]
	PARAM (
		[parameter(Mandatory = $True, ParameterSetName = "One")]
		[String[]]$GPOName,
		[parameter(Mandatory = $True, ParameterSetName = "All")]
		[Switch]$All
	)
	BEGIN
	{
		TRY
		{
			if (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginIpmoAD }
			if (-not (Get-Module -Name GroupPolicy)) { Import-Module -Name GroupPolicy -ErrorAction Stop -ErrorVariable ErrorBeginIpmoGP }
		}
		CATCH
		{
			Write-Warning -Message "[BEGIN] Something wrong happened"
			IF ($ErrorBeginIpmoAD) { Write-Warning -Message "[BEGIN] Error while Importing the module Active Directory" }
			IF ($ErrorBeginIpmoGP) { Write-Warning -Message "[BEGIN] Error while Importing the module Group Policy" }
			Write-Warning -Message "[BEGIN] $($Error[0].exception.message)"
		}
	}
	PROCESS
	{
		FOREACH ($DomainController in ((Get-ADDomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetDC -filter *).hostname))
		{
			TRY
			{
				IF ($psBoundParameters['GPOName'])
				{
					Foreach ($GPOItem in $GPOName)
					{
						$GPO = Get-GPO -Name $GPOItem -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPO
						
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPOItem
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}#Foreach ($GPOItem in $GPOName)
				}#IF ($psBoundParameters['GPOName'])
				IF ($psBoundParameters['All'])
				{
					$GPOList = Get-GPO -All -Server $DomainController -ErrorAction Stop -ErrorVariable ErrorProcessGetGPOAll
					
					foreach ($GPO in $GPOList)
					{
						[pscustomobject][ordered] @{
							GroupPolicyName = $GPO.DisplayName
							DomainController = $DomainController
							UserVersion = $GPO.User.DSVersion
							UserSysVolVersion = $GPO.User.SysvolVersion
							ComputerVersion = $GPO.Computer.DSVersion
							ComputerSysVolVersion = $GPO.Computer.SysvolVersion
						}#PSObject
					}
				}#IF ($psBoundParameters['All'])
			}#TRY
			CATCH
			{
				Write-Warning -Message "[PROCESS] Something wrong happened"
				IF ($ErrorProcessGetDC) { Write-Warning -Message "[PROCESS] Error while running retrieving Domain Controllers with Get-ADDomainController" }
				IF ($ErrorProcessGetGPO) { Write-Warning -Message "[PROCESS] Error while running Get-GPO" }
				IF ($ErrorProcessGetGPOAll) { Write-Warning -Message "[PROCESS] Error while running Get-GPO -All" }
				Write-Warning -Message "[PROCESS] $($Error[0].exception.message)"
			}
		}#FOREACH
	}#PROCESS
}

function Get-NestedMember
{
<#
    .SYNOPSIS
        Find all Nested members of a group
    .DESCRIPTION
        Find all Nested members of a group
    .PARAMETER GroupName
        Specify one or more GroupName to audit
    .Example
        Get-NestedMember -GroupName TESTGROUP

        This will find all the indirect members of TESTGROUP
    .Example
        Get-NestedMember -GroupName TESTGROUP,TESTGROUP2

        This will find all the indirect members of TESTGROUP and TESTGROUP2
    .Example
        Get-NestedMember TESTGROUP | Group Name | select name, count

        This will find duplicate

#>
    [CmdletBinding()]
    PARAM(
    [String[]]$GroupName,
    [String]$RelationShipPath,
    [Int]$MaxDepth
    )
    BEGIN 
    {
        $DepthCount = 1

        TRY{
            if(-not(Get-Module Activedirectory -ErrorAction Stop)){
                Write-Verbose -Message "[BEGIN] Loading ActiveDirectory Module"
                Import-Module ActiveDirectory -ErrorAction Stop}
        }
        CATCH
        {
            Write-Warning -Message "[BEGIN] An Error occured"
            Write-Warning -Message $error[0].exception.message
        }
    }
    PROCESS
    {
        TRY
        {
            FOREACH ($Group in $GroupName)
            {
                # Get the Group Information
                $GroupObject = Get-ADGroup -Identity $Group -ErrorAction Stop
 
                IF($GroupObject)
                {
                    # Get the Members of the group
                    $GroupObject | Get-ADGroupMember -ErrorAction Stop | ForEach-Object -Process {
                        
                        # Get the name of the current group (to reuse in output)
                        $ParentGroup = $GroupObject.Name
                        

                        # Avoid circular
                        IF($RelationShipPath -notlike ".\ $($GroupObject.samaccountname) \*")
                        {
                            if($PSBoundParameters["RelationShipPath"]) {
                            
                                $RelationShipPath = "$RelationShipPath \ $($GroupObject.samaccountname)"
                            
                                }
                            Else{$RelationShipPath = ".\ $($GroupObject.samaccountname)"}

                            Write-Verbose -Message "[PROCESS] Name:$($_.name) | ObjectClass:$($_.ObjectClass)"
                            $CurrentObject = $_
                            switch ($_.ObjectClass)
                            {   
                                "group" {
                                    # Output Object
                                    $CurrentObject | Select-Object Name,SamAccountName,ObjectClass,DistinguishedName,@{Label="ParentGroup";Expression={$ParentGroup}}, @{Label="RelationShipPath";Expression={$RelationShipPath}}
                                
                                    if (-not($DepthCount -lt $MaxDepth)){
                                        # Find Child
                                        Get-NestedMember -GroupName $CurrentObject.Name -RelationShipPath $RelationShipPath
                                        $DepthCount++
                                    }
                                }#Group
                                default { $CurrentObject | Select-Object Name,SamAccountName,ObjectClass,DistinguishedName, @{Label="ParentGroup";Expression={$ParentGroup}},@{Label="RelationShipPath";Expression={$RelationShipPath}}}
                            }#Switch
                        }#IF($RelationShipPath -notmatch $($GroupObject.samaccountname))
                        ELSE{Write-Warning -Message "[PROCESS] Circular group membership detected with $($GroupObject.samaccountname)"}
                    }#ForeachObject
                }#IF($GroupObject)
                ELSE {
                    Write-Warning -Message "[PROCESS] Can't find the group $Group"
                }#ELSE
            }#FOREACH ($Group in $GroupName)
        }#TRY
        CATCH{
            Write-Warning -Message "[PROCESS] An Error occured"
            Write-Warning -Message $error[0].exception.message }
    }#PROCESS
    END
    {
        Write-Verbose -Message "[END] Get-NestedMember"
    }
}

function Get-ParentGroup
{
<#
    .SYNOPSIS
        Find all Nested members of a group
    .DESCRIPTION
        Find all Nested members of a group
    .PARAMETER GroupName
        Specify one or more GroupName to audit
    .Example
        Get-NestedMember -GroupName TESTGROUP

        This will find all the indirect members of TESTGROUP
    .Example
        Get-NestedMember -GroupName TESTGROUP,TESTGROUP2

        This will find all the indirect members of TESTGROUP and TESTGROUP2
    .Example
        Get-NestedMember TESTGROUP | Group Name | select name, count

        This will find duplicate

#>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory = $true)]
        [String[]]$Name
    )
    BEGIN 
    {
        TRY{
            if(-not(Get-Module Activedirectory -ErrorAction Stop)){
                Write-Verbose -Message "[BEGIN] Loading ActiveDirectory Module"
                Import-Module ActiveDirectory -ErrorAction Stop}
        }
        CATCH
        {
            Write-Warning -Message "[BEGIN] An Error occured"
            Write-Warning -Message $error[0].exception.message
        }
    }
    PROCESS
    {
        TRY
        {
            FOREACH ($Obj in $Name)
            {
                # Make an Ambiguous Name Resolution
                $ADObject = Get-ADObject -LDAPFilter "(|(anr=$obj)(distinguishedname=$obj))" -Properties memberof -ErrorAction Stop
                IF($ADObject)
                {
                    # Show a warning if more than 1 object is found
                    if ($ADObject.count -gt 1){Write-Warning -Message "More than one object found with the $obj request"}
                    
                    FOREACH ($Account in $ADObject)
                    {
                        Write-Verbose -Message "[PROCESS] $($Account.name)"
                        $Account | Select-Object -ExpandProperty memberof | ForEach-Object -Process {

                            $CurrentObject = Get-Adobject -LDAPFilter "(|(anr=$_)(distinguishedname=$_))" -Properties Samaccountname
                                
                            
                            Write-Output $CurrentObject | Select-Object Name,SamAccountName,ObjectClass, @{L="Child";E={$Account.samaccountname}}
                            
                            Write-Verbose -Message "Inception - $($CurrentObject.distinguishedname)"
                            Get-ParentGroup -OutBuffer $CurrentObject.distinguishedname

                        }#$Account | Select-Object
                    }#FOREACH ($Account in $ADObject){
                }#IF($ADObject)
                ELSE {
                    #Write-Warning -Message "[PROCESS] Can't find the object $Obj"
                }#ELSE
            }#FOREACH ($Obj in $Object)
        }#TRY
        CATCH{
            Write-Warning -Message "[PROCESS] An Error occured"
            Write-Warning -Message $error[0].exception.message }
    }#PROCESS
    END
    {
        Write-Verbose -Message "[END] Get-NestedMember"
    }

###################
# End AD Functions#
###################
#######################
# Office 365 Functions#
#######################

function Connect-Office365
{
<#
.SYNOPSIS
    This function will prompt for credentials, load module MSOLservice,
	load implicit modules for Office 365 Services (AD, Lync, Exchange) using PSSession.
.DESCRIPTION
    This function will prompt for credentials, load module MSOLservice,
	load implicit modules for Office 365 Services (AD, Lync, Exchange) using PSSession.
.EXAMPLE
    Connect-Office365
   
    This will prompt for your credentials and connect to the Office365 services
.EXAMPLE
    Connect-Office365 -verbose
   
    This will prompt for your credentials and connect to the Office365 services.
	Additionally you will see verbose messages on the screen to follow what is happening in the background
.NOTE
    Francois-Xavier Cat
    lazywinadmin.com
    @lazywinadm
#>
	[CmdletBinding()]
	PARAM ()
	BEGIN
	{
		TRY
		{
			#Modules
			IF (-not (Get-Module -Name MSOnline -ListAvailable))
			{
				Write-Verbose -Message "BEGIN - Import module Azure Active Directory"
				Import-Module -Name MSOnline -ErrorAction Stop -ErrorVariable ErrorBeginIpmoMSOnline
			}
			
			IF (-not (Get-Module -Name LyncOnlineConnector -ListAvailable))
			{
				Write-Verbose -Message "BEGIN - Import module Lync Online"
				Import-Module -Name LyncOnlineConnector -ErrorAction Stop -ErrorVariable ErrorBeginIpmoLyncOnline
			}
		}
		CATCH
		{
			Write-Warning -Message "BEGIN - Something went wrong!"
			IF ($ErrorBeginIpmoMSOnline)
			{
				Write-Warning -Message "BEGIN - Error while importing MSOnline module"
			}
			IF ($ErrorBeginIpmoLyncOnline)
			{
				Write-Warning -Message "BEGIN - Error while importing LyncOnlineConnector module"
			}
			
			Write-Warning -Message $error[0].exception.message
		}
	}
	PROCESS
	{
		TRY
		{
			
			# CREDENTIAL
			Write-Verbose -Message "PROCESS - Ask for Office365 Credential"
			$O365cred = Get-Credential -ErrorAction Stop -ErrorVariable ErrorCredential
			
			# AZURE ACTIVE DIRECTORY (MSOnline)
			Write-Verbose -Message "PROCESS - Connect to Azure Active Directory"
			Connect-MsolService -Credential $O365cred -ErrorAction Stop -ErrorVariable ErrorConnectMSOL
			
			# EXCHANGE ONLINE
			Write-Verbose -Message "PROCESS - Create session to Exchange online"
			$ExchangeURL = "https://ps.outlook.com/powershell/"
			$O365PS = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeURL -Credential $O365cred -Authentication Basic -AllowRedirection -ErrorAction Stop -ErrorVariable ErrorConnectExchange
			
			Write-Verbose -Message "PROCESS - Open session to Exchange online (Prefix: Cloud)"
			Import-PSSession -Session $O365PS –Prefix ExchCloud
			
			# LYNC ONLINE (LyncOnlineConnector)
			Write-Verbose -Message "PROCESS - Create session to Lync online"
			$lyncsession = New-CsOnlineSession –Credential $O365cred -ErrorAction Stop -ErrorVariable ErrorConnectExchange
			Import-PSSession -Session $lyncsession -Prefix LyncCloud
			
			# SHAREPOINT ONLINE
			#Connect-SPOService -Url https://contoso-admin.sharepoint.com –credential $O365cred
		}
		CATCH
		{
			Write-Warning -Message "PROCESS - Something went wrong!"
			IF ($ErrorCredential)
			{
				Write-Warning -Message "PROCESS - Error while gathering credential"
			}
			IF ($ErrorConnectMSOL)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Azure AD"
			}
			IF ($ErrorConnectExchange)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Exchange Online"
			}
			IF ($ErrorConnectLync)
			{
				Write-Warning -Message "PROCESS - Error while connecting to Lync Online"
			}
			
			Write-Warning -Message $error[0].exception.message
		}
	}
}