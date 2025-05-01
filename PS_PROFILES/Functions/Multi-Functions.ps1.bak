####################
# Static Functions #
####################
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
# Begin Remove-UserProfiles
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
}
# End Remove-UserProfiles
# Begin Remove-RemotePrintDrivers
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

}
# End Remove-RemotePrintDrivers
# Begin RDP
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
}
# End RDP
# Begin Get-Lastboot
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
}
# End Get-LastBoot
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
}
# End Get-HotFixes
# Begin Get-GPRemote

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
}
# End Get-GPRemote

# Begin CheckProcess
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
    
}
# End CheckProcess
# Begin Whois

Function WhoIs
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
 {
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
}
# End WhoIs
# Begin Get-NetworkStatistics
function Get-NetworkStatistics
<#
.SYNOPSIS
PowerShell version of netstat
.EXAMPLE
Get-NetworkStatistics
.EXAMPLE
Get-NetworkStatistics | where-object {$_.State -eq "LISTENING"} | Format-Table
#>
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
# Begin Update-SysInternals
function Update-Sysinternals
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
 {
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
# End Update-SysInternals
# Begin Get-ADGPOReplication
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
# Begin Get-LocalAdmin 
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

## Connect to VMware VSphere
Function GoGo-VSphere {

Connect-VIServer -Server 192.168.1.201
}

## End Connect to VMware VSphere

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

# Begin Get-RemoteSysInfo
function Get-RemoteSysInfo {
  <# 
  .SYNOPSIS 
  Retrieve basic system information for specified workstation(s) 

  .EXAMPLE 
  Get-RemoteSysInfo Computer123456 

  .EXAMPLE 
  Get-RemoteSysInfo 123456 
  #> 
param(

    [Parameter(Mandatory=$true)]
    [string[]] $ComputerName
)

$Stamp = (Get-Date -Format G) + ":"
$ComputerArray = @()

$i=0
$j=0

function Systeminformation {
	
    foreach ($Computer in $ComputerName) {

        if(!([String]::IsNullOrWhiteSpace($Computer))) {

            If (Test-Connection -quiet -count 1 -Computer $Computer) {

                Write-Progress -Activity "Getting Sytem Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerName.count) * 100) + "%") -CurrentOperation "Processing $($Computer)..." -PercentComplete ((($j++) / $ComputerName.count) * 100)

	            Start-Job -ScriptBlock { param($Computer) 

	                #Gather specified workstation information; CimInstance only works on 64-bit
	                $computerSystem = Get-CimInstance CIM_ComputerSystem -Computer $Computer
	                $computerBIOS = Get-CimInstance CIM_BIOSElement -Computer $Computer
	                $computerOS = Get-CimInstance CIM_OperatingSystem -Computer $Computer
	                $computerCPU = Get-CimInstance CIM_Processor -Computer $Computer
	                $computerHDD = Get-CimInstance Win32_LogicalDisk -Computer $Computer -Filter "DeviceID = 'C:'"
    
                        [pscustomobject]@{

                            "Computer Name"=$computerSystem.Name
                            "Last Reboot"=$computerOS.LastBootUpTime
                            "Operating System"=$computerOS.OSArchitecture + " " + $computerOS.caption
                             Model=$computerSystem.Model
                             RAM= "{0:N2}" -f [int]($computerSystem.TotalPhysicalMemory/1GB) + "GB"
                            "Disk Capacity"="{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
                            "Total Disk Space"="{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
                            "Current User"=$computerSystem.UserName
                        }
	            } -ArgumentList $Computer
            }

            else {

                Start-Job -ScriptBlock { param($Computer)  
                     
                    [pscustomobject]@{

                        "Computer Name"=$Computer
                        "Last Reboot"="Unable to PING."
                        "Operating System"="$Null"
                        Model="$Null"
                        RAM="$Null"
                        "Disk Capacity"="$Null"
                        "Total Disk Space"="$Null"
                        "Current User"="$Null"
                    }
                } -ArgumentList $Computer                       
            }
        }

        else {
                 
            Start-Job -ScriptBlock { param($Computer)  
                     
                [pscustomobject]@{

                    "Computer Name"="Value is null."
                    "Last Reboot"="$Null"
                    "Operating System"="$Null"
                    Model="$Null"
                    RAM="$Null"
                    "Disk Capacity"="$Null"
                    "Total Disk Space"="$Null"
                    "Current User"="$Null"
                }
            } -ArgumentList $Computer
        }
    } 
}

$SystemInformation = SystemInformation | Wait-Job | Receive-Job | Select "Computer Name", "Current User", "Operating System", Model, RAM, "Disk Capacity", "Total Disk Space", "Last Reboot"
$DocPath = [environment]::getfolderpath("mydocuments") + "\SystemInformation-Report.csv"

	Switch ($CheckBox.IsChecked){
		$true { $SystemInformation | Export-Csv $DocPath -NoTypeInformation -Force; }
		default { $SystemInformation | Out-GridView -Title "System Information"; }
		
    }

	if ($CheckBox.IsChecked -eq $true){

	    Try { 

		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {

		 #Do Nothing 
	    }
	}
	
	else{

	    Try {

	        $listBox.Items.Add("$stamp System Information output processed!`n")
	    } 

	    Catch {

	        #Do Nothing 
	    }
	}
}
# End Get-RemoteSysInfo
#Begin Get-RemoteSoftWare
function Get-RemoteSoftWare {
  <# 
  .SYNOPSIS 
  Grabs all installed Software on specified computer(s) 

  .EXAMPLE 
  Get-RemoteSoftWare Computer123456 

  .EXAMPLE 
  Get-RemoteSoftWare 123456 
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

function SoftwareCheck {

$i=0
$j=0

foreach ($computer in $ComputerArray) {

    Write-Progress -Activity "Retrieving Software Information..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $ComputerArray.count) * 100) + "%") -CurrentOperation "Processing $($computer)..." -PercentComplete ((($j++) / $ComputerArray.count) * 100)

        $keys = '','\Wow6432Node'
        foreach ($key in $keys) {
            try {
                $apps = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall").GetSubKeyNames()
            } catch {
                continue
            }

            foreach ($app in $apps) {
                $program = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$computer).OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall\$app")
                $name = $program.GetValue('DisplayName')
                if ($name -and $name -match $NameRegex) {
                    [pscustomobject]@{
                        "Computer Name" = $computer
                        Software = $name
                        Version = $program.GetValue('DisplayVersion')
                        Publisher = $program.GetValue('Publisher')
                        "Install Date" = $program.GetValue('InstallDate')
                        "Uninstall String" = $program.GetValue('UninstallString')
                        Bits = $(if ($key -eq '\Wow6432Node') {'64'} else {'32'})
                        Path = $program.name
                    }
                }
            }
        } 
    }
}	

foreach ($computer in $ComputerName) {	     
    If (Test-Connection -quiet -count 1 -Computer $Computer) {
		    
        $ComputerArray += $Computer
    }	
}
	$SoftwareCheck = SoftwareCheck | Sort "Computer Name" | Select "Computer Name", Software, Version, Publisher, "Install Date", "Uninstall String", Bits, Path
    	$DocPath = [environment]::getfolderpath("mydocuments") + "\Software-Report.csv"

    		Switch ($CheckBox.IsChecked){
    		    $true { $SoftwareCheck | Export-Csv $DocPath -NoTypeInformation -Force; }
    		    default { $SoftwareCheck | Out-GridView -Title "Software"; }
		}
		
	if ($CheckBox.IsChecked -eq $true){
	    Try { 
		$listBox.Items.Add("$stamp Export-CSV to $DocPath!`n")
	    } 

	    Catch {
		 #Do Nothing 
	    }
	}
	
	else{
	    Try {
	        $listBox.Items.Add("$stamp Software output processed!`n")
	    } 
	    Catch {
	        #Do Nothing 
	    }
	}
}
# End Get-RemoteSoftWare
# Begin Get-OfficeVersion
Function Get-OfficeVersion {
<#
.Synopsis
Gets the Office Version installed on the computer
.DESCRIPTION
This function will query the local or a remote computer and return the information about Office Products installed on the computer
.NOTES   
Name: Get-OfficeVersion
Version: 1.0.5
DateCreated: 2015-07-01
DateUpdated: 2016-07-20
.LINK
https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts
.PARAMETER ComputerName
The computer or list of computers from which to query 
.PARAMETER ShowAllInstalledProducts
Will expand the output to include all installed Office products
.EXAMPLE
Get-OfficeVersion
Description:
Will return the locally installed Office product
.EXAMPLE
Get-OfficeVersion -ComputerName client01,client02
Description:
Will return the installed Office product on the remote computers
.EXAMPLE
Get-OfficeVersion | select *
Description:
Will return the locally installed Office product with all of the available properties
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
       }
    }

    

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if (!($installPath)) { continue }
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE") -and $name.ToUpper() -notlike "*MUI*" -and $name.ToUpper() -notlike "*VISIO*" -and $name.ToUpper() -notlike "*PROJECT*") {
              $primaryOfficeProduct = $true
           }

           $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, "ClickToRunComponent").uValue
           $uninstallString = $regProv.GetStringValue($HKLM, $path, "UninstallString").sValue
           if (!($clickToRunComponent)) {
              if ($uninstallString) {
                 if ($uninstallString.Contains("OfficeClickToRun")) {
                     $clickToRunComponent = $true
                 }
              }
           }

           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 
           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false

           if ($clickToRunComponent) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }

  }

  $results = Get-Unique -InputObject $results 

  return $results;
}

}
# End Get-OfficeVersion
# Begin Get-OfficeVersion2

function Get-OfficeVersion2
{
param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string] $Infile,
    
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string] $outfile
    )
#$outfile = 'C:\temp\office.csv'
#$infile = 'c:\temp\servers.txt'
Begin
    {
    }
 Process
    {
    $office = @()
    $computers = Get-Content $infile
    $i=0
    $count = $computers.count
    foreach($computer in $computers)
     {
     $i++
     Write-Progress -Activity "Querying Computers" -Status "Computer: $i of $count " `
      -PercentComplete ($i/$count*100)
        $info = @{}
        $version = 0
        try{
          $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer) 
          $reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() |% {
            if ($_ -match '(\d+)\.') {
              if ([int]$matches[1] -gt $version) {
                $version = $matches[1]
              }
            }    
          }
          if ($version) {
            Write-Debug("$computer : found $version")
            switch($version) {
                "7" {$officename = 'Office 97' }
                "8" {$officename = 'Office 98' }
                "9" {$officename = 'Office 2000' }
                "10" {$officename = 'Office XP' }
                "11" {$officename = 'Office 97' }
                "12" {$officename = 'Office 2003' }
                "13" {$officename = 'Office 2007' }
                "14" {$officename = 'Office 2010' }
                "15" {$officename = 'Office 2013' }
                "16" {$officename = 'Office 2016' }
                default {$officename = 'Unknown Version'}
            }
    
          }
          }
          catch{
              $officename = 'Not Installed/Not Available'
          }
    $info.Computer = $computer
    $info.Name= $officename
    $info.version =  $version

    $object = new-object -TypeName PSObject -Property $info
    $office += $object
    }
    $office | select computer,version,name | Export-Csv -NoTypeInformation -Path $outfile
    }
}
  write-output ("Done")
  # End Get-OfficeVersion2
  # Begin Get-OutlookClientVersion
  
function Get-OutlookClientVersion {
 
<#
.SYNOPSIS
    Identifies and reports which Outlook client versions are being used to access Exchange.
 
.DESCRIPTION
    Get-MrRCAProtocolLog is an advanced PowerShell function that parses Exchange Server RPC
    logs to determine what Outlook client versions are being used to access the Exchange Server.
 
.PARAMETER LogFile
    The path to the Exchange RPC log files.
 
.EXAMPLE
     Get-MrRCAProtocolLog -LogFile 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\RCA_20140831-1.LOG'
 
.EXAMPLE
     Get-ChildItem -Path '\\servername\c$\Program Files\Microsoft\Exchange Server\V15\Logging\RPC Client Access\*.log' |
     Get-MrRCAProtocolLog |
     Out-GridView -Title 'Outlook Client Versions'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>
 
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
                   ValueFromPipeline)]
        [ValidateScript({
            Test-Path -Path $_ -PathType Leaf -Include '*.log'
        })]
        [string[]]$LogFile
    )
 
    PROCESS {
        foreach ($file in $LogFile) {
            $Headers = (Get-Content -Path $file -TotalCount 5 | Where-Object {$_ -like '#Fields*'}) -replace '#Fields: ' -split ','
                    
            Import-Csv -Header $Headers -Path $file |
            Where-Object {$_.operation -eq 'Connect' -and $_.'client-software' -eq 'outlook.exe'} |
            Select-Object -Unique -Property @{label='User';expression={$_.'client-name' -replace '^.*cn='}},
                                            @{label='DN';expression={$_.'client-name'}},
                                            client-software,
                                            @{label='Version';expression={Get-MrOutlookVersion -OutlookBuild $_.'client-software-version'}},
                                            client-mode,
                                            client-ip,
                                            protocol
        }
    }
}
 
function Get-MrOutlookVersion {
    param (
        [string]$OutlookBuild
    )
    switch ($OutlookBuild) {  
        {$_ -ge '16.0.4266.1001'} {'Outlook 2016 4266.1001'; break}
        {$_ -ge '16.0.4522.1000'} {'Outlook 2016 4522.1000'; break}
        {$_ -ge '16.0.4498.1000'} {'Outlook 2016 4498.1000'; break}
        {$_ -ge '16.0.4229.1024'} {'Outlook 2016 4229.1024'; break}              
        {$_ -ge '15.0.4569.1506'} {'Outlook 2013 SP1'; break}
        {$_ -ge '15.0.4420.1017'} {'Outlook 2013 RTM'; break}
        {$_ -ge '14.0.7015.1000'} {'Outlook 2010 SP2'; break}
        {$_ -ge '14.0.6029.1000'} {'Outlook 2010 SP1'; break}
        {$_ -ge '14.0.4763.1000'} {'Outlook 2010 RTM'; break}
        {$_ -ge '12.0.6672.5000'} {'Outlook 2007 SP3 U2013'; break}
        {$_ -ge '12.0.6423.1000'} {'Outlook 2007 SP2'; break}
        {$_ -ge '12.0.6212.1000'} {'Outlook 2007 SP1'; break}
        {$_ -ge '12.0.4518.1014'} {'Outlook 2007 RTM'; break}
        Default {'$OutlookBuild'}
    }
}
# End Get-OutlookClientVersion
# Begin VMware Functions

<# Enable or Disable Hot Add Memory/CPU
 Enable-MemHotAdd $ServerName
 Disable-MemHotAdd $ServerName
 Enable-vCPUHotAdd $ServerName
 Disable-vCPUHotAdd $ServerName
#>

Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End VMware Functions
# Begin Get-DefragAnalysis
Function Get-DefragAnalysis {

<#
.Synopsis
Run a defrag analysis.
.Description
This command uses WMI to to run a defrag analysis on selected volumes on
local or remote computers. You will get a custom object for each volume like
this:

AverageFileSize               : 64
AverageFragmentsPerFile       : 1
AverageFreeSpacePerExtent     : 17002496
ClusterSize                   : 4096
ExcessFolderFragments         : 0
FilePercentFragmentation      : 0
FragmentedFolders             : 0
FreeSpace                     : 161816576
FreeSpacePercent              : 77
FreeSpacePercentFragmentation : 29
LargestFreeSpaceExtent        : 113500160
MFTPercentInUse               : 100
MFTRecordCount                : 511
PageFileSize                  : 0
TotalExcessFragments          : 0
TotalFiles                    : 182
TotalFolders                  : 11
TotalFragmentedFiles          : 0
TotalFreeSpaceExtents         : 8
TotalMFTFragments             : 1
TotalMFTSize                  : 524288
TotalPageFileFragments        : 0
TotalPercentFragmentation     : 0
TotalUnmovableFiles           : 4
UsedSpace                     : 47894528
VolumeName                    : 
VolumeSize                    : 209711104
Driveletter                   : E:
DefragRecommended             : False
Computername                  : NOVO8

The default drive is C: on the local computer.
.Example
PS C:\> Get-DefragAnalysis
Run a defrag analysis on C: on the local computer
.Example
PS C:\> Get-DefragAnalysis -drive "C:" -computername $servers
Run a defrag analysis for drive C: on a previously defined collection of server names.
.Example
PS C:\> $data = Get-WmiObject Win32_volume -filter "driveletter like '%' AND drivetype=3" -ComputerName Novo8 | Get-DefragAnalysis
PS C:\> $data | Sort Driveletter | Select Computername,DriveLetter,DefragRecommended

Computername                    Driveletter                     DefragRecommended
------------                    -----------                     -----------------
NOVO8                           C:                                          False
NOVO8                           D:                                          True
NOVO8                           E:                                          False

Get all volumes on a remote computer that are fixed but have a drive letter,
this should eliminate CD/DVD drives, and run a defrag analysis on each one.
The results are saved to a variable, $data.
.Notes
Last Updated: 12/5/2012
Author      : Jeffery Hicks (http://jdhitsolutions.com/blog)
Version     : 0.9

.Link
Get-WMIObject
Invoke-WMIMethod

#>

[cmdletbinding(SupportsShouldProcess=$True)]

Param(
[Parameter(Position=0,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("drive")]
[string]$Driveletter="C:",
[Parameter(Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[ValidateNotNullorEmpty()]
[Alias("PSComputername","SystemName")]
[string[]]$Computername=$env:computername
)

Begin {
    Write-Verbose -Message "$(Get-Date) Starting $($MyInvocation.Mycommand)"   
} #close Begin

Process {
    #strip off any extra spaces on the drive letter just in case
    Write-Verbose "$(Get-Date) Processing $Driveletter"
    $Driveletter=$Driveletter.Trim()
    if ($Driveletter.length -gt 2) {
        Write-Verbose "$(Get-Date) Scrubbing drive parameter value"
        $Driveletter=$Driveletter.Substring(0,2)
    }
    #add a colon if not included
    if ($Driveletter -match "^\w$") {
        Write-Verbose "$(Get-Date) Modifying drive parameter value"
        $Driveletter="$($Driveletter):"
    }

    Write-Verbose "$(Get-Date) Analyzing drive $Driveletter"
        
    Foreach ($computer in $computername) {
        Write-Verbose "$(Get-Date) Examining $computer"
        Try {
            $volume=Get-WmiObject -Class Win32_Volume -filter "DriveLetter='$Driveletter'" -computername $computer -errorAction "Stop"
        }
        Catch {
            Write-Warning ("Failed to get volume {0} from  {1}. {2}" -f $driveletter,$computer,$_.Exception.Message)
        }
        if ($volume) {
            Write-Verbose "$(Get-Date) Running defrag analysis"
            $analysis = $volume | Invoke-WMIMethod -name DefragAnalysis
        
            #get properties for DefragAnalysis so we can filter out system properties
            $analysis.DefragAnalysis.Properties | 
            Foreach -begin {$Prop=@()} -process { $Prop+=$_.Name }
        
            Write-Verbose "$(Get-Date) Retrieving results"
            $analysis | Select @{Name="Results";Expression={$_.DefragAnalysis | 
            Select-Object -Property $Prop |
            Foreach-Object { 
              #Add on some additional property values
              $_ | Add-member -MemberType Noteproperty -Name Driveletter -value $DriveLetter
              $_ | Add-member -MemberType Noteproperty -Name DefragRecommended -value $analysis.DefragRecommended 
              $_ | Add-member -MemberType Noteproperty -Name Computername -value $volume.__SERVER -passthru
             } #foreach-object
            }}  | Select -expand Results 
            
            #clean up variables so there are no accidental leftovers
            Remove-Variable "volume","analysis"
        } #close if volume
     } #close Foreach computer
 } #close Process
 
End {
    Write-Verbose "$(Get-Date) Defrag analysis complete"
} #close End

} #close function
# End Get-DefragAnalysis
# Begin Get-NetworkInfo
Function Get-NetworkInfo {
    <#   
        .SYNOPSIS   
            Retrieves the network configuration from a local or remote client.      
             
        .DESCRIPTION   
            Retrieves the network configuration from a local or remote client.        
        
        .PARAMETER Computername
            A single or collection of systems to perform the query against
        
        .PARAMETER Credential
            Alternate credentials to use for query of network information        
        
        .PARAMETER Throttle
            Number of asynchonous jobs that will run at a time
        
        .NOTES   
            Name: Get-NetworkInfo.ps1
            Author: Boe Prox
            Version: 1.0
        
        .EXAMPLE 
             Get-NetworkInfo -Computername 'System1'
            
            NICDescription : Ethernet Network Adapter
            MACAddress     : 00:11:22:33:aa:bb
            NICName        : enthad
            Computername   : System1.domain.com
            DHCPEnabled    : True
            WINSPrimary    : 192.0.0.25
            SubnetMask     : {255.255.255.255}
            WINSSecondary  : 192.0.0.26
            DNSServer      : {192.0.0.31, 192.0.0.30}
            IPAddress      : {192.0.0.5}
            DefaultGateway : {192.0.0.1}         
             
            Description 
            ----------- 
            Retrieves the network information from 'System1'      

        .EXAMPLE
            $Servers = Get-Content Servers.txt
            $Servers | Get-NetworkInfo -Throttle 10
            
            Description
            -----------
            Retrieves all of network information from the remote servers while running 10 runspace jobs at a time.  
            
        .EXAMPLE
            (Get-Content Servers.txt) | Get-NetworkInfo -Credential domain\adminuser -Throttle 10
            
            Description
            -----------
            Gathers all of the network information from the systems in the text file. Also uses alternate administrator credentials provided.                                            
    #>
    
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline = $True,ValueFromPipeLineByPropertyName = $True)]
        [Alias('CN','__Server','IPAddress','Server')]
        [string[]]$Computername = $Env:Computername,
        
        [parameter()]
        [Alias('RunAs')]
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,       
        
        [parameter()]
        [int]$Throttle = 15
    )
    Begin {
        #Function that will be used to process runspace jobs
        Function Get-RunspaceData {
            [cmdletbinding()]
            param(
                [switch]$Wait
            )
            Do {
                $more = $false         
                Foreach($runspace in $runspaces) {
                    If ($runspace.Runspace.isCompleted) {
                        $runspace.powershell.EndInvoke($runspace.Runspace)
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null
                        $Script:i++                  
                    } ElseIf ($runspace.Runspace -ne $null) {
                        $more = $true
                    }
                }
                If ($more -AND $PSBoundParameters['Wait']) {
                    Start-Sleep -Milliseconds 100
                }   
                #Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | Where {
                    $_.runspace -eq $Null
                } | ForEach {
                    Write-Verbose ("Removing {0}" -f $_.computer)
                    $Runspaces.remove($_)
                }             
            } while ($more -AND $PSBoundParameters['Wait'])
        }
            
        Write-Verbose ("Performing inital Administrator check")
        $usercontext = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $IsAdmin = $usercontext.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")                   
        
        #Main collection to hold all data returned from runspace jobs
        $Script:report = @()    
        
        Write-Verbose ("Building hash table for WMI parameters")
        $WMIhash = @{
            Class = "Win32_NetworkAdapterConfiguration"
            Filter = "IPEnabled='$True'"
            ErrorAction = "Stop"
        } 
        
        #Supplied Alternate Credentials?
        If ($PSBoundParameters['Credential']) {
            $wmihash.credential = $Credential
        }
        
        #Define hash table for Get-RunspaceData function
        $runspacehash = @{}

        #Define Scriptblock for runspaces
        $scriptblock = {
            Param (
                $Computer,
                $wmihash
            )           
            Write-Verbose ("{0}: Checking network connection" -f $Computer)
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                #Check if running against local system and perform necessary actions
                Write-Verbose ("Checking for local system")
                If ($Computer -eq $Env:Computername) {
                    $wmihash.remove('Credential')
                } Else {
                    $wmihash.Computername = $Computer
                }
                Try {
                        Get-WmiObject @WMIhash | ForEach {
                            $IpHash =  @{
                                Computername = $_.DNSHostName
                                DNSDomain = $_.DNSDomain
                                IPAddress = $_.IpAddress
                                SubnetMask = $_.IPSubnet
                                DefaultGateway = $_.DefaultIPGateway
                                DNSServer = $_.DNSServerSearchOrder
                                DHCPEnabled = $_.DHCPEnabled
                                MACAddress  = $_.MACAddress
                                WINSPrimary = $_.WINSPrimaryServer
                                WINSSecondary = $_.WINSSecondaryServer
                                NICName = $_.ServiceName
                                NICDescription = $_.Description
                            }
                            $IpStack = New-Object PSObject -Property $IpHash
                            #Add a unique object typename
                            $IpStack.PSTypeNames.Insert(0,"IPStack.Information")
                            $IpStack 
                        }
                    } Catch {
                        Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
                        Break
                }
            } Else {
                Write-Warning ("{0}: Unavailable!" -f $Computer)
                Break
            }        
        }
        
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.Open()  
        
        Write-Verbose ("Creating empty collection to hold runspace jobs")
        $Script:runspaces = New-Object System.Collections.ArrayList        
    }
    Process {        
        $totalcount = $computername.count
        Write-Verbose ("Validating that current user is Administrator or supplied alternate credentials")        
        If (-Not ($Computername.count -eq 1 -AND $Computername[0] -eq $Env:Computername)) {
            #Now check that user is either an Administrator or supplied Alternate Credentials
            If (-Not ($IsAdmin -OR $PSBoundParameters['Credential'])) {
                Write-Warning ("You must be an Administrator to perform this action against remote systems!")
                Break
            }
        }
        ForEach ($Computer in $Computername) {
           #Create the powershell instance and supply the scriptblock with the other parameters 
           $powershell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($computer).AddArgument($wmihash)
           
           #Add the runspace into the powershell instance
           $powershell.RunspacePool = $runspacepool
           
           #Create a temporary collection for each runspace
           $temp = "" | Select-Object PowerShell,Runspace,Computer
           $Temp.Computer = $Computer
           $temp.PowerShell = $powershell
           
           #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
           $temp.Runspace = $powershell.BeginInvoke()
           Write-Verbose ("Adding {0} collection" -f $temp.Computer)
           $runspaces.Add($temp) | Out-Null
           
           Write-Verbose ("Checking status of runspace jobs")
           Get-RunspaceData @runspacehash
        }                        
    }
    End {                     
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@(($runspaces | Where {$_.Runspace -ne $Null}).Count)))
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash
        
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()               
    }
}
# End Get-NetworkInfo
# Begin Get-SNMPTrap
function Get-SnmpTrap {
<#
.SYNOPSIS
Function that will list SNMP Community string, Security Options and Trap Configuration for SNMP version 1 and version 2c.
.DESCRIPTION
** This function will list SNMP settings of  windows server by reading the registry keys under HKLM\SYSTEM\CurrentControlSet\services\SNMP\Parameters\
Example usage:																					  
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server. The meaning of each column is:
AcceptedCommunityStrings => The community string that the SNMP agent is allowed to receive. If the host is not requested with one of these pre-defined 
community strings, then the host will send an authentication trap.
AllowedHosts => The hostnames or IP addresses from which SNMP agent will accept SNMP messages.
CommunityRights => The permission that determines how the SNMP agent processes the incoming request from various communities.
TrapCommunityNames => When an SNMP agent receives a request that does not contain a valid community name or the host that is sending the message 
is not on the list of acceptable hosts, the agent can send an authentication trap message to one or more trap destinations (management systems)
TrapDestinations => The host names or IP addresses of trap destinations which are defined under the TrapCommunityNames.
SendTrap => It indicates whether sending autentication trap is enabled.
Author: phyoepaing3.142@gmail.com
Country: Myanmar(Burma)
Released: 05/07/2017
.EXAMPLE
Get_SnmpTrap
This will list the  SNMP Community string, Security Options and Trap Configuration on the server.
.LINK
You can find this script and more at: https://www.sysadminplus.blogspot.com/
#>

### DATA lookup section to convert registry numeric to corresponding output ###
$ConvertRights = DATA { ConvertFrom-StringData -StringData @'
1 = NONE
2 = NOTIFY
4 = READ-ONLY
8 = READ-WRITE
16 = READ-CREATE
'@}

$rh = '2147483650';  ## This number represents HKLM
$key1 = 'SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers';
$reg = [wmiclass]"\\localhost\root\default:StdRegprov"; 
$obj = New-Object -TypeName PsObject -Property @{AllowedHosts=@(); AcceptedCommunityStrings="";  CommunityRights =@(); TrapCommunityNames=@(); TrapDestinations=@(); SendTrap="" }; 
$AccessDenied = 0;

### Read the registry to find the allowed hosts for incoming community string ###
$i=1;
while ( $reg.GetStringValue($rh, $key1, $i ).sValue )
	{
	$obj.AllowedHosts += $reg.GetStringValue($rh, $key1, $i ).sValue; 
	$i ++;
	}
If ($obj.AllowedHosts.count -eq 1)
	{
	$obj.AllowedHosts = $obj.AllowedHosts[0];
	}

### Read the Community Strings ###	
Try {
	$obj.AcceptedCommunityStrings = (Gi -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities -EA Stop).Property; 

## If there is only one community string, then convert the property type to string from array ###
	If ($obj.AcceptedCommunityStrings.count -eq 1)
	{ $obj.AcceptedCommunityStrings = $obj.AcceptedCommunityStrings[0] }

### If there are multiple community strings, then read through all the security permission of each community string 	via registry ##
	If ($obj.AcceptedCommunityStrings -is [array])
	{	
		$obj.AcceptedCommunityStrings | foreach {
		$securityRight =  [string]((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$_)
		$obj.CommunityRights += $_+":"+$ConvertRights[$securityRight]
			}
		}
	else
		{
		[string]$securityRight = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities).$($obj.AcceptedCommunityStrings)
		$obj.CommunityRights = $ConvertRights[$securityRight]
		}
	}
catch [System.Security.SecurityException]
	{ 
	Write-Host -fore red "Access to Registry is denied. Please make sure you have permission to access registry and run in the elevated command prompt.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	}
catch 
	{ 
	Write-Host -fore red "SMNP Service is not installed on one or more servers.`n"; 
	$obj.AllowedHosts = "N/A"
	$obj.AcceptedCommunityStrings = "N/A"
	$obj.CommunityRights = "N/A"
	$obj.TrapCommunityNames = "N/A"
	$obj.TrapDestinations = "N/A"
	$obj.SendTrap = "N/A"
	$AccessDenied = 1; 
	$obj;
	}

## If the read of registry is not access-denied from previous try-catch statement, then continue ##	
If (!$AccessDenied)	
	{
	Try {
		$TrapConfig = Gci -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration -EA Stop ;
		$TrapConfig | foreach {
			$obj.TrapCommunityNames += $_.PsChildName
			}
	If ($obj.TrapCommunityNames.count -eq 1)	
	{ $obj.TrapCommunityNames = $obj.TrapCommunityNames[0]	}
			
		}
	catch 
		{  }
		
### Find destination for each Trap. The trap's community name will be prefixed on the trap's destination IP/hosts if there are multiple Traps configured, if it's single trap, then use without prefix ###

If ($obj.TrapCommunityNames -is [array])
	{
		$obj.TrapCommunityNames | foreach {
		$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$_";
		$i=1;
			while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $_+":"+$reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	}
else
	{
	$key2 = "SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\$($obj.TrapCommunityNames)";
	$i=1;
		while ( $reg.GetStringValue($rh, $key2, $i ).sValue )
				{
				$obj.TrapDestinations += $reg.GetStringValue($rh, $key2, $i ).sValue; 
				$i ++;
				}
		}
	
### If there is only one entry in the Trap Destination, then convert  the array to string ###
If ($obj.TrapDestinations.count -eq 1)
	{
	$obj.TrapDestinations = $obj.TrapDestinations[0];
	}
	
#### Check if the 'Send Authentication Trap' check box is enabled ###
	Switch ((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters).EnableAuthenticationTraps)
		{
		"0" { $obj.SendTrap = "Disabled" }
		"1" { $obj.SendTrap = "Enabled "}		
		}
	$obj;	
	}
}
# End Get-SNMPTrap
# Begin Get-UserLogon
function Get-UserLogon {
 
[CmdletBinding()]
 
param
 
(
 
[Parameter ()]
[String]$Computer,
 
[Parameter ()]
[String]$OU,
 
[Parameter ()]
[Switch]$All
 
)
 
$ErrorActionPreference="SilentlyContinue"
 
$result=@()
 
If ($Computer) {
 
Invoke-Command -ComputerName $Computer -ScriptBlock {quser} | Select-Object -Skip 1 | Foreach-Object {
 
$b=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($b[2] -like 'Disc*') {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[4]
'Time' = $b[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $b[0]
'Computer' = $Computer
'Date' = $b[5]
'Time' = $b[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
}
}
 
If ($OU) {
 
$comp=Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
}
 
}
 
}
 
}
 
If ($All) {
 
$comp=Get-ADComputer -Filter * -Properties operatingsystem
 
$count=$comp.count
 
If ($count -gt 20) {
 
Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
 
}
 
foreach ($u in $comp) {
 
Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
 
$a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
 
If ($a[2] -like '*Disc*') {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[4]
'Time' = $a[5..6] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
else {
 
$array= ([ordered]@{
'User' = $a[0]
'Computer' = $u.Name
'Date' = $a[5]
'Time' = $a[6..7] -join ' '
})
 
$result+=New-Object -TypeName PSCustomObject -Property $array
 
}
 
}
 
}
}
Write-Output $result
}
# End Get-UserLogon
# Begin Invoke-Ping
Function Invoke-Ping
{
<#
.SYNOPSIS
    Ping or test connectivity to systems in parallel
    
.DESCRIPTION
    Ping or test connectivity to systems in parallel

    Default action will run a ping against systems
        If Quiet parameter is specified, we return an array of systems that responded
        If Detail parameter is specified, we test WSMan, RemoteReg, RPC, RDP and/or SMB

.PARAMETER ComputerName
    One or more computers to test

.PARAMETER Quiet
    If specified, only return addresses that responded to Test-Connection

.PARAMETER Detail
    Include one or more additional tests as specified:
        WSMan      via Test-WSMan
        RemoteReg  via Microsoft.Win32.RegistryKey
        RPC        via WMI
        RDP        via port 3389
        SMB        via \\ComputerName\C$
        *          All tests

.PARAMETER Timeout
    Time in seconds before we attempt to dispose an individual query.  Default is 20

.PARAMETER Throttle
    Throttle query to this many parallel runspaces.  Default is 100.

.PARAMETER NoCloseOnTimeout
    Do not dispose of timed out tasks or attempt to close the runspace if threads have timed out

    This will prevent the script from hanging in certain situations where threads become non-responsive, at the expense of leaking memory within the PowerShell host.

.EXAMPLE
    Invoke-Ping Server1, Server2, Server3 -Detail *

    # Check for WSMan, Remote Registry, Remote RPC, RDP, and SMB (via C$) connectivity against 3 machines

.EXAMPLE
    $Computers | Invoke-Ping

    # Ping computers in $Computers in parallel

.EXAMPLE
    $Responding = $Computers | Invoke-Ping -Quiet
    
    # Create a list of computers that successfully responded to Test-Connection

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Invoke-Ping-Test-in-b553242a

.FUNCTIONALITY
    Computers
	
.NOTES
	Warren F
	http://ramblingcookiemonster.github.io/Invoke-Ping/

#>
	[cmdletbinding(DefaultParameterSetName = 'Ping')]
	param (
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[string[]]$ComputerName,
		
		[Parameter(ParameterSetName = 'Detail')]
		[validateset("*", "WSMan", "RemoteReg", "RPC", "RDP", "SMB")]
		[string[]]$Detail,
		
		[Parameter(ParameterSetName = 'Ping')]
		[switch]$Quiet,
		
		[int]$Timeout = 20,
		
		[int]$Throttle = 100,
		
		[switch]$NoCloseOnTimeout
	)
	Begin
	{
		
		#http://gallery.technet.microsoft.com/Run-Parallel-Parallel-377fd430
		function Invoke-Parallel
		{
			[cmdletbinding(DefaultParameterSetName = 'ScriptBlock')]
			Param (
				[Parameter(Mandatory = $false, position = 0, ParameterSetName = 'ScriptBlock')]
				[System.Management.Automation.ScriptBlock]$ScriptBlock,
				
				[Parameter(Mandatory = $false, ParameterSetName = 'ScriptFile')]
				[ValidateScript({ test-path $_ -pathtype leaf })]
				$ScriptFile,
				
				[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
				[Alias('CN', '__Server', 'IPAddress', 'Server', 'ComputerName')]
				[PSObject]$InputObject,
				
				[PSObject]$Parameter,
				
				[switch]$ImportVariables,
				
				[switch]$ImportModules,
				
				[int]$Throttle = 20,
				
				[int]$SleepTimer = 200,
				
				[int]$RunspaceTimeout = 0,
				
				[switch]$NoCloseOnTimeout = $false,
				
				[int]$MaxQueue,
				
				[validatescript({ Test-Path (Split-Path $_ -parent) })]
				[string]$LogFile = "C:\temp\log.log",
				
				[switch]$Quiet = $false
			)
			
			Begin
			{
				
				#No max queue specified?  Estimate one.
				#We use the script scope to resolve an odd PowerShell 2 issue where MaxQueue isn't seen later in the function
				if (-not $PSBoundParameters.ContainsKey('MaxQueue'))
				{
					if ($RunspaceTimeout -ne 0) { $script:MaxQueue = $Throttle }
					else { $script:MaxQueue = $Throttle * 3 }
				}
				else
				{
					$script:MaxQueue = $MaxQueue
				}
				
				Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"
				
				#If they want to import variables or modules, create a clean runspace, get loaded items, use those to exclude items
				if ($ImportVariables -or $ImportModules)
				{
					$StandardUserEnv = [powershell]::Create().addscript({
						
						#Get modules and snapins in this clean runspace
						$Modules = Get-Module | Select -ExpandProperty Name
						$Snapins = Get-PSSnapin | Select -ExpandProperty Name
						
						#Get variables in this clean runspace
						#Called last to get vars like $? into session
						$Variables = Get-Variable | Select -ExpandProperty Name
						
						#Return a hashtable where we can access each.
						@{
							Variables = $Variables
							Modules = $Modules
							Snapins = $Snapins
						}
					}).invoke()[0]
					
					if ($ImportVariables)
					{
						#Exclude common parameters, bound parameters, and automatic variables
						Function _temp { [cmdletbinding()]
							param () }
						$VariablesToExclude = @((Get-Command _temp | Select -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables)
						Write-Verbose "Excluding variables $(($VariablesToExclude | sort) -join ", ")"
						
						# we don't use 'Get-Variable -Exclude', because it uses regexps. 
						# One of the veriables that we pass is '$?'. 
						# There could be other variables with such problems.
						# Scope 2 required if we move to a real module
						$UserVariables = @(Get-Variable | Where { -not ($VariablesToExclude -contains $_.Name) })
						Write-Verbose "Found variables to import: $(($UserVariables | Select -expandproperty Name | Sort) -join ", " | Out-String).`n"
						
					}
					
					if ($ImportModules)
					{
						$UserModules = @(Get-Module | Where { $StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue) } | Select -ExpandProperty Path)
						$UserSnapins = @(Get-PSSnapin | Select -ExpandProperty Name | Where { $StandardUserEnv.Snapins -notcontains $_ })
					}
				}
				
				#region functions
				
				Function Get-RunspaceData
				{
					[cmdletbinding()]
					param ([switch]$Wait)
					
					#loop through runspaces
					#if $wait is specified, keep looping until all complete
					Do
					{
						
						#set more to false for tracking completion
						$more = $false
						
						#Progress bar if we have inputobject count (bound parameter)
						if (-not $Quiet)
						{
							Write-Progress -Activity "Running Query" -Status "Starting threads"`
										   -CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
										   -PercentComplete $(Try { $script:completedCount / $totalCount * 100 }
							Catch { 0 })
						}
						
						#run through each runspace.           
						Foreach ($runspace in $runspaces)
						{
							
							#get the duration - inaccurate
							$currentdate = Get-Date
							$runtime = $currentdate - $runspace.startTime
							$runMin = [math]::Round($runtime.totalminutes, 2)
							
							#set up log object
							$log = "" | select Date, Action, Runtime, Status, Details
							$log.Action = "Removing:'$($runspace.object)'"
							$log.Date = $currentdate
							$log.Runtime = "$runMin minutes"
							
							#If runspace completed, end invoke, dispose, recycle, counter++
							If ($runspace.Runspace.isCompleted)
							{
								
								$script:completedCount++
								
								#check if there were errors
								if ($runspace.powershell.Streams.Error.Count -gt 0)
								{
									
									#set the logging info and move the file to completed
									$log.status = "CompletedWithErrors"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
									foreach ($ErrorRecord in $runspace.powershell.Streams.Error)
									{
										Write-Error -ErrorRecord $ErrorRecord
									}
								}
								else
								{
									
									#add logging details and cleanup
									$log.status = "Completed"
									Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								}
								
								#everything is logged, clean up the runspace
								$runspace.powershell.EndInvoke($runspace.Runspace)
								$runspace.powershell.dispose()
								$runspace.Runspace = $null
								$runspace.powershell = $null
								
							}
							
							#If runtime exceeds max, dispose the runspace
							ElseIf ($runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout)
							{
								
								$script:completedCount++
								$timedOutTasks = $true
								
								#add logging details and cleanup
								$log.status = "TimedOut"
								Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
								Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"
								
								#Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
								if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
								$runspace.Runspace = $null
								$runspace.powershell = $null
								$completedCount++
								
							}
							
							#If runspace isn't null set more to true  
							ElseIf ($runspace.Runspace -ne $null)
							{
								$log = $null
								$more = $true
							}
							
							#log the results if a log file was indicated
							if ($logFile -and $log)
							{
								($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
							}
						}
						
						#Clean out unused runspace jobs
						$temphash = $runspaces.clone()
						$temphash | Where { $_.runspace -eq $Null } | ForEach {
							$Runspaces.remove($_)
						}
						
						#sleep for a bit if we will loop again
						if ($PSBoundParameters['Wait']) { Start-Sleep -milliseconds $SleepTimer }
						
						#Loop again only if -wait parameter and there are more runspaces to process
					}
					while ($more -and $PSBoundParameters['Wait'])
					
					#End of runspace function
				}
				
				#endregion functions
				
				#region Init
				
				if ($PSCmdlet.ParameterSetName -eq 'ScriptFile')
				{
					$ScriptBlock = [scriptblock]::Create($(Get-Content $ScriptFile | out-string))
				}
				elseif ($PSCmdlet.ParameterSetName -eq 'ScriptBlock')
				{
					#Start building parameter names for the param block
					[string[]]$ParamsToAdd = '$_'
					if ($PSBoundParameters.ContainsKey('Parameter'))
					{
						$ParamsToAdd += '$Parameter'
					}
					
					$UsingVariableData = $Null
					
					
					# This code enables $Using support through the AST.
					# This is entirely from  Boe Prox, and his https://github.com/proxb/PoshRSJob module; all credit to Boe!
					
					if ($PSVersionTable.PSVersion.Major -gt 2)
					{
						#Extract using references
						$UsingVariables = $ScriptBlock.ast.FindAll({ $args[0] -is [System.Management.Automation.Language.UsingExpressionAst] }, $True)
						
						If ($UsingVariables)
						{
							$List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
							ForEach ($Ast in $UsingVariables)
							{
								[void]$list.Add($Ast.SubExpression)
							}
							
							$UsingVar = $UsingVariables | Group Parent | ForEach { $_.Group | Select -First 1 }
							
							#Extract the name, value, and create replacements for each
							$UsingVariableData = ForEach ($Var in $UsingVar)
							{
								Try
								{
									$Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
									$NewName = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									[pscustomobject]@{
										Name = $Var.SubExpression.Extent.Text
										Value = $Value.Value
										NewName = $NewName
										NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
									}
									$ParamsToAdd += $NewName
								}
								Catch
								{
									Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
								}
							}
							
							$NewParams = $UsingVariableData.NewName -join ', '
							$Tuple = [Tuple]::Create($list, $NewParams)
							$bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
							$GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl', $bindingFlags))
							
							$StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast, @($Tuple))
							
							$ScriptBlock = [scriptblock]::Create($StringScriptBlock)
							
							Write-Verbose $StringScriptBlock
						}
					}
					
					$ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
				}
				else
				{
					Throw "Must provide ScriptBlock or ScriptFile"; Break
				}
				
				Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
				Write-Verbose "Creating runspace pool and session states"
				
				#If specified, add variables and modules/snapins to session state
				$sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
				if ($ImportVariables)
				{
					if ($UserVariables.count -gt 0)
					{
						foreach ($Variable in $UserVariables)
						{
							$sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null))
						}
					}
				}
				if ($ImportModules)
				{
					if ($UserModules.count -gt 0)
					{
						foreach ($ModulePath in $UserModules)
						{
							$sessionstate.ImportPSModule($ModulePath)
						}
					}
					if ($UserSnapins.count -gt 0)
					{
						foreach ($PSSnapin in $UserSnapins)
						{
							[void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
						}
					}
				}
				
				#Create runspace pool
				$runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
				$runspacepool.Open()
				
				Write-Verbose "Creating empty collection to hold runspace jobs"
				$Script:runspaces = New-Object System.Collections.ArrayList
				
				#If inputObject is bound get a total count and set bound to true
				$global:__bound = $false
				$allObjects = @()
				if ($PSBoundParameters.ContainsKey("inputObject"))
				{
					$global:__bound = $true
				}
				
				#Set up log file if specified
				if ($LogFile)
				{
					New-Item -ItemType file -path $logFile -force | Out-Null
					("" | Select Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
				}
				
				#write initial log entry
				$log = "" | Select Date, Action, Runtime, Status, Details
				$log.Date = Get-Date
				$log.Action = "Batch processing started"
				$log.Runtime = $null
				$log.Status = "Started"
				$log.Details = $null
				if ($logFile)
				{
					($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
				}
				
				$timedOutTasks = $false
				
				#endregion INIT
			}
			
			Process
			{
				
				#add piped objects to all objects or set all objects to bound input object parameter
				if (-not $global:__bound)
				{
					$allObjects += $inputObject
				}
				else
				{
					$allObjects = $InputObject
				}
			}
			
			End
			{
				
				#Use Try/Finally to catch Ctrl+C and clean up.
				Try
				{
					#counts for progress
					$totalCount = $allObjects.count
					$script:completedCount = 0
					$startedCount = 0
					
					foreach ($object in $allObjects)
					{
						
						#region add scripts to runspace pool
						
						#Create the powershell instance, set verbose if needed, supply the scriptblock and parameters
						$powershell = [powershell]::Create()
						
						if ($VerbosePreference -eq 'Continue')
						{
							[void]$PowerShell.AddScript({ $VerbosePreference = 'Continue' })
						}
						
						[void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)
						
						if ($parameter)
						{
							[void]$PowerShell.AddArgument($parameter)
						}
						
						# $Using support from Boe Prox
						if ($UsingVariableData)
						{
							Foreach ($UsingVariable in $UsingVariableData)
							{
								Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
								[void]$PowerShell.AddArgument($UsingVariable.Value)
							}
						}
						
						#Add the runspace into the powershell instance
						$powershell.RunspacePool = $runspacepool
						
						#Create a temporary collection for each runspace
						$temp = "" | Select-Object PowerShell, StartTime, object, Runspace
						$temp.PowerShell = $powershell
						$temp.StartTime = Get-Date
						$temp.object = $object
						
						#Save the handle output when calling BeginInvoke() that will be used later to end the runspace
						$temp.Runspace = $powershell.BeginInvoke()
						$startedCount++
						
						#Add the temp tracking info to $runspaces collection
						Write-Verbose ("Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring())
						$runspaces.Add($temp) | Out-Null
						
						#loop through existing runspaces one time
						Get-RunspaceData
						
						#If we have more running than max queue (used to control timeout accuracy)
						#Script scope resolves odd PowerShell 2 issue
						$firstRun = $true
						while ($runspaces.count -ge $Script:MaxQueue)
						{
							
							#give verbose output
							if ($firstRun)
							{
								Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
							}
							$firstRun = $false
							
							#run get-runspace data and sleep for a short while
							Get-RunspaceData
							Start-Sleep -Milliseconds $sleepTimer
							
						}
						
						#endregion add scripts to runspace pool
					}
					
					Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f (@($runspaces | Where { $_.Runspace -ne $Null }).Count))
					Get-RunspaceData -wait
					
					if (-not $quiet)
					{
						Write-Progress -Activity "Running Query" -Status "Starting threads" -Completed
					}
					
				}
				Finally
				{
					#Close the runspace pool, unless we specified no close on timeout and something timed out
					if (($timedOutTasks -eq $false) -or (($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false)))
					{
						Write-Verbose "Closing the runspace pool"
						$runspacepool.close()
					}
					
					#collect garbage
					[gc]::Collect()
				}
			}
		}
		
		Write-Verbose "PSBoundParameters = $($PSBoundParameters | Out-String)"
		
		$bound = $PSBoundParameters.keys -contains "ComputerName"
		if (-not $bound)
		{
			[System.Collections.ArrayList]$AllComputers = @()
		}
	}
	Process
	{
		
		#Handle both pipeline and bound parameter.  We don't want to stream objects, defeats purpose of parallelizing work
		if ($bound)
		{
			$AllComputers = $ComputerName
		}
		Else
		{
			foreach ($Computer in $ComputerName)
			{
				$AllComputers.add($Computer) | Out-Null
			}
		}
		
	}
	End
	{
		
		#Built up the parameters and run everything in parallel
		$params = @($Detail, $Quiet)
		$splat = @{
			Throttle = $Throttle
			RunspaceTimeout = $Timeout
			InputObject = $AllComputers
			parameter = $params
		}
		if ($NoCloseOnTimeout)
		{
			$splat.add('NoCloseOnTimeout', $True)
		}
		
		Invoke-Parallel @splat -ScriptBlock {
			
			$computer = $_.trim()
			$detail = $parameter[0]
			$quiet = $parameter[1]
			
			#They want detail, define and run test-server
			if ($detail)
			{
				Try
				{
					#Modification of jrich's Test-Server function: https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a
					Function Test-Server
					{
						[cmdletBinding()]
						param (
							[parameter(
									   Mandatory = $true,
									   ValueFromPipeline = $true)]
							[string[]]$ComputerName,
							
							[switch]$All,
							
							[parameter(Mandatory = $false)]
							[switch]$CredSSP,
							
							[switch]$RemoteReg,
							
							[switch]$RDP,
							
							[switch]$RPC,
							
							[switch]$SMB,
							
							[switch]$WSMAN,
							
							[switch]$IPV6,
							
							[Management.Automation.PSCredential]$Credential
						)
						begin
						{
							$total = Get-Date
							$results = @()
							if ($credssp -and -not $Credential)
							{
								Throw "Must supply Credentials with CredSSP test"
							}
							
							[string[]]$props = write-output Name, IP, Domain, Ping, WSMAN, CredSSP, RemoteReg, RPC, RDP, SMB
							
							#Hash table to create PSObjects later, compatible with ps2...
							$Hash = @{ }
							foreach ($prop in $props)
							{
								$Hash.Add($prop, $null)
							}
							
							function Test-Port
							{
								[cmdletbinding()]
								Param (
									[string]$srv,
									
									$port = 135,
									
									$timeout = 3000
								)
								$ErrorActionPreference = "SilentlyContinue"
								$tcpclient = new-Object system.Net.Sockets.TcpClient
								$iar = $tcpclient.BeginConnect($srv, $port, $null, $null)
								$wait = $iar.AsyncWaitHandle.WaitOne($timeout, $false)
								if (-not $wait)
								{
									$tcpclient.Close()
									Write-Verbose "Connection Timeout to $srv`:$port"
									$false
								}
								else
								{
									Try
									{
										$tcpclient.EndConnect($iar) | out-Null
										$true
									}
									Catch
									{
										write-verbose "Error for $srv`:$port`: $_"
										$false
									}
									$tcpclient.Close()
								}
							}
						}
						
						process
						{
							foreach ($name in $computername)
							{
								$dt = $cdt = Get-Date
								Write-verbose "Testing: $Name"
								$failed = 0
								try
								{
									$DNSEntity = [Net.Dns]::GetHostEntry($name)
									$domain = ($DNSEntity.hostname).replace("$name.", "")
									$ips = $DNSEntity.AddressList | %{
										if (-not (-not $IPV6 -and $_.AddressFamily -like "InterNetworkV6"))
										{
											$_.IPAddressToString
										}
									}
								}
								catch
								{
									$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
									$rst.name = $name
									$results += $rst
									$failed = 1
								}
								Write-verbose "DNS:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
								if ($failed -eq 0)
								{
									foreach ($ip in $ips)
									{
										
										$rst = New-Object -TypeName PSObject -Property $Hash | Select -Property $props
										$rst.name = $name
										$rst.ip = $ip
										$rst.domain = $domain
										
										if ($RDP -or $All)
										{
											####RDP Check (firewall may block rest so do before ping
											try
											{
												$socket = New-Object Net.Sockets.TcpClient($name, 3389) -ErrorAction stop
												if ($socket -eq $null)
												{
													$rst.RDP = $false
												}
												else
												{
													$rst.RDP = $true
													$socket.close()
												}
											}
											catch
											{
												$rst.RDP = $false
												Write-Verbose "Error testing RDP: $_"
											}
										}
										Write-verbose "RDP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
										#########ping
										if (test-connection $ip -count 2 -Quiet)
										{
											Write-verbose "PING:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
											$rst.ping = $true
											
											if ($WSMAN -or $All)
											{
												try
												{
													############wsman
														Test-WSMan $ip -ErrorAction stop | Out-Null
														$rst.WSMAN = $true
													}
													catch
													{
														$rst.WSMAN = $false
														Write-Verbose "Error testing WSMAN: $_"
													}
													Write-verbose "WSMAN:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													if ($rst.WSMAN -and $credssp) ########### credssp
													{
														try
														{
															Test-WSMan $ip -Authentication Credssp -Credential $cred -ErrorAction stop
															$rst.CredSSP = $true
														}
														catch
														{
															$rst.CredSSP = $false
															Write-Verbose "Error testing CredSSP: $_"
														}
														Write-verbose "CredSSP:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													}
												}
												if ($RemoteReg -or $All)
												{
													try ########remote reg
													{
														[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ip) | Out-Null
														$rst.remotereg = $true
													}
													catch
													{
														$rst.remotereg = $false
														Write-Verbose "Error testing RemoteRegistry: $_"
													}
													Write-verbose "remote reg:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($RPC -or $All)
												{
													try ######### wmi
													{
														$w = [wmi] ''
														$w.psbase.options.timeout = 15000000
														$w.path = "\\$Name\root\cimv2:Win32_ComputerSystem.Name='$Name'"
														$w | select none | Out-Null
														$rst.RPC = $true
													}
													catch
													{
														$rst.rpc = $false
														Write-Verbose "Error testing WMI/RPC: $_"
													}
													Write-verbose "WMI/RPC:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
												}
												if ($SMB -or $All)
												{
													
													#Use set location and resulting errors.  push and pop current location
													try ######### C$
													{
														$path = "\\$name\c$"
														Push-Location -Path $path -ErrorAction stop
														$rst.SMB = $true
														Pop-Location
													}
													catch
													{
														$rst.SMB = $false
														Write-Verbose "Error testing SMB: $_"
													}
													Write-verbose "SMB:  $((New-TimeSpan $dt ($dt = get-date)).totalseconds)"
													
												}
											}
											else
											{
												$rst.ping = $false
												$rst.wsman = $false
												$rst.credssp = $false
												$rst.remotereg = $false
												$rst.rpc = $false
												$rst.smb = $false
											}
											$results += $rst
										}
									}
									Write-Verbose "Time for $($Name): $((New-TimeSpan $cdt ($dt)).totalseconds)"
									Write-Verbose "----------------------------"
								}
							}
							end
							{
								Write-Verbose "Time for all: $((New-TimeSpan $total ($dt)).totalseconds)"
								Write-Verbose "----------------------------"
								return $results
							}
						}
						
						#Build up parameters for Test-Server and run it
						$TestServerParams = @{
							ComputerName = $Computer
							ErrorAction = "Stop"
						}
						
						if ($detail -eq "*")
						{
							$detail = "WSMan", "RemoteReg", "RPC", "RDP", "SMB"
						}
						
						$detail | Select -Unique | Foreach-Object { $TestServerParams.add($_, $True) }
						Test-Server @TestServerParams | Select -Property $("Name", "IP", "Domain", "Ping" + $detail)
					}
					Catch
					{
						Write-Warning "Error with Test-Server: $_"
					}
				}
				#We just want ping output
				else
				{
					Try
					{
						#Pick out a few properties, add a status label.  If quiet output, just return the address
						$result = $null
						if ($result = @(Test-Connection -ComputerName $computer -Count 2 -erroraction Stop))
						{
							$Output = $result | Select -first 1 -Property Address,
													   IPV4Address,
													   IPV6Address,
													   ResponseTime,
													   @{ label = "STATUS"; expression = { "Responding" } }
							
							if ($quiet)
							{
								$Output.address
							}
							else
							{
								$Output
							}
						}
					}
					Catch
					{
						if (-not $quiet)
						{
							#Ping failed.  I'm likely making inappropriate assumptions here, let me know if this is the case : )
							if ($_ -match "No such host is known")
							{
								$status = "Unknown host"
							}
							elseif ($_ -match "Error due to lack of resources")
							{
								$status = "No Response"
							}
							else
							{
								$status = "Error: $_"
							}
							
							"" | Select -Property @{ label = "Address"; expression = { $computer } },
										IPV4Address,
										IPV6Address,
										ResponseTime,
										@{ label = "STATUS"; expression = { $status } }
						}
					}
				}
			}
		}
	}
# End Invoke-Ping
# Begin New-ISOFile
function New-IsoFile  
{  
  <#  
   .Synopsis  
    Creates a new .iso file  
   .Description  
    The New-IsoFile cmdlet creates a new .iso file containing content from chosen folders  
   .Example  
    New-IsoFile "c:\tools","c:Downloads\utils"  
    This command creates a .iso file in $env:temp folder (default location) that contains c:\tools and c:\downloads\utils folders. The folders themselves are included at the root of the .iso image.  
   .Example 
    New-IsoFile -FromClipboard -Verbose 
    Before running this command, select and copy (Ctrl-C) files/folders in Explorer first.  
   .Example  
    dir c:\WinPE | New-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE" 
    This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx  
   .Notes 
    NAME:  New-IsoFile  
    AUTHOR: Chris Wu 
    LASTEDIT: 03/23/2016 14:46:50  
 #>  
  
  [CmdletBinding(DefaultParameterSetName='Source')]Param( 
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]$Source,  
    [parameter(Position=2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch]$Force, 
    [parameter(ParameterSetName='Clipboard')][switch]$FromClipboard 
  ) 
 
  Begin {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @' 
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
  
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
  
    if ($BootFile) { 
      if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
      $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
    } 
 
    $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
 
    Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
  
    if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
  }  
 
  Process { 
    if($FromClipboard) { 
      if($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
      $Source = Get-Clipboard -Format FileDropList 
    } 
 
    foreach($item in $Source) { 
      if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
        $item = Get-Item -LiteralPath $item 
      } 
 
      if($item) { 
        Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
        try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
      } 
    } 
  } 
 
  End {  
    if ($Boot) { $Image.BootImageOptions=$Boot }  
    $Result = $Image.CreateResultImage()  
    [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
    Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
    $Target 
  } 
} 
# End New-IsoFile
# Begin Set-FileTime
function Set-FileTime{
  param(
    [string[]]$paths,
    [bool]$only_modification = $false,
    [bool]$only_access = $false
  );

  begin {
    function updateFileSystemInfo([System.IO.FileSystemInfo]$fsInfo) {
      $datetime = get-date
      if ( $only_access )
      {
         $fsInfo.LastAccessTime = $datetime
      }
      elseif ( $only_modification )
      {
         $fsInfo.LastWriteTime = $datetime
      }
      else
      {
         $fsInfo.CreationTime = $datetime
         $fsInfo.LastWriteTime = $datetime
         $fsInfo.LastAccessTime = $datetime
       }
    }
   
    function touchExistingFile($arg) {
      if ($arg -is [System.IO.FileSystemInfo]) {
        updateFileSystemInfo($arg)
      }
      else {
        $resolvedPaths = resolve-path $arg
        foreach ($rpath in $resolvedPaths) {
          if (test-path -type Container $rpath) {
            $fsInfo = new-object System.IO.DirectoryInfo($rpath)
          }
          else {
            $fsInfo = new-object System.IO.FileInfo($rpath)
          }
          updateFileSystemInfo($fsInfo)
        }
      }
    }
   
    function touchNewFile([string]$path) {
      #$null > $path
      Set-Content -Path $path -value $null;
    }
  }
 
  process {
    if ($_) {
      if (test-path $_) {
        touchExistingFile($_)
      }
      else {
        touchNewFile($_)
      }
    }
  }
 
  end {
    if ($paths) {
      foreach ($path in $paths) {
        if (test-path $path) {
          touchExistingFile($path)
        }
        else {
          touchNewFile($path)
        }
      }
    }
  }
}
# End Set-FileTime
# Begin Get-PendingUpdate
Function Get-PendingUpdate { 
    <#    
      .SYNOPSIS   
        Retrieves the updates waiting to be installed from WSUS   
      .DESCRIPTION   
        Retrieves the updates waiting to be installed from WSUS  
      .PARAMETER Computername 
        Computer or computers to find updates for.   
      .EXAMPLE   
       Get-PendingUpdates 
    
       Description 
       ----------- 
       Retrieves the updates that are available to install on the local system 
      .NOTES 
      Author: Boe Prox                                           
                                        
    #> 
      
    #Requires -version 3.0   
    [CmdletBinding( 
        DefaultParameterSetName = 'computer' 
        )] 
    param( 
        [Parameter(ValueFromPipeline = $True)] 
            [string[]]$Computername = $env:COMPUTERNAME
        )     
    Process { 
        ForEach ($computer in $Computername) { 
            If (Test-Connection -ComputerName $computer -Count 1 -Quiet) { 
                Try { 
                #Create Session COM object 
                    Write-Verbose "Creating COM object for WSUS Session" 
                    $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$computer)) 
                    } 
                Catch { 
                    Write-Warning "$($Error[0])" 
                    Break 
                    } 
 
                #Configure Session COM Object 
                Write-Verbose "Creating COM object for WSUS update Search" 
                $updatesearcher = $updatesession.CreateUpdateSearcher() 
 
                #Configure Searcher object to look for Updates awaiting installation 
                Write-Verbose "Searching for WSUS updates on client" 
                $searchresult = $updatesearcher.Search("IsInstalled=0")     
             
                #Verify if Updates need installed 
                Write-Verbose "Verifing that updates are available to install" 
                If ($searchresult.Updates.Count -gt 0) { 
                    #Updates are waiting to be installed 
                    Write-Verbose "Found $($searchresult.Updates.Count) update\s!" 
                    #Cache the count to make the For loop run faster 
                    $count = $searchresult.Updates.Count 
                 
                    #Begin iterating through Updates available for installation 
                    Write-Verbose "Iterating through list of updates" 
                    For ($i=0; $i -lt $Count; $i++) { 
                        #Create object holding update 
                        $Update = $searchresult.Updates.Item($i)
                        [pscustomobject]@{
                            Computername = $Computer
                            Title = $Update.Title
                            KB = $($Update.KBArticleIDs)
                            SecurityBulletin = $($Update.SecurityBulletinIDs)
                            MsrcSeverity = $Update.MsrcSeverity
                            IsDownloaded = $Update.IsDownloaded
                            Url = $($Update.MoreInfoUrls)
                            Categories = ($Update.Categories | Select-Object -ExpandProperty Name)
                            BundledUpdates = @($Update.BundledUpdates)|ForEach{
                               [pscustomobject]@{
                                    Title = $_.Title
                                    DownloadUrl = @($_.DownloadContents).DownloadUrl
                                }
                            }
                        } 
                    }
                } 
                Else { 
                    #Nothing to install at this time 
                    Write-Verbose "No updates to install." 
                }
            } 
            Else { 
                #Nothing to install at this time 
                Write-Warning "$($c): Offline" 
            }  
        }
    }  
}
# End Get-PendingUpdate
# Begin Get-Set-NetworkLevelAuthentication
function Get-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This function will get the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This function will get the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computer to query

	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.
	
	.EXAMPLE
		Get-NetworkLevelAuthentication
		
		This will get the NLA setting on the localhost
	
		ComputerName     : XAVIERDESKTOP
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp	

    .EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01
		
		This will get the NLA setting on the server DC01
	
		ComputerName     : DC01
		NLAEnabled       : True
		TerminalName     : RDP-Tcp
		TerminalProtocol : Microsoft RDP 8.0
		Transport        : tcp
	
	.EXAMPLE
		Get-NetworkLevelAuthentication -ComputerName DC01, SERVER01 -verbose
	
	.EXAMPLE
		Get-Content .\Computers.txt | Get-NetworkLevelAuthentication -verbose
		
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				[pscustomobject][ordered]@{
					'ComputerName' = $NLAinfo.PSComputerName
					'NLAEnabled' = $NLAinfo.UserAuthenticationRequired -as [bool]
					'TerminalName' = $NLAinfo.TerminalName
					'TerminalProtocol' = $NLAinfo.TerminalProtocol
					'Transport' = $NLAinfo.transport
				}
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{
		
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}


function Set-NetworkLevelAuthentication
{
<#
	.SYNOPSIS
		This function will set the NLA setting on a local machine or remote machine

	.DESCRIPTION
		This function will set the NLA setting on a local machine or remote machine

	.PARAMETER  ComputerName
		Specify one or more computers
	
	.PARAMETER EnableNLA
		Specify if the NetworkLevelAuthentication need to be set to $true or $false
	
	.PARAMETER  Credential
		Specify the alternative credential to use. By default it will use the current one.

	.EXAMPLE
		Set-NetworkLevelAuthentication -EnableNLA $true

		ReturnValue                             PSComputerName                         
		-----------                             --------------                         
		                                        XAVIERDESKTOP      
	
	.NOTES
		DATE	: 2014/04/01
		AUTHOR	: Francois-Xavier Cat
		WWW		: http://lazywinadmin.com
		Twitter	: @lazywinadm
#>
	#Requires -Version 3.0
	[CmdletBinding()]
	PARAM (
		[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String[]]$ComputerName = $env:ComputerName,
		
		[Parameter(Mandatory)]
		[Bool]$EnableNLA,
		
		[Alias("RunAs")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)#Param
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name CimCmdlets))
			{
				Write-Verbose -Message 'BEGIN - Import Module CimCmdlets'
				Import-Module CimCmdlets -ErrorAction 'Stop' -ErrorVariable ErrorBeginCimCmdlets
				
			}
		}
		CATCH
		{
			IF ($ErrorBeginCimCmdlets)
			{
				Write-Error -Message "BEGIN - Can't find CimCmdlets Module"
			}
		}
	}#BEGIN
	
	PROCESS
	{
		FOREACH ($Computer in $ComputerName)
		{
			TRY
			{
				# Building Splatting for CIM Sessions
				$CIMSessionParams = @{
					ComputerName = $Computer
					ErrorAction = 'Stop'
					ErrorVariable = 'ProcessError'
				}
				
				# Add Credential if specified when calling the function
				IF ($PSBoundParameters['Credential'])
				{
					$CIMSessionParams.credential = $Credential
				}
				
				# Connectivity Test
				Write-Verbose -Message "PROCESS - $Computer - Testing Connection..."
				Test-Connection -ComputerName $Computer -count 1 -ErrorAction Stop -ErrorVariable ErrorTestConnection | Out-Null
				
				# CIM/WMI Connection
				#  WsMAN
				IF ((Test-WSMan -ComputerName $Computer -ErrorAction SilentlyContinue).productversion -match 'Stack: 3.0')
				{
					Write-Verbose -Message "PROCESS - $Computer - WSMAN is responsive"
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# DCOM
				ELSE
				{
					# Trying with DCOM protocol
					Write-Verbose -Message "PROCESS - $Computer - Trying to connect via DCOM protocol"
					$CIMSessionParams.SessionOption = New-CimSessionOption -Protocol Dcom
					$CimSession = New-CimSession @CIMSessionParams
					$CimProtocol = $CimSession.protocol
					Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Opened"
				}
				
				# Getting the Information on Terminal Settings
				Write-Verbose -message "PROCESS - $Computer - [$CimProtocol] CIM SESSION - Get the Terminal Services Information"
				$NLAinfo = Get-CimInstance -CimSession $CimSession -ClassName Win32_TSGeneralSetting -Namespace root\cimv2\terminalservices -Filter "TerminalName='RDP-tcp'"
				$NLAinfo | Invoke-CimMethod -MethodName SetUserAuthenticationRequired -Arguments @{ UserAuthenticationRequired = $EnableNLA } -ErrorAction 'Continue' -ErrorVariable ErrorProcessInvokeWmiMethod
			}
			
			CATCH
			{
				Write-Warning -Message "PROCESS - Error on $Computer"
				$_.Exception.Message
				if ($ErrorTestConnection) { Write-Warning -Message "PROCESS Error - $ErrorTestConnection" }
				if ($ProcessError) { Write-Warning -Message "PROCESS Error - $ProcessError" }
				if ($ErrorProcessInvokeWmiMethod) { Write-Warning -Message "PROCESS Error - $ErrorProcessInvokeWmiMethod" }
			}#CATCH
		} # FOREACH
	}#PROCESS
	END
	{	
		if ($CimSession)
		{
			Write-Verbose -Message "END - Close CIM Session(s)"
			Remove-CimSession $CimSession
		}
		Write-Verbose -Message "END - Script is completed"
	}
}
# End Get-Set-NetworkLevelAuthentication
# Begin Get-FolderSize
Function Get-FolderSize 
{

<#
.SYNOPSIS
	Get-FolderSize Function displays size of all folders in a specified path. 
	
.DESCRIPTION

	Get-FolderSize Function allows you to return folders greater than a specified size.
	See examples below for more info. 
	
.PARAMETER FolderPath
	Specifies the path you wish to check folder sizes. For example \\70411SRV\EventLogs
	Will return sizes (in GB) of all folders in \\70411SRV\EventLogs. FolderPath accepts
	both UNC and local path format. You can specify multiple paths in quotes, seperated
	by commas. 
	
.PARAMETER FoldersOver
	This parameter is specified in whole numbers (but represents values in GB). It instructs
	the Get-FolderSize function to return only folders greater than or equal to the specified
	value in GB. 
	
.PARAMETER Recurse
	If this parameter is specified, size of all folders and subfolders are displayed 
	If the Recurse parameter is not spefified (default), size of base folders are displayed.
	
.EXAMPLE
	To return size for all folders in C:\EventLogs, run the command:
	PS C:\>Get-FolderSize -FolderPath C:\EventLogs
	The command returns the following output:
	Permorning initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...
	
	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	70411SRV1                               C:\EventLogs\70411SRV1                  128 MB
	70411SRV2                               C:\EventLogs\70411SRV2                  128 MB
	70411SRV3                               C:\EventLogs\70411SRV3                  128 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB

.EXAMPLE
	To return size for folders in C:\EventLogs greater or equal to 200BM, run the command:
	PS C:\> Get-FolderSize C:\EventLogs -FoldersOver 0.2
	Result of the above command is shown below:
	Permorning initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...

	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB
	Notice that only folders greater than 200 MB were returned 
	
.EXAMPLE
	To return size of all folders and subfolders, specify the Recurse parameter:
	PS C:\> Get-FolderSize C:\EventLogs -Recurse
	
	Performing initial tasks, please wait...
	Calculating size of folders in C:\EventLogs. This may take sometime, please wait...

	Folder Name                             Full Path                               Size
	-----------                             ---------                               ----
	70411SRV                                C:\EventLogs\70411SRV                   384 MB
	70411SRV1                               C:\EventLogs\70411SRV1                  128 MB
	70411SRV2                               C:\EventLogs\70411SRV2                  128 MB
	70411SRV3                               C:\EventLogs\70411SRV3                  128 MB
	Softwares                               C:\EventLogs\Softwares                  2.34 GB
	Acrobat 7                               C:\EventLogs\Softwares\Acrobat 7        209 MB
	Citrix                                  C:\EventLogs\Softwares\Citrix           96 MB
	Dell OM station                         C:\EventLogs\Softwares\Dell OM station  577 MB
	JAWS                                    C:\EventLogs\Softwares\JAWS             227 MB
	MDT 2012 Update1                        C:\EventLogs\Softwares\MDT 2012 Update1 118 MB
	OpenManageEssentials                    C:\EventLogs\Softwares\OpenManageEss... 891 MB
	Adobe Acrobat 7.0 Professional          C:\EventLogs\Softwares\Acrobat 7\Ado... 200 MB
	windows                                 C:\EventLogs\Softwares\Dell OM stati... 271 MB
	ManagementStation                       C:\EventLogs\Softwares\Dell OM stati... 254 MB
	support                                 C:\EventLogs\Softwares\Dell OM stati... 107 MB
	
	Notice that, we now have size of all folders and subfolders
	
	
#>

[CmdletBinding(DefaultParameterSetName='FolderPath')]
param 
(
[Parameter(Mandatory=$true,Position=0,ParameterSetName='FolderPath')]
[String[]]$FolderPath,
[Parameter(Mandatory=$false,Position=1,ParameterSetName='FolderPath')]
[String]$FoldersOver,
[Parameter(Mandatory=$false,Position=2,ParameterSetName='FolderPath')]
[switch]$Recurse

)

Begin 
{
#$FoldersOver and $ZeroSizeFolders cannot be used together
#Convert the size specified by Greaterhan parameter to Bytes
$size = 1000000000 * $FoldersOver

}

Process {#Check whether user has access to the folders.
	
	
		Try {
		Write-Host "Performing initial tasks, please wait... " -ForegroundColor Magenta
		$ColItems = If ($Recurse) {Get-ChildItem $FolderPath -Recurse -ErrorAction Stop } 
		Else {Get-ChildItem $FolderPath -ErrorAction Stop } 
		
		} 
		Catch [exception]{}
		
		#Calculate folder size
		If ($ColItems) 
		{
		Write-Host "Calculating size of folders in $FolderPath. This may take sometime, please wait... " -ForegroundColor Magenta
		$Items = $ColItems | Where-Object {$_.PSIsContainer -eq $TRUE -and `
		@(Get-ChildItem -LiteralPath $_.Fullname -Recurse -ErrorAction SilentlyContinue | Where-Object {!$_.PSIsContainer}).Length -gt '0'}}
		

		ForEach ($i in $Items)
		{

		$subFolders = 
		If ($FoldersOver)
		{Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object {$_.Sum -ge $size -and $_.Sum -gt 100000000  } }
		Else
		{Get-ChildItem -Path $i.FullName -Recurse | Measure-Object -sum Length | Where-Object {$_.Sum -gt 100000000  }} #added 25/12/2014: returns folders over 100MB
		#Return only values not equal to 0
		ForEach ($subFolder in $subFolders) {
		#If folder is less than or equal to 1GB, display in MB, If above 1GB, display in GB 
		$si = If (($subFolder.Sum -ge 1000000000)  ) {"{0:N2}" -f ($subFolder.Sum / 1GB) + " GB"} 
 	  	ElseIf (($subFolder.Sum -lt 1000000000)  ) {"{0:N0}" -f ($subFolder.Sum / 1MB) + " MB"} 
		$Object = New-Object PSObject -Property @{            
        'Folder Name'    = $i.Name                
        'Size'    =  $si
        'Full Path'    = $i.FullName          
        } 

		$Object | Select-Object 'Folder Name', 'Full Path',Size



} 

}


}
End {

Write-Host "Task completed...if nothing is displayed:
you may not have access to the path specified or 
all folders are less than 100 MB" -ForegroundColor Cyan


}

}
# End Get-FolderSize

# Begin Get-Software
Function Get-Software  {

  [OutputType('System.Software.Inventory')]

  [Cmdletbinding()] 

  Param( 

  [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 

  [String[]]$Computername=$env:COMPUTERNAME

  )         

  Begin {

  }

  Process  {     

  ForEach  ($Computer in  $Computername){ 

  If  (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {

  $Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         

  ForEach($Path in $Paths) { 

  Write-Verbose  "Checking Path: $Path"

  #  Create an instance of the Registry Object and open the HKLM base key 

  Try  { 

  $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 

  } Catch  { 

  Write-Error $_ 

  Continue 

  } 

  #  Drill down into the Uninstall key using the OpenSubKey Method 

  Try  {

  $regkey=$reg.OpenSubKey($Path)  

  # Retrieve an array of string that contain all the subkey names 

  $subkeys=$regkey.GetSubKeyNames()      

  # Open each Subkey and use GetValue Method to return the required  values for each 

  ForEach ($key in $subkeys){   

  Write-Verbose "Key: $Key"

  $thisKey=$Path+"\\"+$key 

  Try {  

  $thisSubKey=$reg.OpenSubKey($thisKey)   

  # Prevent Objects with empty DisplayName 

  $DisplayName =  $thisSubKey.getValue("DisplayName")

  If ($DisplayName  -AND $DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {

  $Date = $thisSubKey.GetValue('InstallDate')

  If ($Date) {

  Try {

  $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)

  } Catch{

  Write-Warning "$($Computer): $_ <$($Date)>"

  $Date = $Null

  }

  } 

  # Create New Object with empty Properties 

  $Publisher =  Try {

  $thisSubKey.GetValue('Publisher').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('Publisher')

  }

  $Version = Try {

  #Some weirdness with trailing [char]0 on some strings

  $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))

  } 

  Catch {

  $thisSubKey.GetValue('DisplayVersion')

  }

  $UninstallString =  Try {

  $thisSubKey.GetValue('UninstallString').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('UninstallString')

  }

  $InstallLocation =  Try {

  $thisSubKey.GetValue('InstallLocation').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('InstallLocation')

  }

  $InstallSource =  Try {

  $thisSubKey.GetValue('InstallSource').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('InstallSource')

  }

  $HelpLink = Try {

  $thisSubKey.GetValue('HelpLink').Trim()

  } 

  Catch {

  $thisSubKey.GetValue('HelpLink')

  }

  $Object = [pscustomobject]@{

  Computername = $Computer

  DisplayName = $DisplayName

  Version  = $Version

  InstallDate = $Date

  Publisher = $Publisher

  UninstallString = $UninstallString

  InstallLocation = $InstallLocation

  InstallSource  = $InstallSource

  HelpLink = $thisSubKey.GetValue('HelpLink')

  EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))

  }

  $Object.pstypenames.insert(0,'System.Software.Inventory')

  Write-Output $Object

  }

  } Catch {

  Write-Warning "$Key : $_"

  }   

  }

  } Catch  {}   

  $reg.Close() 

  }                  

  } Else  {

  Write-Error  "$($Computer): unable to reach remote system!"

  }

  } 

  } 

}  
# End Get-Software
# Begin Get-AssetTagAndSerialNumber
function Get-AssetTagAndSerialNumber {

   param  ( [string[]]$computerName = @('.') );

   $computerName | % {

       if ($_) {

           Get-WmiObject -ComputerName $_ Win32_SystemEnclosure | Select-Object __Server, SerialNumber, SMBiosAssetTag

       }

   }

}
# End Get-AssetTagAndSerialNumber
# Begin Enable-MemHotAdd
Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Enable-MemHotAdd
# Begin Disable-MemHotAdd
Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Disable-MemHotAdd
# Begin Enable-vCPUHotAdd
Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Enable-vCPUHotAdd
# Begin Disable-vCPUHotAdd
Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
# End Disable-vCPUHotAdd