<#
	Created by Jose Gabriel Ortega C.
	All right reserved
	mail: j0rt3g4@j0rt3g4.com
    web: https://www.j0rt3g4.com
	Version 1.0
		* Initial Version, add/remove domain suffixes to active directory from a csv file
			
    ARdomainSuffix.ps1
    version 2.0 (03/26/2015)
		* Check if AD Module Exists
			*If it exists check IF it's loaded	
				If not loaded, it will load it
				If it's loaded, will skip the loading.
			If it doesn't exists will exit the program
		* Add or Remove domain suffixes to domain using  2 parameters: NewOldDomain and a Action
			*Parameter NewOldDomain: It's a string specified by the user, Mandatory and must match a Regex for domain names in http://stackoverflow.com/questions/10306690/domain-name-validation-with-regex, it's in the 1st position
			*Parameter Action: It's a action to do with the domain: Add, add, Remove or remove are valid options, if you write anything else more than those will get a warning about it. 
    Version 3.0 (04/11/2016)

        * Added Several Validations in the process
        * Clearer "clean up" step.

	Version 4.0 (09/14/2016)
	
		* Corrected bug when there's nothing in null the actual domains in the "AD Domain and trust", checked if that variable is null or empty.
		* Added correct dates into the logs with US format: mm/dd/yyyy
		
		
    Examples:
        * To Add a domain:
            .\ARdomainSuffix.ps1 "DomainToBeAdded.com" -Action Add

        * To Remove a domain:
            .\ARdomainSuffix.ps1 "DomainToBeRemoved.com" -Action remove


    Version 5.0 (03/10/2017)
        * Added Get Option
    	* Added new ways to run the script:
    
    Previous way still compatibles:
    Examples:
        * To Add a domain:
            .\ARdomainSuffix.ps1 "DomainToBeAdded.com" -Action Add

        * To Remove a domain:
            .\ARdomainSuffix.ps1 "DomainToBeRemoved.com" -Action remove

    New way: (Allows you to see if you have or not any domain allowed in your active directory)
        * To Remove a domain:
            .\ARdomainSuffix.ps1 -get

#>

#####################################################################################################################################
#####                                              VARIABLES AND FUNCTIONS                                                      #####
#####################################################################################################################################
#[CmdletBinding(DefaultParameterSetName="default")]
param(
	#Regex Source: http://stackoverflow.com/questions/10306690/domain-name-validation-with-regex
    [Parameter(ParametersetName="default",ValueFromPipelineByPropertyName,Mandatory=$true,position=0)][ValidateNotNullOrEmpty()] [ValidatePattern('^(([a-zA-Z]{1})|([a-zA-Z]{1}[a-zA-Z]{1})|([a-zA-Z]{1}[0-9]{1})|([0-9]{1}[a-zA-Z]{1})|([a-zA-Z0-9][a-zA-Z0-9-_]{1,61}[a-zA-Z0-9]))\.([a-zA-Z]{2,6}|[a-zA-Z0-9-]{2,30}\.[a-zA-Z]{2,3})$')] [string]$NewOldDomain,
    [Parameter(ParametersetName="default",ValueFromPipelineByPropertyName,Mandatory=$true,position=1)][Parameter(ParameterSetName="other",position=0)][ValidateSet("Add","Remove","Get")][String]$Action
    
 )

#$StartTime= Get-time

$global:ScriptLocation = $(get-location).Path
$global:DefaultLog = "$global:ScriptLocation\AddDomain.Log"
#Functions
# ---------------------------------------------------------------------------------------------------------------------------------
# StartFunction Write-Log : Log the information into PsDiskHtml.log file
# ---------------------------------------------------------------------------------------------------------------------------------
function Write-Log{
    [CmdletBinding()]
    #[Alias('wl')]
    [OutputType([int])]
    Param
    (
        # The string to be written to the log.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        # The path to the log file.
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [Alias('LogPath')]
        [string]$Path=$DefaultLog,

        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true,
                    Position=2)]
        [ValidateSet("Error","Warn","Info","Load","Execute")]
        [string]$Level="Info",

        [Parameter(Mandatory=$false)]
        [switch]$NoClobber
    )

   
    Process
    {
        
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Warning "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist
        # to create the file include path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Now do the logging and additional output based on $Level
        switch ($Level) {
            'Error' {
                Write-Host $Message -ForegroundColor Red
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ERROR: `t $Message" | Out-File -FilePath $Path -Append
                break;
                }
            'Warn' {
                Write-Warning $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") WARNING: `t $Message" | Out-File -FilePath $Path -Append
                break;
                }
            'Info' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") INFO: `t $Message" | Out-File -FilePath $Path -Append
                break;
                }
            'Load' {
                Write-Host $Message -ForegroundColor Magenta
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") LOAD: `t $Message" | Out-File -FilePath $Path -Append
                break;
                }
            'Execute' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") EXEC: `t $Message" | Out-File -FilePath $Path -Append
                break;
                }
            }
    }
}
function ShowTimeMS{
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline=$True,position=0,mandatory=$true)]	[datetime]$timeStart,
	[Parameter(ValueFromPipeline=$True,position=1,mandatory=$true)]	[datetime]$timeEnd
  )
  BEGIN {
    
  }
  PROCESS {
		write-Verbose "Stamping time"
		write-Verbose  "initial time: $TimeStart"
		write-Verbose "End time: $TimeEnd"
		$diff=New-TimeSpan $TimeStart $TimeEnd
		Write-verbose "Timediff= $diff"
		$miliseconds = $diff.TotalMilliseconds
		Write-output " Total Time in miliseconds is: $miliseconds ms"
		
  }
}
function Domain-Exists{
    [CmdletBinding()]
    [OutputType([bool])]
    Param
    (
        # The string to be written to the log.
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)][ValidateNotNullOrEmpty()][string]$DomainChecked,
        
        # The array of Actual Domain Names
        [Parameter(Mandatory=$true,Position=1)]$array
    )
        $arrayisNull = [string]::IsNullOrEmpty($array)
        #case #false, empty
        if($arrayisNull){
            return $output;
        }
        else{

		    [bool]$output=$false
		    foreach($domain in $array){
			    if($DomainChecked -eq $domain){
				    $output= $true
			    }
		    }
        }
		#True Exists
		#False not exists
		return $output
}
function ADModule-Exists{
    [CmdletBinding()]
    [OutputType([bool])]
   
    $Getmd = Get-Module -ListAvailable | Where-Object{ $_.Name -like "ActiveD*"} | Select Name
    [bool]$output=$false

    if( $Getmd.Name -clike "ActiveDirectory"){
		if(! (Get-Module ActiveDirectory)){
			Write-Log -Level Execute "Loading the Active Directory Module"
			Import-Module ActiveDirectory
			Write-Log -Level Load "Loaded"
		}
		else{
			Write-Log -Level Execute "Active Directory Module for Powershell is already loaded"
		}
        $output= $true
    }
	return $output
}
 #Functions

Write-Log -Level Info -Message "#####**********************************#  START SCRIPT  #******************************#####"

$Global:CleanUpGlobal=@()
$Global:CleanUpVar=@()

if(-not(ADModule-Exists)){
    Write-Log -Level Error "Active Directory Module was not found"
    Write-Log -Level Warn "For this script to run it is required the ActiveDirectory Module, Please run this script in a Windows Server with the Active Directory Module Enabled (or a Domain Controller)"
    exit(0)
}

Write-Log -Level Info "Getting Information about the Forest"
$Global:Forest = Get-ADForest | Select *

Write-Log -Level Info "Getting Information about the Domain"
$Global:Domain = Get-ADDomain | Select *

Write-Log -Level Info "Setting Up Variables and Objects"

#FunctionalLevelds
$Global:FunctionalLevels = New-Object PSObject -Property @{
    Forest = $Global:Forest.ForestMode
    Domain = $Global:Domain.DomainMode
}

#ADRoles
$Global:ADRoles= New-Object PSObject -Property @{
        RIDMaster=$Global:Domain.RIDMaster
        PDCEmulator=$Global:Domain.PDCEmulator
        InfrastructureMaster=$Global:Domain.InfrastructureMaster
        DomainNamingMaster=$Global:Forest.DomainNamingMaster
        SchemaMaster=$Global:Forest.SchemaMaster
}

#Showing AD roles
Write-Log -Level Info -Message "$Global:ADRoles"


#DomainName and Today
$DomainName= $Global:Domain.DNSRoot
$today = get-date -format MM-dd-yyyy

#add cleanupvars
$Global:CleanUpGlobal+="Forest"
$Global:CleanUpGlobal+="Domain"
$Global:CleanUpGlobal+="ADRoles"
$Global:CleanUpGlobal+="FunctionalLevels"
$Global:CleanUpVar+="Today"
$Global:CleanUpVar+="DomainName"
$Suffixes = $Global:Forest.UPNSuffixes

####Check AD module Exists ####

if( $Action -ne "Get" -or $Action -ne "get"){
    $exist = Domain-Exists $NewOldDomain $Suffixes
}


switch -wildcard ($Action){
    '?dd' {
        if($Exist){
            Write-Log -Level Error "The domain $NewOldDomain already Exists and can't be added again. Exiting"
			Write-Log -Level Info -Message "#####**********************************#  END SCRIPT  #********************************#####"
            exit(0)
        }
        Write-Log -Level Execute "Adding the $NewOldDomain to the Active Directory suffixes for the domain: $DomainName on $Today"
        Set-ADForest -identity $DomainName -UPNSuffixes @{Add="$NewOldDomain"}
        break;
    }
    '?emove'{
        if( -not ($Exist) ){
            Write-Log -Level Error "The domain $NewOldDomain does not exists, so can not be removed. Exiting"
			Write-Log -Level Info -Message "#####**********************************#  END SCRIPT  #********************************#####"
            exit(0)
        }
        Set-ADForest -identity $DomainName -UPNSuffixes @{Remove="$NewOldDomain"}
        Write-Log -Level Info -Message "the domain $NewOldDomain has been removed"
    }
    '?et'{
        if($Suffixes.Count -gt 0){
            if($Suffixes.Count -eq 1){
              Write-Log -Level Info -Message "The domains allowed in the Domain is:"
              Write-Log -Level Execute -Message $Suffixes
            }
            else{

               Write-Log -Level Info -Message "The domains allowed in the Domain are:"
               Write-Log -Level Execute -Message $Suffixes
            }
        }
        else{
               Write-Log -Level Execute -Message "You have no suffixes in the current domain"
        }
    }

    default {
         Write-Log -Level Error "The action you required is not allowed, please select 'Add' or 'Remove' domain are the valid options for this item"
		 Write-Log -Level Info -Message "#####**********************************#  END SCRIPT  #********************************#####"
         exit(0)
    }
}


#cleanup
$CleanUpVar| ForEach-Object{
	Remove-Variable $_
	}
$CleanUpGlobal | ForEach-Object{
	Remove-Variable -Scope global $_
}
Remove-Variable -Scope Global CleanUpGlobal,CleanUpVar

Write-Log -Level Info -Message "#####**********************************#  END SCRIPT  #********************************#####"
