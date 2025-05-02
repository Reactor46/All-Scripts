## Begin Get-AntiSpyware
Function Get-AntiSpyware{

<#
.Synopsis
   Get AntiSpyware information

.DESCRIPTION
   Obtain from cim (wmi) as SecutiryCenter shows.
   make sure your OS is Workstation, not as Server. (Because server does not have secutiry Center.)

.EXAMPLE
    # this will obtain from localhost
    Get-AntiSpyware

.EXAMPLE
    # this will obtain from 192.168.100.1 with credential you enter.
    $cred = Get-Credential
    Get-AntiSpyware -computerName 192.168.100.1 -credential $cred

.EXAMPLE
    # this will obtain from 192.168.100.1 with credential you enter.
    $cred = Get-Credential
    "server01","server02" | Get-AntiSpyware -credential $cred

.EXAMPLE
    # Output sample
    --------------------
    isplayName               : Windows Defender
    instanceGuid             : {D68DDC3A-831F-4fae-9E44-DA132C1ACF46}
    pathToSignedProductExe   : %ProgramFiles%\Windows Defender\MSASCui.exe
    pathToSignedReportingExe : %ProgramFiles%\Windows Defender\MsMpeng.exe
    productState             : 397568
    timestamp                : Fri, 25 Oct 2013 14:31:11 GMT
    PSComputerName           : 127.0.0.1
    --------------------    

#>

    [CmdletBinding()]
    Param
    (
        # Input ComputerName you want to check
        [Parameter(Mandatory = 0, 
                   ValueFromPipeline,
                   ValueFromPipelineByPropertyName, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $computerName = [System.Environment]::MachineName,

        # Input PSCredential for $ComputerName
        [Parameter(Mandatory = 0, 
                   Position=1)]
        [System.Management.Automation.PSCredential]
        $credential
    )

    Begin
    {
        $nameSpace = "SecurityCenter2"
        $className = "AntiSpywareProduct"
    }

    Process
    {
        if ($PSBoundParameters.count -eq 0)
        {
            if ((Get-CimInstance -namespace "root" -className "__Namespace").Name -contains $nameSpace)
            {
                Write-Verbose ("localhost cim session")
                Get-CimInstance -Namespace "root\$nameSpace" -ClassName $className
            }
            else
            {
                Write-Warning ("You can not check AntiSpyware with {0} as it not contain SecutiryCenter2" -f $OSName)
            }
        }
        else
        {
            try
            {
                Write-Verbose ("creating cim session for {0}" -f $computerName)
                $cimSession = New-CimSession @PSBoundParameters
                if ((Get-CimInstance -namespace "root" -className "__Namespace" -cimsession $cimSession).Name -contains $nameSpace)
                {
                    Get-CimInstance -Namespace "root\$nameSpace" -ClassName $className -CimSession $cimSession
                }
                else
                {
                    Write-Warning ("{0} not contains namespace {1}, you can not check {2}." -f $computerName, $nameSpace, $className)
                }
            }
            finally
            {
                $cimSession.Dispose()
            }
        }
    }

    End
    {
    }
}
## End Get-AntiSpyware
## Begin Get-AntiVirusProduct
Function Get-AntiVirusProduct { 
[CmdletBinding()] 
param ( 
[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] 
[Alias('name')] 
$computername=$env:computername 
)
$AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct  -ComputerName $computername

#Switch to determine the status of antivirus definitions and real-time protection. 
#The values in this switch-statement are retrieved from the following website: http://community.kaseya.com/resources/m/knowexch/1020.aspx 
switch ($AntiVirusProduct.productState) { 
"262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
default {$defstatus = "Unknown" ;$rtstatus = "Unknown"} 
    }

#Create hash-table for each computer 
$ht = @{} 
$ht.Computername = $computername 
$ht.Name = $AntiVirusProduct.displayName 
$ht.ProductExecutable = $AntiVirusProduct.pathToSignedProductExe 
$ht.'Definition Status' = $defstatus 
$ht.'Real-time Protection Status' = $rtstatus

#Create a new object for each computer 
New-Object -TypeName PSObject -Property $ht

}
## End Get-AntiVirusProduct
## Begin Get-AVStatus2
Function Get-AVStatus2 {

<#
.Synopsis
Get anti-virus product information
.Description
This command uses WMI via the Get-CimInstance command to query the state of installed anti-virus products. The default behavior is to only display enabled products, unless you use -All. You can query by computername or existing CIMSessions.
.Example
PS C:\> Get-AVStatus chi-win10

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : True
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Thu, 21 Jul 2016 15:20:18 GMT
Computername : CHI-WIN10

.Example
PS C:\>  import-csv s:\computers.csv | Get-AVStatus -All | Group Displayname | Select Name,Count | Sort Count,Name

Name                           Count
----                           -----
ESET NOD32 Antivirus 9.0.386.0    12
ESET Endpoint Security 5.0         6
Windows Defender                   4
360 Total Security                 1

Import a CSV file which includes a Computername heading. The imported objects are piped to this command. The results are sent to Group-Object.

.Example
PS C:\> $cs | Get-AVStatus | where {-Not $_.UptoDate}

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : False
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Wed, 20 Jul 2016 11:10:13 GMT
Computername : CHI-WIN11

Displayname  : ESET NOD32 Antivirus 9.0.386.0
ProductState : 266256
Enabled      : True
UpToDate     : False
Path         : C:\Program Files\ESET\ESET NOD32 Antivirus\ecmd.exe
Timestamp    : Thu, 07 Jul 2016 15:15:26 GMT
Computername : CHI-WIN81

You can also pipe CIMSession objects. In this example, the output are enabled products that are not up to date.
.Notes
version: 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

.Inputs
[string[]]
[Microsoft.Management.Infrastructure.CimSession[]]

.Outputs
[pscustomboject]

.Link
Get-CimInstance
#>

[cmdletbinding(DefaultParameterSetName="computer")]

Param(
[Parameter(
 Position = 0, 
 ValueFromPipeline, 
 ValueFromPipelineByPropertyName,
 ParameterSetName="computer")]
[ValidateNotNullorEmpty()]
#The name of a computer to query.
[string[]]$Computername = $env:COMPUTERNAME,

[Parameter(ValueFromPipeline,ParameterSetName = "session")]
#An existing CIMsession.
[Microsoft.Management.Infrastructure.CimSession[]]$WimSession,

#The default is enabled products only.
[switch]$All

)

Begin {
    Write-Verbose "[BEGIN  ] Starting: $($MyInvocation.Mycommand)"  

    Function ConvertTo-Hex {
    Param([int]$Number)
    '0x{0:x}' -f $Number
    }

    #initialize an hashtable of paramters to splat to Get-CimInstance
    $wimParams = @{
    Namespace = "root/SecurityCenter2"
    ClassName = "AntiVirusProduct"
    ErrorAction = "Stop"

    }

    If ($All) {
        Write-Verbose "[BEGIN  ] Getting all AV products"
    }
    
    $results = @()
} #begin
## End Get-AVStatus2

Process {
 
    #initialize an empty array to hold results
    $AV=@()
 
    #display PSBoundparameters formatted nicely for Verbose output  
    [string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
    Write-Verbose "[PROCESS] PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 
    Write-Verbose "[PROCESS] Using parameter set: $($pscmdlet.ParameterSetName)"

    if ($pscmdlet.ParameterSetName -eq 'computer') {
        foreach ($computer in $Computername) {

            Write-Verbose "[PROCESS] Querying $($computer.ToUpper())"
            $wimParams.ComputerName = $computer
            Try {    
                $AV += Get-WMIObject @wimParams
         
            }
            Catch {
                Write-Warning "[$($computer.ToUpper())] $($_.Exception.Message)"
                $wimParams.ComputerName = $null
            }

        } #foreach computer
    } 
    else {
        foreach ($session in $WimSession) {

            Write-Verbose "[PROCESS] Using session $($session.computername.toUpper())"
            $wimParams.CimSession = $session
            Try {    
                $AV += Get-WMIObject @wimParams
         
            }
            Catch {
                Write-Warning "[$($session.computername.ToUpper())] $($_.Exception.Message)"
                $wimParams.cimsession = $null
            }

        } #foreach computer
    }

       foreach ($item in $AV) {
                Write-Verbose "[PROCESS] Found $($item.Displayname)"
                $hx = ConvertTo-Hex $item.ProductState
                $mid = $hx.Substring(3,2)
                if ($mid -match "00|01") {
                    $Enabled = $False
                }
                else {
                    $Enabled = $True
                }
                $end = $hx.Substring(5)
                if ($end -eq "00") {
                    $UpToDate = $True
                }
                else {
                    $UpToDate = $False
                }

                $results += $item | Select Displayname,ProductState,
                @{Name="Enabled";Expression = {$Enabled}},
                @{Name = "UpToDate";Expression = {$UptoDate}},
                @{Name = "Path"; Expression = {$_.pathToSignedProductExe}},
                Timestamp,
                @{Name = "Computername"; Expression = {$_.PSComputername.toUpper()}}

            } #foreach

} #process
## End Get-AVStatus2

End {
    If ($All) {
      $results
    }
    else {
        #filter for enabled only
        ($results).Where({$_.enabled})
    }

    Write-Verbose "[END    ] Ending: $($MyInvocation.Mycommand)"
} #end
## End Get-AVStatus2

} 
## End Get-AVStatus2
## Begin Get-LHSAntiVirusProduct
Function Get-LHSAntiVirusProduct{
<#
.SYNOPSIS
    Get the status of Antivirus Product on local and Remote Computers.

.DESCRIPTION
    It works with MS Security Center and detects the status for most AV products.
    
    Note that this script will only work on Windows XP SP2, Vista, 7, 8.x, 10 
    operating systems as Windows Servers does not have 
    the required WMI SecurityCenter\SecurityCenter(2) name spaces.

.PARAMETER ComputerName
    The computer name(s) to retrieve the info from. 

.EXAMPLE
    Get-LHSAntiVirusProduct
    
    ComputerName             : Localhost
    Name                     : Kaspersky Endpoint Security 10 für Windows
    ProductExecutable        : C:\Program Files (x86)\Kaspersky Lab\Kaspersky Endpoint 
                               Security 10 for Windows SP1\wmiav.exe
    DefinitionStatus         : UP_TO_DATE
    RealTimeProtectionStatus : ON
    ProductState             : 266240
 
.EXAMPLE
    Get-LHSAntiVirusProduct –ComputerName PC1,PC2,PC3

    ComputerName             : PC1
    Name                     : Kaspersky Endpoint Security 10 für Windows
    ProductExecutable        : C:\Program Files (x86)\Kaspersky Lab\Kaspersky Endpoint 
                               Security 10 for Windows SP1\wmiav.exe
    DefinitionStatus         : UP_TO_DATE
    RealTimeProtectionStatus : ON
    ProductState             : 266240
    (..)

.EXAMPLE
    (get-content PClist.txt) | Get-LHSAntiVirusProduct

 .INPUTS
    System.String, you can pipe ComputerNames to this Function

.OUTPUTS
    Custom PSObjects 

.NOTE
    WMI query to get anti-virus infor­ma­tion has been changed.
    Pre-Vista clients used the root/SecurityCenter name­space, 
    while Post-Vista clients use the root/SecurityCenter2 name­space.
    But not only the name­space has been changed, The properties too. 


    More info at http://neophob.com/2010/03/wmi-query-windows-securitycenter2/
    and from this MSDN Blog 
    http://blogs.msdn.com/b/alejacma/archive/2008/05/12/how-to-get-antivirus-information-with-wmi-vbscript.aspx


    AUTHOR: Pasquale Lantella 
    LASTEDIT: 23.06.2016
    KEYWORDS: Antivirus
    Version :1.1
    History :1.1 support for Win 10, changed the use of WMI productState   

.LINK
    WSC_SECURITY_PRODUCT_STATE enumeration
    https://msdn.microsoft.com/en-us/library/jj155490%28v=vs.85%29

.LINK
    Windows Security Center
    https://msdn.microsoft.com/en-us/library/gg537273%28v=vs.85%29

.LINK
    http://neophob.com/2010/03/wmi-query-windows-securitycenter2/

#Requires -Version 2.0
#>


[CmdletBinding()]

param (
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('CN')]
    [String[]]$ComputerName=$env:computername
)

BEGIN {

    Set-StrictMode -Version Latest
    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name

} # end BEGIN
## End Get-LHSAntiVirusProduct

PROCESS {
    
    ForEach ($Computer in $computerName) 
    {
        IF (Test-Connection -ComputerName $Computer -count 2 -quiet) 
        { 
            Try
            {
                [system.Version]$OSVersion = (Get-WmiObject win32_operatingsystem -computername $Computer).version

                IF ($OSVersion -ge [system.version]'6.0.0.0') 
                {
                    Write-Verbose "OS Windows Vista/Server 2008 or newer detected."
                    $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ComputerName $Computer -ErrorAction Stop
                } 
                Else 
                {
                    Write-Verbose "Windows 2000, 2003, XP detected" 
                    $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter -Class AntiVirusProduct  -ComputerName $Computer -ErrorAction Stop
                } # end IF ($OSVersion -ge 6.0) 
 
                <#
                it appears that if you convert the productstate to HEX then you can read the 1st 2nd or 3rd block 
                to get whether product is enabled/disabled and whether definitons are up-to-date or outdated
                #>

                $productState = $AntiVirusProduct.productState

                # convert to hex, add an additional '0' left if necesarry
                $hex = [Convert]::ToString($productState, 16).PadLeft(6,'0')

                # Substring(int startIndex, int length)  
                $WSC_SECURITY_PROVIDER = $hex.Substring(0,2)
                $WSC_SECURITY_PRODUCT_STATE = $hex.Substring(2,2)
                $WSC_SECURITY_SIGNATURE_STATUS = $hex.Substring(4,2)

                #n ot used yet
                $SECURITY_PROVIDER = switch ($WSC_SECURITY_PROVIDER)
                {
                    0  {"NONE"}
                    1  {"FIREWALL"}
                    2  {"AUTOUPDATE_SETTINGS"}
                    4  {"ANTIVIRUS"}
                    8  {"ANTISPYWARE"}
                    16 {"INTERNET_SETTINGS"}
                    32 {"USER_ACCOUNT_CONTROL"}
                    64 {"SERVICE"}
                    default {"UNKNOWN"}
                }


                $RealTimeProtectionStatus = switch ($WSC_SECURITY_PRODUCT_STATE)
                {
                    "00" {"OFF"} 
                    "01" {"EXPIRED"}
                    "10" {"ON"}
                    "11" {"SNOOZED"}
                    default {"UNKNOWN"}
                }

                $DefinitionStatus = switch ($WSC_SECURITY_SIGNATURE_STATUS)
                {
                    "00" {"UP_TO_DATE"}
                    "10" {"OUT_OF_DATE"}
                    default {"UNKNOWN"}
                }  

<#  
                # Switch to determine the status of antivirus definitions and real-time protection.
                # The values in this switch-statement are retrieved from the following website: http://community.kaseya.com/resources/m/knowexch/1020.aspx
                switch ($AntiVirusProduct.productState) {
                     #AVG Internet Security 2012 (from antivirusproduct WMI)
                     "262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                     "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
 
                     "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"}
                     "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"}
                     "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}
                     "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     #Windows Defender
                     "393472" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
                     "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"}
                     "397568" {$defstatus = "Up to date" ;$rtstatus = "Enabled"}

                     default {$defstatus = "Unknown" ;$rtstatus = "Unknown"}
                }
#>

              
                # Output PSCustom Object
                $AV = $Null
                $AV = New-Object -TypeName PSObject -ErrorAction Stop -Property @{
             
                    ComputerName = $AntiVirusProduct.__Server;
                    Name = $AntiVirusProduct.displayName;
                    ProductExecutable = $AntiVirusProduct.pathToSignedProductExe;
                    DefinitionStatus = $DefinitionStatus;
                    RealTimeProtectionStatus = $RealTimeProtectionStatus;
                    ProductState = $productState;
                
                } | Select-Object ComputerName,Name,ProductExecutable,DefinitionStatus,RealTimeProtectionStatus,ProductState  
                
                Write-Output $AV 
            }
            Catch 
            {
                Write-Error "\\$Computer : WMI Error"
                Write-Error $_
            }                              
        } 
        Else 
        {
            Write-Warning "\\$computer DO NOT reply to ping" 
        } # end IF (Test-Connection -ComputerName $Computer -count 2 -quiet)
	   
    } # end ForEach ($Computer in $computerName)

} # end PROCESS
## End Get-LHSAntiVirusProduct

END { Write-Verbose "Function Get-LHSAntiVirusProduct finished." } 
} # end Function Get-LHSAntiVirusProduct
## End Get-LHSAntiVirusProduct

## Begin Get-SEPVersion
Function Get-SEPVersion { 
# All registry keys: http://www.symantec.com/business/support/index?page=content&id=HOWTO75109 
[CmdletBinding()] 
param( 
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for", 
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
[Alias('CN','__SERVER','IPAddress','Server')] 
[System.String] 
$ComputerName 
) 
# Create object to enable access to the months of the year 
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo 
#Set registry value to look for definitions path (depending on 32/64 bit OS) 
$osType=Get-WmiObject Win32_OperatingSystem -ComputerName $computername| Select OSArchitecture 
if ($osType.OSArchitecture -eq "32-bit")  
{ 
# Connect to Registry 
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName) 
# Set Registry keys to query 
$SMCKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
# Obtain Product Version value 
$SMCRegKey = $reg.opensubkey($SMCKey) 
$SEPVersion = $SMCRegKey.GetValue('ProductVersion') 
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate') 
# Convert PatternFileDate to readable date 
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970) 
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]] 
$AVDayFileDate = [string]$AVPatternFileDate[2] 
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate 
# Obtain Sylink Group value 
#$SylinkRegKey = $reg.opensubkey($SylinkKey) 
#$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup') 
}  
## End Get-SEPVersion
else  
{ 
# Connect to Registry 
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName) 
# Set Registry keys to query 
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
 
# Obtain Product Version value 
$SMCRegKey = $reg.opensubkey($SMCKey) 
$SEPVersion = $SMCRegKey.GetValue('ProductVersion') 
  
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue("PatternFileDate") 
 
# Obtain Pattern File Date Value 
$AVRegKey = $reg.opensubkey($AVKey) 
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate') 
  
# Convert PatternFileDate to readable date 
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970) 
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]] 
$AVDayFileDate = [string]$AVPatternFileDate[2] 
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate 
} 
## End Get-SEPVersion
$MYObject = ""| Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate 
$MYObject.ComputerName = $ComputerName 
$MYObject.SEPProductVersion = $SEPVersion 
$MYObject.SEPDefinitionDate = $AVFileVersionDate 
$MYObject 
} 
## End Get-SEPVersion

## Begin Get-SEPVersion
Function Get-SEPVersion {

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage='Name of the computer to query SEP for',
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object -TypeName System.Globalization.DateTimeFormatInfo

# Set Registry keys to query

If((Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem).OSArchitecture -eq '32-bit')
{
$SMCKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC'
$AVKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV'
$SylinkKey = 'SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink'
}

Else
{
$SMCKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC' 
$AVKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV' 
$SylinkKey = 'SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink' 
}

    }


process {


try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + ' ' + $AVMonthFileDate + ' ' + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}


catch [System.Management.Automation.MethodInvocationException]

{
$SEPVersion = 'Unable to connect to computer'
$AVFileVersionDate = ''
$SylinkGroup = ''
}

$MYObject = '' | Select-Object -Property ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}
}
## End Get-SEPVersion

## Begin Get-SEPVersion
Function Get-SEPVersion {
	#### Symantec Enterprise Protection ####
<#
.SYNOPSIS
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.DESCRIPTION
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.PARAMETER  ComputerName
Name of the computer to query SEP info for

.EXAMPLE
PS C:\> Get-SEPVersion -ComputerName Server01

.EXAMPLE
PS C:\> $servers | Get-SEPVersion

.NOTES
Author: Jonathan Medd
Date: 23/12/2011
#>

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for",
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo

# Set Registry keys to query
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
}

process {

try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}
catch [System.Management.Automation.MethodInvocationException]
{
$SEPVersion = "Unable to connect to computer"
$AVFileVersionDate = ""
$SylinkGroup = ""
}

$MYObject = “” | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}

}
## End Get-SEPVersion

## Begin Get-ProcessForeignAddress
Function Get-ProcessForeignAddress{
<#
.SYNOPSIS
	Get all foreignIPAddress for all or specific processname
	
.DESCRIPTION
	Get all foreignIPAddress for all or specific processname
	
.PARAMETER ProcessName
	Specifies the ProcessName to filter on
	
.EXAMPLE
	Get-ProcessForeignAddress
	
	Retrieve all the foreign addresses
	
.EXAMPLE
	Get-ProcessForeignAddress chrome
	
	Show all the foreign address(es) for the process chrome
	
.EXAMPLE
	Get-ProcessForeignAddress chrome | select ForeignAddress -Unique
	
	Show all the foreign address(es) for the process chrome and show only the ForeignAddress(es) once
	
.NOTES
	Author	: Francois-Xavier Cat
	Website	: www.lazywinadmin.com
	Github	: github.com/lazywinadmin
	Twitter	: @lazywinadm
#>
	PARAM ($ProcessName)
	$netstat = netstat -no
	
	$Result = $netstat[4..$netstat.count] |
	ForEach-Object {
		$current = $_.trim() -split '\s+'
		
		New-Object -TypeName PSobject -Property @{
			ProcessName = (Get-Process -id $current[4]).processname
			ForeignAddressIP = ($current[2] -split ":")[0] #-as [ipaddress]
			ForeignAddressPort = ($current[2] -split ":")[1]
			State = $current[3]
		}
	}
	
	if ($ProcessName)
	{
		$result | Where-Object { $_.processname -like "$processname" }
	}
	else { $Result }
}
## End Get-ProcessForeignAddress