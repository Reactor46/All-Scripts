function Get-SEPVersion {
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

If((Get-WmiObject -ComputerName $ComputerName Win32_OperatingSystem).OSArchitecture -eq '32-bit')
{
$SMCKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC"
$AVKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\AV"
$SylinkKey = "SOFTWARE\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink"
}
Else
{
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
}
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