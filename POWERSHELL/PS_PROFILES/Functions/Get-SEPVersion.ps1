function Get-SEPVersion { 
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
$MYObject = ""| Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate 
$MYObject.ComputerName = $ComputerName 
$MYObject.SEPProductVersion = $SEPVersion 
$MYObject.SEPDefinitionDate = $AVFileVersionDate 
$MYObject 
} 