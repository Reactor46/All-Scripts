#############################################################################
#                                     			 		                    #
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 		                    #
#   Author: Koos Botha                                                      #
#   Public Release Version 1.2                            			 	    #
#   Last Update Date:23 August 2016                           	            #
#   Usage  Run .\main.ps1 with Powershell               		            #
#   Requirements Exchange 2010 or 2013 Servers, AD PS module 		        #
#############################################################################

<#
.Synopsis
   Gather information on your exchange environment and document for operational documentation

.DESCRIPTION
   The script is a collection of function that will collect various information within the exchange organization. Once the Data is collected a word document can be generated that contains this information
   .
.EXAMPLE
    This Example shows how to execute the command once the data has been collected to generate a word document
   .\Main.ps1 -GenerateReport

.EXAMPLE
    This Example shows how to execute the command once the data has been collected to generate a word document and include hybrid information
   .\Main.ps1 -GenerateReport -IncludeHybrid

.EXAMPLE
  This Example shows how to start collection of information including Hybrid and Office 365 Subscription information
  .\Main.ps1 -CollectInformation -ExchangeServer <ServerName> -IncludeHybrid
   
.EXAMPLE
  This Example shows how to start collection of information
  .\Main.ps1 -CollectInformation -ExchangeServer <ServerName>

.EXAMPLE
   This Example show how to start the collection of information, with excluding some servers from the report. These servers might be unavailble, etc.
   .\Main.ps1 -CollectInformation -ExchangeServer <ServerName> -Exclude ServerName2,ServerName3

.EXAMPLE
   This Example show how to start Drawing a Visio Diagram. {beta}
   .\Main.ps1 -DrawDiagram


.PARAMETER ExchangeServer
    The parameter is a the value that indicates the Exchange server to establish a remote session with.

.PARAMETER CollectInformation
 The parameter is a Switch to indicate collection of information
   

.PARAMETER GenerateReport
 The parameter is a Switch to indicate to start word document generation. Word must be availble on machine this is execute on with the data directory included.

.PARAMETER DrawDiagram
 The parameter is a Switch to indicate to start drawing a diagram using visio from the collected information.

.PARAMETER IncludeHybrid
 The parameter is a Switch to indicate inclusion of hybrid Configuration data and O365 Subscription information.

#>
#requires -version 2.0
#Set-StrictMode -Version 2.0
param(
[Parameter(ParameterSetName='Parameter Set 1',Mandatory=$true)][String]$ExchangeServer,
[Parameter(ParameterSetName='Parameter Set 1')][Switch]$CollectInformation,
[Parameter(ParameterSetName='Parameter Set 2')][Switch]$GenerateReport,
[Parameter(ParameterSetName='Parameter Set 3')][switch]$DrawDiagram,
[Parameter()][switch]$IncludeHybrid,
[Parameter(ParameterSetName='Parameter Set 1')][Array]$Exclude=@())


$Global:ExcludedServers = @($Exclude)

If ($IncludeHybrid -and !($GenerateReport))
{
$HybridCredentials = Get-Credential -Message "Please provide your Office 365 Credentials"
}

Function Add-text 
{
param([string]$string, $color="Green")
Write-Host $string -ForegroundColor $color -BackgroundColor Black
}

IF ($CollectInformation) 
{
Add-text "Creating Remote session with Exchange server $ExchangeServer"
$commands = "Get-OrganizationConfig","Get-MailboxServer","Get-MailboxDatabase","Get-ExchangeServer","Get-ClientAccessServer","Get-OwaVirtualDirectory"," Get-OabVirtualDirectory","Get-WebServicesVirtualDirectory","Get-ActiveSyncVirtualDirectory","Get-AutodiscoverVirtualDirectory","Get-EcpVirtualDirectory","Get-DatabaseAvailabilityGroup","Get-OutlookAnywhere",`
"Get-receiveconnector","Get-SendConnector","Get-TransportRule","Get-AddressList","Get-AcceptedDomain","Get-OfflineAddressBook","Get-RetentionPolicy","Get-RetentionPolicyTag","Get-AddressBookPolicy","Get-OwaMailboxPolicy","Get-RemoteDomain","Get-MobileDeviceMailboxPolicy","Get-ActiveSyncMailboxPolicy","Get-TransportConfig","Get-UMDialPlan","Get-UMIPGateway",`
"Get-UMMailboxPolicy","Get-UMAutoAttendant","Get-EmailAddressPolicy","Get-DatabaseAvailabilityGroupNetwork","Get-ClientAccessarray","Get-ManagementRoleAssignment","Get-MailboxDatabaseCopyStatus","Get-HybridConfiguration",`
"Get-MsolCompanyInformation","Get-MsolDomain","Connect-MsolService","Get-OabVirtualDirectory","Get-Mailbox","Get-AvailabilityAddressSpace","Get-AdSiteLink","Get-AdSite","Get-ADServerSettings"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($ExchangeServer)/PowerShell/" -Authentication Kerberos -ErrorAction SilentlyContinue -ErrorVariable remoteLoad
Import-PSSession $Session -CommandName $commands -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -AllowClobber  | Out-Null

If ($remoteLoad.COUNT -ge 1)
{
Write-host "An error occured trying to establish a session with server $($ExchangeServer). The error was $($remoteload[0].Exception)" -ForegroundColor Red -BackgroundColor Black
Remove-PSSession $Session -ErrorAction SilentlyContinue
Exit
}
}
#Region Data Directory
IF (Test-Path ((Get-location).path + "\Data"))
	{Add-text "Data directory found"}
	ELSE
	{
		Try{
			New-Item -Name "Data" -ItemType Container
			Write-Host ".\Data directory created for output" -ForegroundColor Green -BackgroundColor Black
			}
	Catch {Add-text $_.message}
	}
#endregion

	
IF ($CollectInformation) 
		{
		#$buttonRun.Enabled = $false
		Write-Host "`nStart Collecting Information..." -ForegroundColor Green -BackgroundColor Black
		$error.Clear()
		$ErrorActionPreference = "Continue"
		
#region Includes

. .\includes\ADsettings.ps1
. .\includes\Exchange.ps1
. .\includes\WMICollection.ps1
. .\includes\Formats.ps1

#endregion includes

#region document Variables

$ExOrg = ""
$Heading = ""
$body = ""
$ShellVersion = ""
#endregion Document Variables


#region Start Exchange Content
                                                    try{
                                                        $ErrorActionPreference = "Stop"
                                                        Get-OrganizationConfig | select Name | Export-Csv ".\data\orgname.csv"
                                                        }
                                                        Catch {
                                                        Write-host $_.Message -ForegroundColor Red
                                                        }
	
                                                        $CSVEXImport = Import-Clixml .\includes\ExchangeContent.xml
                                                       If ($IncludeHybrid)
                                                       {
                                                       Write-Host "Including Hybrid" -ForegroundColor Green -BackgroundColor Black
                                                      # Connect-MsolService -Credential $HybridCredentials 

                                                       $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credential $HybridCredentials -Authentication Basic -AllowRedirection 
                                                       Import-PSSession $O365Session -Prefix 'O365' -CommandName "Get-AvailabilityAddressSpace","Get-AcceptedDomain","Get-OrganizationRelationship","Get-InboundConnector","Get-OutboundConnector","Get-HostedConnectionFilterPolicy","Get-TransportRule","Get-FederationTrust","get-SharingPolicy","Get-IntraOrganizationConnector"

                                                       $CSVEXImport += Import-Clixml .\includes\MSOnline.xml
                                                       }
                                                        
                                                        
                                                        Foreach($ExGroup in $CSVEXImport)
                                                        {
                                                                IF (!($ExGroup.CallFunc -eq "[]") -and !($ExGroup.CallFunc -eq "Get-DAGDistribution") )
                                                                {
			                                                    $Data = @(&($ExGroup.Callfunc)) 
			                                                    $Data | Export-Csv ".\Data\$($ExGroup.Export)"
			                                                    }
                                                                ELSEIF ($ExGroup.CallFunc -eq "Get-DAGDistribution")
                                                                {
                                                                $Data = @(&($ExGroup.Callfunc))
                                                                #Data export Happens in Function
                                                                }
            
                                                        }




#region WMI Import
$WMIContent = @(Import-Clixml .\includes\wmiContent.xml)
#try {
$ErrorActionPreference = "Stop"
       ForEach ($WMIGroup in $WMIContent)
        {
 
            try{
	           IF (!($WMIGroup.Class -eq "[]"))
	             {
	            #Get WMI Content for each item in CSV File
	            $TableArray = Get-WMI $WMIGroup.Class $WMIGroup.Property.split(",") 
	                If (!($TableArray -eq $null) )
	                {
	                #Write the Wmi Data Content to Word Document after Body Text for each Section
	                $TableArray | Export-Csv ".\Data\$($WMIGroup.Export.tostring())" -Force
	                }
             	 }
            	} 
            Catch [system.exception] { Add-text $_.message }
        }

Set-Content -Value $error -Path .\Error.log

#$buttonRun.Enabled = $true
Add-text "Completed"

}
	ELSEIF ($GenerateReport) 
        {
#$buttonRun.Enabled = $false
Add-text "Generate Report"
$error.Clear()
$ErrorActionPreference = "Continue"


#region includes
. .\includes\Application.ps1
#endregion includes

#region document Variables
$ExOrg = @(Import-Csv ".\data\orgname.csv")[0].Name
$Heading = ""
$body = ""

#endregion Document Variables

#region Create Word Application and Copy Template for use

TRY {
$ErrorActionPreference = "Stop"
$FileName = (Get-location).path + "\ExchangeConfig$(get-date -Format ddMMyyhhmm).doc"
Copy-Item ((Get-location).path + "\Template.doc") $FileName -ErrorAction Stop
}
Catch [system.exception] {Add-text $_.Message
}

Try{
$ErrorActionPreference = "Stop"
$word = New-Object -ComObject "Word.application"
$doc = $word.Documents.open($FileName)
$word.visible = $False
}
Catch [system.exception]
{
 Add-text $_.Message -color red
 Add-text  "Microsoft Word is a requirement for this script. The script will now exit." -color red
 Exit
}
#endregion Create Word Application and Copy Template for use

#region Opening Page
# Write the First Page to the Word Document via Write-Landing Function
Write-Landing 

#endregion Opening Page

#region Start Exchange Content
$CSVEXImport = Import-Clixml .\includes\ExchangeContent.xml
If ($IncludeHybrid)
{
$CSVEXImport += Import-Clixml .\includes\MSonline.xml
}
    Foreach($ExGroup in $CSVEXImport)
    {
	If (Test-Path ".\Data\$($ExGroup.Export.tostring())")
	{
        Write-Doc $ExGroup.TextPara $ExGroup.Heading $ExGroup.HeadingFormat
            IF (!($ExGroup.CallFunc -eq "[]") -and !($ExGroup.CallFunc -eq "Get-DAGDistribution"))
            {
            $Array = @(Import-Csv ".\Data\$($ExGroup.Export.tostring())")
            Update-Table $Array $ExGroup.HeaderDirection $ExGroup.HeaderHeight
            }
            ELSEIF ($ExGroup.CallFunc -eq "Get-DAGDistribution")
            {

                $Array  = @()
                $Array += Import-Csv .\Data\DAGDistribution_DAG.csv
               
                Update-Table $Array $ExGroup.HeaderDirection $ExGroup.HeaderHeight
                

            }
	}
	ELSE
	{Add-text "Cannot find .\Data\$($ExGroup.Export.tostring())"}
    }
#endregion Start Exchange Content

#region Start WMI Import
$WMIContent = @(Import-Clixml .\includes\wmiContent.xml)
#try {
$ErrorActionPreference = "Stop"
       ForEach ($WMIGroup in $WMIContent)
        {
			IF (Test-Path ".\Data\$($WMIGroup.Export.tostring())")
			{
            #Write Heading and body (Text paragraph) information to word Document
            try{
			#Add-text "Starting  $WMIGroup.Heading"
            Write-Doc $WMIGroup.TextPara $WMIGroup.Heading $WMIGroup.HeadingFormat    # Write Document Function positional paramters (posision 1 = Text Body)(posision 2 = header)(posision 3 = Header Type formating)
            }
            Catch [system.exception]{
            Add-text $_.message 
            }
            try{
			
            IF (!($WMIGroup.Class -eq "[]"))
             {
            #Get WMI Content for each item in CSV File
            $TableArray =@(Import-Csv ".\Data\$($WMIGroup.Export.tostring())")
                If (!($TableArray -eq $null) )
                {
                #Write the Wmi Data Content to Word Document after Body Text for each Section
                Update-Table $TableArray $WMIGroup.HeaderDirection $WMIGroup.HeaderHeight
                }
            }
            } 
            Catch [system.exception] { Add-text $_.message -color red }
			}
			ELSE
			{Add-text "Cannot find .\Data\$($WMIGroup.Export.tostring())" -color red}
        }

#endregion

#region Finishing Word document and saving

$doc.TablesOfContents.Item(1).update()            #Update Table of Content
$doc.Save()
$Word.Quit()
#endregion Finishing Word Doc

Set-Content -Value $error -Path .\Error.log

Add-text "Completed"
}
    ELSEIF ($drawDiagram) 
        {
       & "$((Get-Location).Path)\Visio.ps1" 
        }
    



