#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.14
# Created on:   2/5/2013 6:36 AM
# Created by:   Zachary Loeber
# Organization: 
# Filename: Globals.ps1
# Description: These are used across both the gui and the called script
#              for storing and loading script state data.
# Requires: VMware-Report-GUI.ps1
#           VMware-Report.ps1
#========================================================================
#A few Constants 
$Styles = @('Default','Style1')

#Our base variables which may get overwritten if a config file is loaded.
#With no config file this is where you want to set your default settings.
$varEmailReport=$false
$varEmailSubject=""
$varEmailRecipient=""
$varEmailSender=""
$varSMTPServer=""
$varSaveReportsLocally=$false
$varReportName="report.html"
$varReportFolder="."
$varVIServer="localhost"
$varUseCurrentUser=$false
$varVIUser=""
$varVIPassword=""
$varScopeWholeFarm=$true
$varScopeDatacenter=""
$varScopeCluster=""
$varScopeHost=""
$varReportHostsDatastore=$false
$varReportHostsDatastoreThreshold="10"
$varReportVMSnapshots=$true
$varReportVMSnapshotsThreshold="2"
$varReportVMThinProvisioned=$false
$varReportSelective=$false
$varReportVCVMsCreated = $false
$varReportVCVMsDeleted = $false
$varReportVCErrors = $false
$varReportHostsNotResponding = $false
$varReportHostsInMaint = $false
$varReportVMTools = $false
$varReportCDConnected = $false
$varReportFloppyConnected = $false
$varReportVCVMsCreatedAge = 5
$varReportVCEventLogsAge = 1
$varReportVCErrorsAge = 1
$varReportVCEvntlogs = $false
$varReportVCServices = $false
$varReportStyle = "Default"

#For each variable an xml attribute should exist which maps to our save file
$ConfigTemplate = @"
<Configuration>
    <EmailReport>{0}</EmailReport>
    <EmailSubject>{1}</EmailSubject>
    <EmailRecipient>{2}</EmailRecipient>
    <EmailSender>{3}</EmailSender>
    <SMTPServer>{4}</SMTPServer>
    <SaveReportsLocally>{5}</SaveReportsLocally>
    <ReportName>{6}</ReportName>
    <ReportFolder>{7}</ReportFolder>
    <VIServer>{8}</VIServer>
    <UseCurrentUser>{9}</UseCurrentUser>
    <VIUser>{10}</VIUser>
    <VIPassword>{11}</VIPassword>
    <ScopeWholeFarm>{12}</ScopeWholeFarm>
    <ScopeDatacenter>{13}</ScopeDatacenter>
    <ScopeCluster>{14}</ScopeCluster>
    <ScopeHost>{15}</ScopeHost>
    <ReportDatastore>{16}</ReportDatastore>
    <ReportDatastoreThreshold>{17}</ReportDatastoreThreshold>
    <ReportSnapshots>{18}</ReportSnapshots>
    <ReportSnapshotsThreshold>{19}</ReportSnapshotsThreshold>
    <ReportThinProvisioned>{20}</ReportThinProvisioned>
    <ReportSelective>{21}</ReportSelective>
    <ReportVCVMsCreated>{22}</ReportVCVMsCreated>
    <ReportVCVMsDeleted>{23}</ReportVCVMsDeleted>
    <ReportVCErrors>{24}</ReportVCErrors>
    <ReportHostNotResponding>{25}</ReportHostNotResponding>
    <ReportHostsInMaint>{26}</ReportHostsInMaint>
    <ReportVMTools>{27}</ReportVMTools>
    <ReportCDConnected>{28}</ReportCDConnected>
    <ReportFloppyConnected>{29}</ReportFloppyConnected>
    <ReportVCEventAge>{30}</ReportVCEventAge>
    <ReportVCErrorsAge>{31}</ReportVCErrorsAge>
    <ReportVCEvntlogs>{32}</ReportVCEvntlogs>
    <ReportVCServices>{33}</ReportVCServices>
    <ReportStyle>{34}</ReportStyle>
</Configuration>
"@

# VMware specific globals
#$VIConnection
$VIConnected = $false

#Provides the location of the script
function Get-ScriptDirectory
{ 
	if($hostinvocation -ne $null)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

#Config file location
$ConfigFile = $ScriptDirectory + "\Config.xml"

#Possible credential file
$CredFile = $ScriptDirectory + "\Cred.crd"

#This is the non-GUI version of the script.
$StarterScript = "\VMware-Report.ps1"

# Extra scripts
function Colorize-Table 
{ 
[CmdletBinding(DefaultParameterSetName = "ObjectSet")] 
param ( 
    [Parameter( 
        Mandatory=$true, 
        Position=0, 
        ValueFromPipeline=$true, 
        ParameterSetName="ObjectSet" 
    )] 
    [PSObject[]]$InputObject, 
    [Parameter( 
        Mandatory=$true, 
        Position=0, 
        ValueFromPipeline=$true, 
        ParameterSetName="StringSet" 
    )] 
    [String[]]$InputString='', 
    [Parameter( 
        Mandatory=$true, 
        ValueFromPipeline=$false 
    )] 
    [String]$Column, 
    [Parameter( 
        Mandatory=$true, 
        ValueFromPipeline=$false 
    )] 
    [String]$ColumnValue, 
    [Parameter( 
        Mandatory=$true, 
        ValueFromPipeline=$false 
    )] 
    [String]$Attr, 
    [Parameter( 
        Mandatory=$true, 
        ValueFromPipeline=$false 
    )] 
    [String]$AttrValue, 
    [Parameter( 
        Mandatory=$false, 
        ValueFromPipeline=$false 
    )] 
    [Bool]$WholeRow=$false, 
    [Parameter( 
        Mandatory=$false, 
        ValueFromPipeline=$false, 
        ParameterSetName="ObjectSet" 
    )] 
    [String]$HTMLHead='<title>HTML Table</title>') 
 
BEGIN 
{ 
    Add-Type -ErrorAction SilentlyContinue -Language CSharpVersion3 -ReferencedAssemblies System.Xml, System.Xml.Linq -UsingNamespace System.Linq -Name XUtilities -Namespace Huddled -MemberDefinition @" 
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex( System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index) { 
        return from e in doc.Descendants(element) where e.NodesBeforeSelf().Count() == index select e; 
    } 
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue( System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value) { 
        return from e in doc.Descendants(element) where e.Value == value select e; 
    } 
"@ 
    $Objects = @() 
} 
 
PROCESS 
{ 
    # Handle passing object via pipe 
    $Objects += $InputObject 
} 
 
END 
{ 
    # Convert our data to x(ht)ml 
    if ($InputString)    # If a string was passed just parse it 
    { 
        $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")  
    } 
    else                # Otherwise we have to convert it to html first 
    { 
        $xml = [System.Xml.Linq.XDocument]::Parse("$($Objects | ConvertTo-Html -Head $HTMLHead)")     
    } 
     
    # Find the index of the column you want to format 
    $ColumnLoc = [Huddled.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
    $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
     
    # Process each xml element based on the index for the column we are highlighting 
    switch([Huddled.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
    { 
        {$_.Value -eq $ColumnValue} { 
            $_.SetAttributeValue($Attr, $AttrValue) 
        }  
    } 
    Return $xml.Document.ToString() 
} 
 
<# 
.SYNOPSIS 
Colorize-Table 
 
.DESCRIPTION 
Colorize cells of an array of objects. Otherwise, if an html table is passed through then colorize 
individual cells of it based on row header and value. 
 
.PARAMETER  [PSObject[]]$InputObject 
An array of objects (ie. (Get-process | select Name,Company) 
 
.PARAMETER  Column 
The column you want to modify1 
 
.PARAMETER  ColumnValue 
The column value you will modify if found. 
 
.PARAMETER  Attr 
The attribute to change should ColumnValue be found in the Column specified. 
- A good example is using "style" 
 
.PARAMETER  AttrValue 
The attribute value to set when the ColumnValue is found in the Column specified 
- A good example is using "background: red;" 
 
.EXAMPLE 
This will highlight the process name of Dropbox with a red background. 
 
$tabletocolorize = $(Get-Process | ConvertTo-Html -Head $TableStyle) 
$colorizedtable = Colorize-Table $tabletocolorize "Name" "Dropbox" "style" "background: red;" 
 
 
.EXAMPLE 
This will highlight the process name of Dropbox with a red background. Then ccSvcHost process names 
will be highlighted yellow. Finally Any Company Name of some big global conglomerate will be  
highlighted green. This all gets wrapped up in a nice single line outlined style for the table, gets 
saved to a file in the local directory, then gets invoked (opened, usually via IE). 
 
$TableStyle = "<title>Process Report</title> 
             
            BODY{font-family: Arial; font-size: 8pt;} 
            H1{font-size: 16px;} 
            H2{font-size: 14px;} 
            H3{font-size: 12px;} 
            TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
            TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
            TD{border: 1px solid black; padding: 5px; } 
            " 
 
$tabletocolorize = $(Get-Process | Select Name,Company) 
 
$colorizedtable = (Colorize-Table $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;" -HTMLHead $TableStyle) 
$test = Colorize-Table $colorizedtable -Column "Name" -ColumnValue "ccSvcHst" -Attr "style" -AttrValue "background: yellow;" 
$test2 = Colorize-Table $test -Column "Company" -ColumnValue "Microsoft Corporation" -Attr "style" -AttrValue "background: green;" 
$test2 | Out-File "$pwd/procs4.html" 
ii "$pwd/procs4.html" 
 
.NOTES 
If you are going to convert something to html with convertto-html in powershell v2 there is a bug where the  
header will show up as an asterick if you only are converting one object property. 
 
This has a long way to go before really being robust and useful but it is a start... 
.LINK 
http://www.the-little-things.net 
#> 
} 

Function Get-SaveData
{
    [xml]$x=($ConfigTemplate) -f 
        $varEmailReport,`
        $varEmailSubject,`
        $varEmailRecipient,`
        $varEmailSender,`
        $varSMTPServer,`
        $varSaveReportsLocally,`
        $varReportName,`
        $varReportFolder,`
        $varVIServer,`
        $varUseCurrentUser,`
        $varVIUser,`
        $varVIPassword,`
        $varScopeWholeFarm,`
        $varScopeDatacenter,`
        $varScopeCluster,`
        $varScopeHost,`
        $varReportHostsDatastore,`
        $varReportHostsDatastoreThreshold,`
        $varReportVMSnapshots,`
        $varReportVMSnapshotsThreshold,`
        $varReportVMThinProvisioned,`
        $varReportSelective,`
        $varReportVCVMsCreated,`
        $varReportVCVMsDeleted,`
        $varReportVCErrors,`
        $varReportHostsNotResponding,`
        $varReportHostsInMaint,`
        $varReportVMTools,`
        $varReportCDConnected,`
        $varReportFloppyConnected,`
        $varReportVCEventLogsAge,`
        $varReportVCErrorsAge,`
        $varReportVCEvntlogs,`
        $varReportVCServices,`
        $varReportStyle
    return $x
}

Function Load-Config
{
	if (Test-Path $ConfigFile)
	{
		[xml]$configuration = get-content $($ConfigFile)
        $Script:varEmailReport = [System.Convert]::ToBoolean($configuration.Configuration.EmailReport)
        $Script:varEmailSubject = $configuration.Configuration.EmailSubject
        $Script:varEmailRecipient = $configuration.Configuration.EmailRecipient
        $Script:varEmailSender = $configuration.Configuration.EmailSender
        $Script:varSMTPServer = $configuration.Configuration.SMTPServer
        $Script:varSaveReportsLocally = [System.Convert]::ToBoolean($configuration.Configuration.SaveReportsLocally)
        $Script:varReportName = $configuration.Configuration.ReportName
        $Script:varReportFolder = $configuration.Configuration.ReportFolder
        $Script:varVIServer = $configuration.Configuration.VIServer
        $Script:varUseCurrentUser = [System.Convert]::ToBoolean($configuration.Configuration.UseCurrentUser)
        $Script:varVIUser = '' #$configuration.Configuration.VIUser
        $Script:varVIPassword = '' #$configuration.Configuration.VIPassword
        $Script:varScopeWholeFarm = [System.Convert]::ToBoolean($configuration.Configuration.ScopeWholeFarm)
        $Script:varScopeDatacenter = $configuration.Configuration.ScopeDatacenter
        $Script:varScopeCluster = $configuration.Configuration.ScopeCluster
        $Script:varScopeHost = $configuration.Configuration.ScopeHost
        $Script:varReportHostsDatastore = [System.Convert]::ToBoolean($configuration.Configuration.ReportDatastore)
        $Script:varReportHostsDatastoreThreshold = $configuration.Configuration.ReportDatastoreThreshold
        $Script:varReportVMSnapshots = [System.Convert]::ToBoolean($configuration.Configuration.ReportSnapshots)
        $Script:varReportVMSnapshotsThreshold = $configuration.Configuration.ReportSnapshotsThreshold
        $Script:varReportVMThinProvisioned = [System.Convert]::ToBoolean($configuration.Configuration.ReportThinProvisioned)
        $Script:varReportSelective = [System.Convert]::ToBoolean($configuration.Configuration.ReportSelective)
        $Script:varReportVCVMsCreated = [System.Convert]::ToBoolean($configuration.Configuration.ReportVCVMsCreated)
        $Script:varReportVCVMsDeleted = [System.Convert]::ToBoolean($configuration.Configuration.ReportVCVMsDeleted)
        $Script:varReportVCErrors = [System.Convert]::ToBoolean($configuration.Configuration.ReportVCErrors)
        $Script:varReportHostsNotResponding = [System.Convert]::ToBoolean($configuration.Configuration.ReportHostNotResponding)
        $Script:varReportHostsInMaint = [System.Convert]::ToBoolean($configuration.Configuration.ReportHostsInMaint)
        $Script:varReportVMTools = [System.Convert]::ToBoolean($configuration.Configuration.ReportVMTools)
        $Script:varReportCDConnected = [System.Convert]::ToBoolean($configuration.Configuration.ReportCDConnected)
        $Script:varReportFloppyConnected = [System.Convert]::ToBoolean($configuration.Configuration.ReportFloppyConnected)
        $Script:varReportVCEventLogsAge = $configuration.Configuration.ReportVCEventAge
        $Script:varReportVCErrorsAge = $configuration.Configuration.ReportVCErrorsAge
        $Script:varReportVCEvntlogs = [System.Convert]::ToBoolean($configuration.Configuration.ReportVCEvntlogs)
        $Script:varReportVCServices = [System.Convert]::ToBoolean($configuration.Configuration.ReportVCServices)
        $Script:varReportStyle = $configuration.Configuration.ReportStyle
        if (-not $varUseCurrentUser)
        {
            $Snapin='VMware.VimAutomation.Core'
            Add-PSSnapin $Snapin –ErrorAction SilentlyContinue
    		if ((Get-PSSnapin $Snapin) –ne $NULL) 
    		{
    			$creds = Get-VICredentialStoreItem -file $CredFile 
                $Script:varVIUser = [string]$creds.User
                $Script:varVIPassword = [string]$creds.Password
    		} 
        }
        Return $true
	}
    else
    {
        Return $false
    }
}

# Save exceptions
Function Save-Config
{
    $SanitizedConfig = $true
	if (($varEmailReport) -and`
        (($varEmailSubject -eq "") -or`
		 ($varEmailRecipient -eq "") -or`
         ($varEmailSender -eq "") -or`
         ($varSMTPServer -eq "")))
	{
		#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		[void][System.Windows.Forms.MessageBox]::Show("You selected to send an email but didn't fill in the right stuff to make it happen buddy.","Sorry, try again.")
        $SanitizedConfig = $false
	}
	elseif (($varSaveReportsLocally) -and`
			(($varReportName -eq "") -or`
			 ($varReportFolder -eq "")))
	{
		#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		[void][System.Windows.Forms.MessageBox]::Show("You selected to not save locally (so are assumed to be attempting to email the reports) but didn't fill in email configuration information.","Sorry, not going to do it.")
        $SanitizedConfig = $false
	}
   	elseif ((!$varScopeWholeFarm) -and`
			($varScopeDatacenter -eq ""))
    {
   		#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
		[void][System.Windows.Forms.MessageBox]::Show("You selected to not report on the whole farm and then made no other selections. `n`n Not Saved.","Sorry, that will not work..")
        $SanitizedConfig = $false
	}
 
    if ($SanitizedConfig)
	{
		# save the data
		[xml]$x=Get-SaveData
        $x.save($ConfigFile)
        Return $true
	}
    else
    {
        Return $false
	}
}
