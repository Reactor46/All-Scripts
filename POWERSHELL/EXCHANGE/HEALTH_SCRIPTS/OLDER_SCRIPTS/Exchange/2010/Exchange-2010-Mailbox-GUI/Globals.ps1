#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.14
# Created on:   2/5/2013 6:36 AM
# Created by:   Zachary Loeber
# Organization: 
# Filename: Globals.ps1
# Description: These are used across both the gui and the called script
#              for storing and loading script state data.
#========================================================================

# Some dot sourcing love
. .\Custom-Script.ps1

# RequiredSnapins
$Snapins=@(’Microsoft.Exchange.Management.PowerShell.E2010’)

#Our base variables
$varEmailReport=$false
$varEmailSubject=""
$varEmailRecipient=""
$varEmailSender=""
$varSMTPServer=""
$varSaveReportsLocally=$false
$varReportName="report.html"
$varReportFolder="."
$varServer="localhost"
$varUseCurrentUser=$true
$varUser=""
$varPassword=""
$varScopeEnterprise=$true
$varScopeDAG=""
$varScopeServer=""
$varScopeMailboxDatabase=""

#For each variable an xml attribute should exist
$ConfigTemplate = 
@"
<Configuration>
    <EmailReport>{0}</EmailReport>
    <EmailSubject>{1}</EmailSubject>
	<EmailRecipient>{2}</EmailRecipient>
	<EmailSender>{3}</EmailSender>
	<SMTPServer>{4}</SMTPServer>
	<SaveReportsLocally>{5}</SaveReportsLocally>
	<ReportName>{6}</ReportName>
	<ReportFolder>{7}</ReportFolder>
    <Server>{8}</Server>
    <UseCurrentUser>{9}</UseCurrentUser>
    <User>{10}</User>
    <Password>{11}</Password>
    <ScopeEnterprise>{12}</ScopeEnterprise>
    <ScopeDAG>{13}</ScopeDAG>
    <ScopeServer>{14}</ScopeServer>
    <ScopeMailboxDatabase>{15}</ScopeMailboxDatabase>
</Configuration>
"@

# Exchange specific globals that do not get saved
$EXConnected = $false
$MailboxServerScope = @()
$Mailboxes = @()
$SelectedMailboxes =@()

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

#This is the non-GUI version of the script.
#$StarterScript = "\MainScript.ps1"

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
        $varServer,`
        $varUseCurrentUser,`
        $varUser,`
        $varPassword,`
        $varScopeEnterprise,`
        $varScopeDAG,`
        $varScopeServer,`
        $varScopeMailboxDatabase
    return $x
}

function Load-Config
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
        $Script:varServer = $configuration.Configuration.VIServer
        $Script:varUseCurrentUser = [System.Convert]::ToBoolean($configuration.Configuration.UseCurrentUser)
        $Script:varUser = $configuration.Configuration.User
        $Script:varPassword = $configuration.Configuration.Password
        $Script:varScopeEnterprise = [System.Convert]::ToBoolean($configuration.Configuration.ScopeEnterprise)
        $Script:varScopeDAG = $configuration.Configuration.ScopeDAG
        $Script:varScopeServer = $configuration.Configuration.ScopeServer
        $Script:varScopeMailboxDatabase = $configuration.Configuration.ScopeMailboxDatabase
        Return $true
	}
    else
    {
        Return $false
    }
}

# Save exceptions
function Save-Config
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
   	elseif ((!$varScopeEnterprise) -and`
			($varScopeDAG -eq ""))
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

function LoadSnapins
{
    $AllRequiredSnapinsLoaded=$True
    if (($Snapins.Count -ge 1) -and $(SnapinsAvailable)) 
    {
    	Foreach ($Snapin in $Snapins)
    	{
            Add-PSSnapin $Snapin –ErrorAction SilentlyContinue 
    		if ((Get-PSSnapin $Snapin –ErrorAction SilentlyContinue) –eq $NULL) 
    		{
    			$AllRequiredSnapinsLoaded=$false
    		}
     	}
    }
    else
    {
        $AllRequiredSnapinsLoaded=$false
    }
    Return $AllRequiredSnapinsLoaded
}

function SnapinsAvailable
{
    $RegisteredSnapins=@(Get-PSSnapin -Registered)
    $RequiredSnapinsRegistered = $true
    if ($Snapins.Count -ge 1) 
    {
    	Foreach ($Snapin in $Snapins)
    	{
            if (!($RegisteredSnapins -match $Snapin))
            {
                $RequiredSnapinsRegistered = $false
            }
		}
	}
    
    Return $RequiredSnapinsRegistered
}

####################### 
function Get-Type 
{ 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
} #Get-Type 
 
####################### 
<# 
.SYNOPSIS 
Creates a DataTable for an object 
.DESCRIPTION 
Creates a DataTable based on an objects properties. 
.INPUTS 
Object 
    Any object can be piped to Out-DataTable 
.OUTPUTS 
   System.Data.DataTable 
.EXAMPLE 
$dt = Get-psdrive| Out-DataTable 
This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable 
.NOTES 
Adapted from script by Marc van Orsouw see link 
Version History 
v1.0  - Chad Miller - Initial Release 
v1.1  - Chad Miller - Fixed Issue with Properties 
v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
v1.4  - Chad Miller - Corrected issue with DBNull 
v1.5  - Chad Miller - Updated example 
v1.6  - Chad Miller - Added column datatype logic with default to string 
v1.7 - Chad Miller - Fixed issue with IsArray 
.LINK 
http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx 
#> 
function Out-DataTable 
{ 
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = new-object Data.datatable   
        $First = $true  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
} #Out-DataTable

Function New-ExchangeSession 
{ 
<# 
   .Synopsis 
    This function creates an implicit remoting connection to an Exchange Server  
   .Description 
    This function creates an implicit remoting session to a remote Exchange  
    Server. It has been tested on Exchange 2010. The Exchange commands are 
    brought into the local PowerShell environment. This works in both the 
    Windows PowerShell console as well as the Windows PowerShell ISE. It requires 
    two parameters: the computername and the user name with rights on the remote  
    Exchange server. 
   .Example 
    New-ExchangeSession -computername ex1 -user iammred\administrator 
    Makes an implicit remoting connection to a remote Exchange 2010 server 
    named ex1 using the administrator account from the iammred domain. The user 
    is prompted for the administrator password. 
   .Parameter ComputerName 
    The name of the remote Exchange server 
   .Parameter User 
    The user account with rights on the remote Exchange server. The user 
    account is specified as domain\username 
   .Notes 
    NAME:  New-ExchangeSession 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 01/13/2012 17:05:32 
    KEYWORDS: Messaging & Communication, Microsoft Exchange 2010, Remoting 
    HSG: HSG-1-23-12 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
     Param( 
      [Parameter(Mandatory=$true,Position=0)] 
      [String] 
      $computername, 
      [Parameter(Mandatory=$true,Position=1)] 
      [String] 
      $user,
      [Parameter(Mandatory=$true,Position=2)] 
      [String] 
      $pass
      ) 
    $pass2 = ConvertTo-SecureString -AsPlainText $pass -Force
    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $user,$pass2
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$computername/powershell -Credential $credential
    Import-PSSession $session
} #end function New-ExchangeSession