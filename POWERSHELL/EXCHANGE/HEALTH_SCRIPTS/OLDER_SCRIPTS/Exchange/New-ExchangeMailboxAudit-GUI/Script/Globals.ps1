#========================================================================
# Created by:   Zachary Loeber
# Filename: Globals.ps1
# Description: These are used across both the gui and the called script
#              for storing and loading script state data.
#========================================================================

# RequiredSnapins
$Snapins=@(’Microsoft.Exchange.Management.PowerShell.E2010’)

#Our base variables
$varEmailReport=$false
$varEmailSubject="Mailbox Audit Report"
$varEmailRecipient="italerts@usonv.com"
$varEmailSender="MailboxAudit@uson.local"
$varSMTPServer="smtp.uson.local"
$varSaveReportsLocally=$true
$varReportName="ExchangeMailboxAuditReport.html"
$varReportFolder="\\msoit01\E$\web\reports\All_Results\Exchange\"
$varServer="USONVSVREX01"
$varUseCurrentUser=$true
$varUser=""
$varPassword=""
$varScopeEnterprise=$true
$varScopeDAG=""
$varScopeServer=""
$varScopeMailboxDatabase=""

#Different Report Option defaults
$varMailboxReportPermissions = @()
$hashAllPermReports = @(@{Option="Mailbox Summary Information"; Selected=$false},
                        @{Option="Full Access Permissions"; Selected=$false},
                        @{Option="Send On Behalf Permissions"; Selected=$false},
                        @{Option="Send As Permissions"; Selected=$false},
						@{Option="Calendar Permissions"; Selected=$false},
                        @{Option="Mailbox Rule - Forwarding"; Selected=$false},
                        @{Option="Mailbox Rule - Redirecting"; Selected=$false})
        
foreach ($hashOption in $hashAllPermReports)
{
    $Newobject = New-Object PSObject -Property $hashOption
    $varMailboxReportPermissions += $Newobject
}

$varMailboxReportIgnoredUsers = @(  "NT AUTHORITY\SYSTEM",
                                    "NT AUTHORITY\SELF")
$varIncludeInherited = $true
$varExcludeZeroResults = $true
$varExcludeUnknownUsers = $true
$varMailboxRuleForwarding = $false
$varMailboxRuleRedirecting = $false
$varSummaryReport = $true
$varFullAccessReport = $false
$varSendAsReport = $false
$varSendOnBehalfReport = $false
$varCalendarPermReport = $false
$varFlagWarnings = $false
$varMailboxSizeWarning = 8000
$varMailboxSizeAlert = 10000
$varDeletedSizeWarning = 2048
$varDeletedSizeAlert = 4096

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
    <IncludeInherited>{16}</IncludeInherited>
    <ExcludeZeroResults>{17}</ExcludeZeroResults>
    <ExcludeUnknownUsers>{18}</ExcludeUnknownUsers>
	<FlagWarnings>{19}</FlagWarnings>
	<MailboxSizeWarning>{20}</MailboxSizeWarning>
	<MailboxSizeAlert>{21}</MailboxSizeAlert>
	<DeletedSizeWarning>{22}</DeletedSizeWarning>
	<DeletedSizeAlert>{23}</DeletedSizeAlert>
</Configuration>
"@

# Exchange specific globals that do not get saved
$EXConnected = $false
#$MailboxServerScope = @()
$datagridMailboxes = @()
$SelectedMailboxes =@()

#Provides the location of the script

#Config file location
$ConfigFile = "E:\Scripts\Exchange\New-ExchangeMailboxAudit-GUI\Script\Config.xml"

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
        $varScopeMailboxDatabase,`
		$varIncludeInherited,`
		$varExcludeZeroResults,`
		$varExcludeUnknownUsers,`
		$varFlagWarnings, `
		$varMailboxSizeWarning, `
		$varMailboxSizeAlert, `
		$varDeletedSizeWarning, `
		$varDeletedSizeAlert
    return $x
}

function Load-Config
{
	if (Test-Path $ConfigFile)
	{
		[xml]$configuration = Get-Content $ConfigFile
        $Script:varEmailReport = [System.Convert]::ToBoolean($configuration.Configuration.EmailReport)
        $Script:varEmailSubject = $configuration.Configuration.EmailSubject
        $Script:varEmailRecipient = $configuration.Configuration.EmailRecipient
        $Script:varEmailSender = $configuration.Configuration.EmailSender
        $Script:varSMTPServer = $configuration.Configuration.SMTPServer
        $Script:varSaveReportsLocally = [System.Convert]::ToBoolean($configuration.Configuration.SaveReportsLocally)
        $Script:varReportName = $configuration.Configuration.ReportName
        $Script:varReportFolder = $configuration.Configuration.ReportFolder
        $Script:varServer = $configuration.Configuration.Server
        $Script:varUseCurrentUser = [System.Convert]::ToBoolean($configuration.Configuration.UseCurrentUser)
        $Script:varUser = $configuration.Configuration.User
        $Script:varScopeEnterprise = [System.Convert]::ToBoolean($configuration.Configuration.ScopeEnterprise)
        $Script:varScopeDAG = $configuration.Configuration.ScopeDAG
        $Script:varScopeServer = $configuration.Configuration.ScopeServer
        $Script:varScopeMailboxDatabase = $configuration.Configuration.ScopeMailboxDatabase
		$Script:varIncludeInherited = [System.Convert]::ToBoolean($configuration.Configuration.IncludeInherited)
		$Script:varExcludeZeroResults = [System.Convert]::ToBoolean($configuration.Configuration.ExcludeZeroResults)
		$Script:varExcludeUnknownUsers = [System.Convert]::ToBoolean($configuration.Configuration.ExcludeUnknownUsers)
		$Script:varFlagWarnings = [System.Convert]::ToBoolean($configuration.Configuration.FlagWarnings)
		$Script:varMailboxSizeWarning = $configuration.Configuration.MailboxSizeWarning
		$Script:varMailboxSizeAlert = $configuration.Configuration.MailboxSizeAlert
		$Script:varDeletedSizeWarning = $configuration.Configuration.DeletedSizeWarning
		$Script:varDeletedSizeAlert = $configuration.Configuration.DeletedSizeAlert
        $Script:varMailboxReportPermissions = @()
        foreach ($mboxoption in $configuration.Configuration.MailboxReportPermissions)
        {
            $hash = @{
				Option = $mboxoption.Option;
				Selected = [System.Convert]::ToBoolean($mboxoption.Selected)
			}			
            $Newobject = New-Object PSObject -Property $hash
            $Script:varMailboxReportPermissions += $Newobject
            switch ($mboxoption.Option) {
            	"Mailbox Summary Information" {
            		$Script:varSummaryReport = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
            	"Full Access Permissions" {
            		$Script:varFullAccessReport = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
            	"Send On Behalf Permissions" {
            		 $Script:varSendOnBehalfReport = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
                "Send As Permissions" {
            		 $Script:varSendAsReport = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
            	"Calendar Permissions" {
            		 $Script:varCalendarPermReport = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
                "Mailbox Rule - Forwarding" {
            		 $Script:varMailboxRuleForwarding = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
                "Mailbox Rule - Redirecting" {
            		 $Script:varMailboxRuleRedirecting = [System.Convert]::ToBoolean($mboxoption.Selected)
            	}
            }
		}
        $Script:varMailboxReportIgnoredUsers = @()
        foreach ($ignoreduser in $configuration.Configuration.MailboxReportIgnoredUser)
        {
            $Script:varMailboxReportIgnoredUsers += $ignoreduser.User
		}
        
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
 
    if ($SanitizedConfig)
	{
		# save the data
		[xml]$x=Get-SaveData
        foreach ($mboxperm in $varMailboxReportPermissions)
        {
            $newpermoption = $x.CreateElement("MailboxReportPermissions")
            $newpermoption.InnerXML = "<Option>$($mboxperm.Option)</Option><Selected>$($mboxperm.Selected)</Selected>"
            $x.Configuration.AppendChild($newpermoption)
		}
        foreach($ignoreduser in $varMailboxReportIgnoredUsers)
        {
			$ignoreduser2 = $x.CreateElement("MailboxReportIgnoredUser")
            $ignoreduser2.InnerXML = "<User>$($ignoreduser)</User>"
            $x.Configuration.AppendChild($ignoreduser2)
		}
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

Function Get-MailboxList {
    [CmdletBinding()] 
    param ( 
        [Parameter(HelpMessage="Return all mailboxes in a specific DAG.")]
        [String]$DAGName = '',
        [Parameter(HelpMessage="Return all mailboxes on a specific server.")]
        [String]$ServerName = '', 
        [Parameter(HelpMessage="Return all mailboxes in a specific mailbox.")]
        [String]$DatabaseName = '',
        [Parameter(HelpMessage="Return all mailboxes in the enterprise.")]
        [switch]$WholeEnterprise,
        [Parameter(HelpMessage="Return specific mailboxes.")]
        [string[]]$Mailboxes,
        [Parameter(HelpMessage="Returns only the mailbox identities instead of an array of mailbox objects.")]
        [switch]$ReturnIdentitiesOnly
    )
    BEGIN {
        $date = get-date -Format MM-dd-yyyy
    }
    PROCESS {
        $Results = @()
        if ($Mailboxes) {
            $MailboxInput = @()
            $MailboxInput += $Mailboxes
        }

        if ($Mailboxes.Count -ge 1) {
            Foreach ($Mbox in $MailboxInput) {
                try {
                    $_Mbox = Get-Mailbox $Mbox -ErrorAction 'Stop' -ResultSize Unlimited
                    $Results += $_Mbox
                }
                catch {
                    $time = get-date -Format hh.mm
                    $erroroutput = "$date;$time;$Mbox;$_"
                    Write-Warning $erroroutput
                }
            }
        }
        elseif ($DatabaseName -ne '') {
            try {
                $Results = @(Get-Mailbox -Database $DatabaseName -ResultSize Unlimited)
            }
            catch {
                $time = get-date -Format hh.mm
                $erroroutput = "$date;$time;$DatabaseName;$_"
                Write-Warning $erroroutput
            }
                  
        }
        elseif ($ServerName -ne '') {
            try {
                $Results += @(Get-Mailbox -Server $ServerName -ResultSize Unlimited)
            }
            catch {
                $time = get-date -Format hh.mm
                $erroroutput = "$date;$time;$ServerName;$_"
                Write-Warning $erroroutput
            }
        }
        elseif ($DAGName -ne '') {
            try {
                $Servers = @(Get-DatabaseAvailabilityGroup $DAGName | foreach {$_.Servers})
                foreach ($Server in $Servers) {
                    $Results += @(Get-Mailbox -Server $Server -ResultSize Unlimited)
                }
            }
            catch {
                $time = get-date -Format hh.mm
                $erroroutput = "$date;$time;$DAGName;$_"
                Write-Warning $erroroutput
            }
        }
        elseif ($WholeEnterprise) {
            $Results += @(Get-Mailbox -ResultSize Unlimited)
        }
        
        if ($ReturnIdentitiesOnly) {
            $Results = @($Results | %{[string]$_.Identity})
        }
        Return $Results 
    }
}

function Colorize-Table {
<# 
.SYNOPSIS 
Colorize-Table 
 
.DESCRIPTION 
Create an html table and colorize individual cells or rows of an array of objects based on row header and value. Optionally, you can also
modify an existing html document or change only the styles of even or odd rows.
 
.PARAMETER  InputObject 
An array of objects (ie. (Get-process | select Name,Company) 
 
.PARAMETER  Column 
The column you want to modify. (Note: If the parameter ColorizeMethod is not set to ByValue the 
Column parameter is ignored)

.PARAMETER ScriptBlock
Used to perform custom cell evaluations such as -gt -lt or anything else you need to check for in a
table cell element. The scriptblock must return either $true or $false and is, by default, just
a basic -eq comparisson. You must use the variables as they are used in the following example.
(Note: If the parameter ColorizeMethod is not set to ByValue the ScriptBlock parameter is ignored)

[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}

$args[0] will be the cell value in the table
$args[1] will be the value to compare it to

Strong typesetting is encouraged for accuracy.

.PARAMETER  ColumnValue 
The column value you will modify if ScriptBlock returns a true result. (Note: If the parameter 
ColorizeMethod is not set to ByValue the ColumnValue parameter is ignored)
 
.PARAMETER  Attr 
The attribute to change should ColumnValue be found in the Column specified. 
- A good example is using "style" 
 
.PARAMETER  AttrValue 
The attribute value to set when the ColumnValue is found in the Column specified 
- A good example is using "background: red;" 
 
.EXAMPLE 
This will highlight the process name of Dropbox with a red background. 

$TableStyle = @'
<title>Process Report</title> 
    <style>             
    BODY{font-family: Arial; font-size: 8pt;} 
    H1{font-size: 16px;} 
    H2{font-size: 14px;} 
    H3{font-size: 12px;} 
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
    TD{border: 1px solid black; padding: 5px;} 
    </style>
'@

$tabletocolorize = $(Get-Process | ConvertTo-Html -Head $TableStyle) 
$colorizedtable = Colorize-Table $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;"
$colorizedtable | Out-File "$pwd/testreport.html" 
ii "$pwd/testreport.html"

You can also strip out just the table at the end if you are working with multiple tables in your report:
if ($colorizedtable -match '(?s)<table>(.*)</table>')
{
    $result = $matches[0]
}

.EXAMPLE 
Using the same $TableStyle variable above this will create a table of top 5 processes by memory usage,
color the background of a whole row yellow for any process using over 150Mb and red if over 400Mb.

$tabletocolorize = $(get-process | select -Property ProcessName,Company,@{Name="Memory";Expression={[math]::truncate($_.WS/ 1Mb)}} | Sort-Object Memory -Descending | Select -First 5 ) 

[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
$testreport = Colorize-Table $tabletocolorize -Column "Memory" -ColumnValue 150 -Attr "style" -AttrValue "background:yellow;" -ScriptBlock $ScriptBlock -HTMLHead $TableStyle -WholeRow $true
$testreport = Colorize-Table $testreport -Column "Memory" -ColumnValue 400 -Attr "style" -AttrValue "background:red;" -ScriptBlock $ScriptBlock -WholeRow $true
$testreport | Out-File "$pwd/testreport.html" 
ii "$pwd/testreport.html"

.NOTES 
If you are going to convert something to html with convertto-html in powershell v2 there is a bug where the  
header will show up as an asterick if you only are converting one object property. 

This script is a modification of something I found by some rockstar named Jaykul at this site
http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

I believe that .Net 4.0 is a requirement for using the Linq libraries

.LINK 
http://www.the-little-things.net 
#> 
[CmdletBinding(DefaultParameterSetName = "ObjectSet")] 
param ( 
    [Parameter( Position=0,
                Mandatory=$true, 
                ValueFromPipeline=$true, 
                ParameterSetName="ObjectSet")]
        [PSObject[]]$InputObject, 
    [Parameter( Position=0, 
                Mandatory=$true, 
                ValueFromPipeline=$true, 
                ParameterSetName="StringSet")]
        [String[]]$InputString='', 
    [Parameter( Position=1 )]
        [String]$Column="Name", 
    [Parameter( Position=2 )] 
        $ColumnValue=0,
    [Parameter( Position=3 )]
        [ScriptBlock]$ScriptBlock = {[string]$args[0] -eq [string]$args[1]}, 
    [Parameter( Position=4, 
                Mandatory=$true )]
        [String]$Attr, 
    [Parameter( Position=5, 
                Mandatory=$true )]
        [String]$AttrValue, 
    [Parameter( Position=6 )]
        [Bool]$WholeRow=$false, 
    [Parameter( Position=7, 
                ParameterSetName="ObjectSet")] 
        [String]$HTMLHead='<title>HTML Table</title>',
    [Parameter( Position=8 )]
    [ValidateSet('ByValue','ByEvenRows','ByOddRows')]
        [String]$ColorizeMethod='ByValue'
    )
    
BEGIN 
{ 
    # A little note on Add-Type, this adds in the assemblies for linq with some custom code. The first time this 
    # is run in your powershell session it is compiled and loaded into your session. If you run it again in the same
    # session and the code was not changed at all powershell skips the command (otherwise recompiling code each time
    # the function is called in a session would be pretty ineffective so this is by design). If you make any changes
    # to the code, even changing one space or tab, it is detected as new code and will try to reload the same namespace
    # which is not allowed and will cause an error. So if you are debugging this or changing it up, either change the
    # namespace as well or exit and restart your powershell session.
    #
    # And some notes on the actual code. It is my first jump into linq (or C# for that matter) so if it looks not so 
    # elegant or there is a better way to do this I'm all ears. I define four methods which names are self-explanitory:
    # - GetElementByIndex
    # - GetElementByValue
    # - GetOddElements
    # - GetEvenElements
    $LinqCode = @"
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index)
    {
        return doc.Descendants(element)
                .Where  (e => e.NodesBeforeSelf().Count() == index)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value)
    {
        return  doc.Descendants(element) 
                .Where  (e => e.Value == value)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetOddElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 != 0)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetEvenElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 == 0)
                .Select (e => e);
    }
"@

    Add-Type -ErrorAction SilentlyContinue -Language CSharpVersion3 `
    -ReferencedAssemblies System.Xml, System.Xml.Linq `
    -UsingNamespace System.Linq `
    -Name XUtilities `
    -Namespace Huddled `
    -MemberDefinition $LinqCode
    
    $Objects = @() 
} 
 
PROCESS 
{ 
    $Objects += $InputObject 
} 
 
END 
{ 
    # Convert our data to x(ht)ml 
    if ($InputString)    # If a string was passed just parse it 
    { 
        $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")  
    } 
    else    # Otherwise we have to convert it to html first 
    { 
        $xml = [System.Xml.Linq.XDocument]::Parse("$($Objects | ConvertTo-Html -Head $HTMLHead)")
    } 
    
    switch ($ColorizeMethod) {
        "ByEvenRows" {
            $evenrows = [Huddled.XUtilities]::GetEvenElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
            foreach ($row in $evenrows)
            {
                $row.SetAttributeValue($Attr, $AttrValue)
            }            
        }

        "ByOddRows" {
            $oddrows = [Huddled.XUtilities]::GetOddElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
            foreach ($row in $oddrows)
            {
                $row.SetAttributeValue($Attr, $AttrValue)
            }
        }
        "ByValue" {
            # Find the index of the column you want to format 
            $ColumnLoc = [Huddled.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
            $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
    
            # Process each xml element based on the index for the column we are highlighting 
            switch([Huddled.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
            { 
                {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $ColumnValue))} {
                    if ($WholeRow)
                    {
                        $_.Parent.SetAttributeValue($Attr, $AttrValue)
                    }
                    else
                    {
                        $_.SetAttributeValue($Attr, $AttrValue)
                    }
                }
            }
        }
    }
    Return $xml.Document.ToString()
}
}

Function Get-CalendarPermission
{
<#
.Synopsis
    Retrieves a list of mailbox calendar permissions
.DESCRIPTION
    Get-CalendarPermission uses the exchange 2010 snappin to get a list of permissions for mailboxes in an exchange environment.
    As different languages spell calendar differently this script first pulls the actual name of the calendar by using
    get-mailboxfolderstatistics and has proven to work across multi-lingual organizations.
.PARAMETER Mailbox
    One or more mailbox names.
.PARAMETER LogErrors
    By default errors are not logged. Use -LogErrors to enable logging errors.
.PARAMETER ErrorLog
    When used with -LogErrors it specifies the full path and location for the ErrorLog. Defaults to "D:\errorlog.txt"
.LINK
    
.NOTES        
Name        :   Get Exchange Calendar Permissions
Last edit   :   April 14th 2013
Version     :   1.2.0 May 6 2013    :   Fixed issue where a mailbox name produces more than one mailbox
                1.1.0 April 24 2013 :   Used new script template from http://blog.bjornhouben.com
                1.0.0 March 10 2013 :   Created script

Author      :   Zachary Loeber
Website     :   http://www.the-little-things.net
Linkedin    :   http://nl.linkedin.com/in/zloeber
Keywords    :   Exchange, Calendar, Permissions Report
Disclaimer  :   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
                any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
                performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
                (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
                or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 

Copyright   :   I believe in sharing knowledge, so this script and its use is subject to : http://creativecommons.org/licenses/by-sa/3.0/

.EXAMPLE
    Get-CalendarPermission -MailboxName "Test User1" -LogErrors -logfile "C:\logfile.txt" -Verbose
    
    Description
    -----------
    Gets the calendar permissions for "Test User1", logs errors to "C:\myerrorlog.txt" and shows verbose information.
    
.EXAMPLE
    Get-CalendarPermission -MailboxName "user1","user2" -LogErrors -ErrorLog "C:\myerrorlog.txt" | Format-List
    
    Description
    -----------
    
    Gets the calendar permissions for "user1" and "user2", logs errors to "C:\myerrorlog.txt" and returns the info as a format-list.

.EXAMPLE
    (Get-Mailbox -Database "MDB1") | Get-CalendarPermission -LogErrors -Logfile "C:\myerrorlog.txt" | Format-Table Mailbox,User,Permission
    
    Description
    -----------
    Gets all mailboxes in the MDB1 database and pipes it to Get-CalendarPermission. Get-CalendarPermission logs errors to "C:\myerrorlog.txt" and returns the info as an autosized format-table containing the Mailbox,User, and Permission
#>
    [CmdletBinding()]
    param(
        [Parameter( Mandatory=$True,
                    ValueFromPipeline=$True,
                    Position=0,
                    HelpMessage="Enter an Exchange mailbox name")]
        [string[]]$MailboxName,
        [Parameter( Position=1,
                    HelpMessage='Enter the full path for your log file. By example: "C:\Windows\log.txt"')]
        [Alias("LogFile")]
            [String]$ErrorLog = ".\errorlog.txt",    
        [switch]$LogErrors
    )
    PROCESS
    {
        $Mboxes = @()
        $Mboxes += $MailboxName
        Foreach($Mailbox in $Mboxes)
        {
            TRY
            { 
                $Mbox = @(Get-Mailbox $Mailbox -erroraction Stop)
                $CheckSuccesful = $True
            }
            CATCH
            {
                $CheckSuccesful = $False
                $date = get-date -Format dd-MM-yyyy
                $time = get-date -Format hh.mm
                $erroroutput = "$date;$time;$Mailbox;MailboxError;$_.Exception.Message"

                Write-Warning $erroroutput

                IF($LogErrors -eq $True)
                {
                    Write-Verbose "Writing error for Mailbox : $Mailbox to $ErrorLog"
                    $ErrorOutput | Out-File $ErrorLog -Append -Encoding ASCII
                }
            }
             
            IF ($CheckSuccesful -eq $True) #If Mailbox was found keep processing
            {
                ForEach ($MailUser in $Mbox) {
                    # Construct the full path to the calendar folder regardless of the language
                    $Calfolder = $MailUser.Name
                    $Calfolder = $Calfolder + ':\'
                    $CalFolder = $Calfolder + [string](Get-MailboxFolderStatistics $MailUser.Identity -folderscope calendar).Name
                    $CalPerm = Get-MailboxFolderPermission $Calfolder
                    $Results = @()
                    foreach ($Perm in $CalPerm)
                    {
                        $TempHash = @{
                            'Mailbox'=$MailUser.Name;
                            'User'=$Perm.User;
                            'Permission'=$Perm.AccessRights;
                        }
                        $Tempobject = New-Object PSObject -Property $TempHash
                        $Results = $Results + $Tempobject
                    }
                }
                $Results
            }
        }
    }
}