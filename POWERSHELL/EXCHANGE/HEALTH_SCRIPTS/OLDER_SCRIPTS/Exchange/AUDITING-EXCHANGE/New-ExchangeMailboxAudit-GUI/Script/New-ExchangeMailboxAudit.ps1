Function New-ExchangeMailboxAuditReport {
    <#
    .SYNOPSIS
        Retrieves a list of various mailbox permissions and other settings one might use in an audit to determine
        where email may be leaking out of the organization or there are security risks.
        
    .DESCRIPTION
        Generate an HTML report with your choice of the follwing information:
            Full access permissions
            Send As permissions
            Send on behalf permissions
            Calendar permissions
            Forwarding Mailbox Rules
            Redirection Mailbox Rules
        
        Optionally you can also do the following:
            Report only on non-inherited permissions
            Include both a mailbox summary report with links to detailed subreports
            Generate just an email summary report
            Filter out specific users from  permissions reports
            Filter out unknown users from permissions reports
                
    .PARAMETER Mailboxes
        One or more mailbox names to process
    .PARAMETER DAGName
        Generate report on entire DAG
    .PARAMETER ServerName
        Generate report on entire database server
    .PARAMETER DatabaseName
        Generate report on entire mailbox database
    .PARAMETER WholeEnterprise
        Generate report on entire exchange organization
    .PARAMETER SendOnBehalfPermReport
        Generate send on behalf of permissions report
    .PARAMETER SummaryReport
        Generate summary report
    .PARAMETER SendAsPermReport
        Generate send as permissions report
    .PARAMETER FullPermReport
        Generate full permissions report
    .PARAMETER CalendarPermReport
        Generate calendar permissions report
    .PARAMETER ForwardingToReport
        Generate administratively set forwarding to report
    .PARAMETER ExcludeInheritedPerm
        Exclude inherited permissions where applicable
    .PARAMETER MailboxRuleForwardingReport
        Generate forwarding mailbox rules report
    .PARAMETER MailboxRuleRedirectingReport
        Generate redirecting mailbox rules report
    .PARAMETER ExcludeZeroResults
        Do not create mailbox report section if no sub-reports were generated
    .PARAMETER ExcludeUnknown
        Exclude unknown user accounts from permissions report (accounts starting with S-1-)
    .PARAMETER FlagWarnings
        Highlight cells based on my arbitrary cell highlighting rules
    .PARAMETER IgnoredUsersPermissions
        Ignore this array of user names in the permissions sub-reports
    .PARAMETER EmailRelay
        Email server to relay report through
    .PARAMETER EmailSender
        Email sender
    .PARAMETER EmailRecipient
        Email recipient
    .PARAMETER SendMail
        Send email of resulting report?
    .PARAMETER SaveReport
        Save the report?
    .PARAMETER ReportName
        If saving the report, what do you want to call it?
    .PARAMETER LogErrors
        By default errors are not logged. Use -LogErrors to enable logging errors.
    .PARAMETER ErrorLog
        When used with -LogErrors it specifies the full path and location for the ErrorLog. Defaults to "D:\errorlog.txt"
    .PARAMETER LoadConfiguration
        When set to be true the function attempts to load all parameter settings from an external configuration file.
    .PARAMETER ConfigurationFile
        Configuration file to use if LoadConfiguration is set to true. If the configuration file is not found the passed
        parameters are used instead.
    .LINK
        http://www.the-little-things.net
    .LINK
        http://nl.linkedin.com/in/zloeber
    .NOTES        
    Name            :   New-ExchangeMailboxAuditReport
    Author          :   Zachary Loeber
    Last edit       :   July 1st 2013
    Version         :   1.0.0 May 4th 2013        
                         - First release
                        1.0.1 June 13th 2013
                         - Fixed issues with mailbox statistics erring out if multiple mailboxes exist with the same alias
                         - Included better logic for summary only reporting
                         - Changed function name from Create-ExchangeMailboxAuditReport to New-ExchangeMailboxAuditReport
                         - Elminitated any default parameter options (as including them was just superfluous)
                         - Added Email subject parameters
                         - Added mailbox summary flag alert/warn trigger size parameters for deleted and total items in a mailbox
                         - Fixed some glaring variable name mistakes in the email delivery parameters
                         - Added parameters and logic for loading from saved config file (required for the upcoming GUI!
                        1.0.2 July 1st 2013
                         - Fixed incorrect default parameter state for full permissions report to be false
                         - Updated the Get-MailboxList function in globals.ps1 to spit out more friendly warnings instead of ugly red errors
                        1.0.3 July 17th 2013
                         - Added the view entire forest scope modifier at the begining of the script
                         - Refactored the xml configuration loading in globals.ps1
                         - Fixed a really crappy logic fault with reporting of inherited permissions when loading a config file
                         - Removed the folder selection for the report saving option in the GUI (it hangs on 2008 R2 for some silly reason
                         - Changed the GUI script name to align with the non-GUI script name
                         - Setup the ScriptPath variable to be locally set within either the GUI or non-GUI script instead of via the Globals.ps1 dot sourced file.

    Disclaimer      :   
        This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
        any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
        performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
        (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
        or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
    To improve      :   ???
    Copyright       :   I believe in sharing knowledge, so this script and its use is subject to : http://creativecommons.org/licenses/by-sa/3.0/

    .EXAMPLE
        $Report = New-ExchangeMailboxAuditReport -Mailboxes test.user1
        $Report | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"
        
        Description
        -----------
        Creates then immediately opens up a mailbox permissions report for test.user1 which includes the following default elements:
            - Summary Report
            - Full Mailbox Permissions
            - Send As Mailbox Permissions
            - Send On Behalf Mailbox Permissions
            - Calendar Permissions
            - Forwarding To Mailbox Rules
            - Redirecting To Mailbox Rules
        This will also take the following default actions
            - Ignore unknown users
            - Dieplay inherited permissions
            - Flag for warning any mailbox total size over 500Mb
            - Flag for alert any mailbox total size over 1024M
            - Flag for warning any mailbox deleted item total size over 500Mb
            - Flag for alert any mailbox deleted item total size over 1024M
            - Flag calendar permissions for default and anonymous users (with dashed blue outline)
            - Flag calendar permissions for warning the following permissions (default yellow background):
                Editor
                PublishingAuthor
                Author
                NonEditingAuthor
                PublishingEditor
                Reviewer

    .EXAMPLE
        $Report = New-ExchangeMailboxAuditReport -DAGName TEST-DAG1
        $Report | Out-File "$pwd/testreport.html" 
        ii "$pwd/testreport.html"
        
        Description
        -----------
        Does the same as the prior example but for all mailboxes in TEST-DAG1

    .EXAMPLE
        $MailUsers = @('test.user1','test.user2','test.user3')
        $Ignored_PermissionUsers = @( 'DOMAINNAME\Exchange Administrators',
                            'FORESTNAME\Domain Admins',
                            'FORESTNAME\Enterprise Admins',
                            'FORESTNAME\Organization Management',
                            'FORESTNAME\Exchange Servers',
                            'DOMAINNAME\Exchange Domain Servers',
                            'DOMAINNAME\Exchange Administrators',
                            'DOMAINNAME\Exchange Services',
                            'DOMAINNAME\BESAdmin',
                            'FORESTNAME\Organization Management',
                            'FORESTNAME\Exchange Trusted Subsystem',
                            'FORESTNAME\Administrator',
                            'NT AUTHORITY\SELF',
                            'NT AUTHORITY\SYSTEM')
        $Report1 = New-ExchangeMailboxAuditReport -DatabaseName 'MDB-01' -IgnoredUsersPermissions $Ignored_PermissionUsers
        $Report2 = New-ExchangeMailboxAuditReport -Mailboxes $Mailusers -IgnoredUsersPermissions $Ignored_PermissionUsers
        Description
        -----------
        Creates a report like the first two but ignoring most of the default permissions found in a forest with a sub-domain.
        One is created for a single database and the other is created for only 3 mailboxes.
        
    #>
    #region Parameters
    [CmdletBinding()]
    param
    (
        [Parameter( Position=0,
                    ValueFromPipeline=$true,
                    HelpMessage="Enter an Exchange mailbox name or an array of mailbox names")]
        [String[]]$Mailboxes,
        [Parameter( HelpMessage='Reporting by DAG')]
        [String]$DAGName='',
        [Parameter( HelpMessage='Reporting by server')]
        [String]$ServerName='', 
        [Parameter( HelpMessage='Reporting by Mailbox Database')]
        [String]$DatabaseName='',
        [Parameter( HelpMessage='Reporting by entire enterprise')]
        [Switch]$WholeEnterprise,
        [Parameter( HelpMessage="Generate send on behalf of permissions report")]
        [Switch]$SendOnBehalfPermReport,
        [Parameter( HelpMessage="Generate summary report")]
        [Switch]$SummaryReport,
        [Parameter( HelpMessage="Generate send as permissions report")]
        [Switch]$SendAsPermReport,
        [Parameter( HelpMessage="Generate full permissions report")]
        [Switch]$FullPermReport,
        [Parameter( HelpMessage="Generate calendar permissions report")]
        [Switch]$CalendarPermReport,
        [Parameter( HelpMessage="Generate administratively set forwarding to report")]
        [Switch]$ForwardingToReport,
        [Parameter( HelpMessage="Exclude inherited permissions where applicable")]
        [Switch]$ExcludeInheritedPerms,
        [Parameter( HelpMessage="Generate forwarding mailbox rules report")]
        [Switch]$MailboxRuleForwardingReport,
        [Parameter( HelpMessage="Generate redirecting mailbox rules report")]
        [Switch]$MailboxRuleRedirectingReport,
        [Parameter( HelpMessage="Generate mailbox delegation report")]
        [Switch]$MailboxDelegateReport,
        [Parameter( HelpMessage="Do not create mailbox report section if no sub-reports were generated")]
        [Switch]$ExcludeZeroResults,
        [Parameter( HelpMessage="Exclude unknown user accounts from permissions report (accounts starting with S-1-)")]
        [Switch]$ExcludeUnknown,
        [Parameter( HelpMessage="Highlight cells based on my arbitrary cell highlighting rules")]
        [Switch]$FlagWarnings,
        [Parameter( HelpMessage="Size of deleted items to flag with a warning.")]
        [Int]$DeletedSizeWarning = '512',
        [Parameter( HelpMessage="Size of deleted items to flag with an alert.")]
        [Int]$DeletedSizeAlert = '1024',
        [Parameter( HelpMessage="Size of total items to flag with a warning.")]
        [Int]$TotalSizeWarning = '512',
        [Parameter( HelpMessage="Size of total items to flag with an alert.")]
        [Int]$TotalSizeAlert = '1024',
        [Parameter( HelpMessage="Ignore this array of user names in the permissions sub-reports")]
        [String[]]$IgnoredUsersPermissions=@(),
        [Parameter( HelpMessage="Email server to relay report through")]
        [String]$EmailRelay = ".",
        [Parameter( HelpMessage="Email sender")]
        [String]$EmailSender='systemreport@localhost',
        [Parameter( HelpMessage="Email subject")]
        [String]$EmailSubject='Exchange Permission Report',
        [Parameter( HelpMessage="Email recipient")]
        [String]$EmailRecipient='default@yourdomain.com',
        [Parameter( HelpMessage="Send email of resulting report?")]
        [Switch]$SendMail,
        [Parameter( HelpMessage="Save the report?")]
        [Switch]$SaveReport,
        [Parameter( HelpMessage="If saving the report, what do you want to call it?")]
        [String]$ReportName=".\ExchangeMailboxPermissionReport.html",
        [Parameter( HelpMessage='Enter the full path for your log file. By example: "C:\Windows\log.txt"')]
        [String]$ErrorLog = ".\errorlog.txt",    
        [Parameter( HelpMessage='Set to true in order to log any errors.')]        
        [Switch]$LogErrors,
        [Parameter( HelpMessage='The name of a configuration file to load. This is ignored if the file does not exist.')]        
        [Switch]$LoadConfiguration,        
        [Parameter( HelpMessage='The name of a configuration file to load. This is ignored if the file does not exist.')]        
        [String]$ConfigurationFile='C:\LazyWinAdmin\RESULTS\Exchange\Config.xml'  
    )
#endregion Parameters

BEGIN 
{
    #region Globals
    Function Test-CommandExists
    {
        Param ($command)
        $oldPreference = $ErrorActionPreference
        $ErrorActionPreference = 'stop'
        try {
           if(Get-Command $command -ErrorAction Stop) {
              $true
           }
        }
        Catch {
           $false
        }
    }
    # Validate that the globals functions do not already exist. This is a quick hack to prevent
    #  double dot sourcing if the GUI is being used.
    if (!(Test-CommandExists 'Get-SaveData')) {
        . .\globals.ps1    
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
        [string]$ScriptDirectory = $PSScriptRoot
    }
    
    Set-ADServerSettings -ViewEntireForest $true

    #region Report Variables
    $ReportHead = 
@'
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"[]>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script src="../../../js/sorttable.js"></script>
<title>Mailbox Permissions Report</title>
'@

    # Modify this to change the entire layout of your report
    $ReportStyle = @'

    <style media="screen" type="text/css">
    /* <!-- */
    
    /* Summary Report Style */
    #mailboxsummaryreport {
        text-align: center;
		border-width: 1px;
        border-spacing: 1px;
        border-color: black;
        border-style: solid;        
        -moz-user-select:none;
        font-size : 77%;
        font-family : "Myriad Web",Verdana,Helvetica,Arial,sans-serif;
    }
    #mailboxsummaryreport tr td {
        text-align: center;
		border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        padding: 1px;
    }
    #mailboxsummaryreport tr {
        text-align: center;
		border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: black;
        padding: 1px;
    }
    #mailboxsummaryreport tr th.title {
        text-align: center;
		background-color: #696969;
        align: center;
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: black;
        color: #FFFFF0;
        padding: 1px;
    }
    #mailboxsummaryreport tr th {
        text-align: center;
		background-color: #878787;
        align: center;
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: black;
        color: #FFFFFF;
    }
    #mailboxsummaryreport tr.odd {
        text-align: center;
		border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        background-color: #E8E8E8;
        padding: 1px;
    }
    #mailboxsummaryreport tr.even {
		text-align: center;
		border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        background-color: #C8C8C8;
        padding: 1px;
    }
    #mailboxsummaryreport tr td.warn {
        text-align: center;
		border-width: 2px;
        border-spacing: 1px;
        background-color: yellow;
        border-color: black;
        border-style: dashed;
    }
    #mailboxsummaryreport tr td.alert {
        text-align: center;
		border-width: 2px;
        border-spacing: 1px;
        background-color: #CC3300;
        border-color: black;
        border-style: dashed;
        font-weight: bold;
        color: white;
    }
    
    #mailboxsubreport {
        border-width: 1px;
        border-spacing: 1px;
        border-color: white;
        border-style: solid;        
        -moz-user-select:none;
        font-size : 77%;
        font-family : "Myriad Web",Verdana,Helvetica,Arial,sans-serif;
        table-layout: fixed; 
        width: 100%;
    }
    #mailboxsubreport tr td.warn {
        border-width: 1px;
        border-spacing: 1px;
        background-color: yellow;
        border-color: black;
        border-style: dashed;
    }
    #mailboxsubreport tr.warn td {
        border-width: 1px;
        border-spacing: 1px;
        background-color: yellow;
        border-color: black;
        border-style: dashed;
    }
    #mailboxsubreport tr.odd {
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        background-color: #E8E8E8;
    }
    #mailboxsubreport tr.even {
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        background-color: #C8C8C8;
    }
    #mailboxsubreport tr.info {
        border-width: 2px;
        border-spacing: 1px;
        border-color: blue;
        border-style: dashed;
    } 
    #mailboxsubreport tr.info td{
        border-width: 2px;
        border-spacing: 1px;
        border-color: blue;
        border-style: dashed;
    } 
    #mailboxsubreport tr td {
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        word-wrap: break-word; !important
    }
    #mailboxsubreport tr {
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
    }
    #mailboxsubreport tr th.title {
        background-color: #696969;
        align: center;
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        color: #FFFFF0;
        padding: 1px;
    }
    #mailboxsubreport tr th {
        background-color: #878787;
        align: center;
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: white;
        color: #FFFFFF;
    }
    
    /* General styles */
    body {
        margin:0;
        padding:0;
        border:0;            /* This removes the border around the viewport in old versions of IE */
        width:100%;
        background:#fff;
        min-width:600px;        /* Minimum width of layout - remove line if not required */
                        /* The min-width property does not work in old versions of Internet Explorer */
        font-size:90%;
    }
    a {
        color:#369;
    }
    a:hover {
        color:#fff;
        background:#369;
        text-decoration:none;
    }
    h1, h2, h3 {
        margin:.8em 0 .2em 0;
        padding:0;
    }
    p {
        margin:.4em 0 .8em 0;
        padding:0;
    }
    img {
        margin:10px 0 5px;
    }
    /* Header styles */
    #header {
        clear:both;
        float:left;
        width:100%;
    }
    #header {
        border-bottom:1px solid #000;
    }
    #header p,
    #header h1,
    #header h2 {
        padding:.4em 15px 0 15px;
        margin:0;
        background-color: #878787;
        align: center;
        border-width: 1px;
        border-spacing: 1px;
        border-style: solid;
        border-color: black;
        color: #FFFFFF;
    }
    #header ul {
        clear:left;
        float:left;
        width:100%;
        list-style:none;
        margin:10px 0 0 0;
        padding:0;
    }
    #header ul li {
        display:inline;
        list-style:none;
        margin:0;
        padding:0;
    }
    #header ul li a {
        display:block;
        float:left;
        margin:0 0 0 1px;
        padding:3px 10px;
        text-align:center;
        background:#eee;
        color:#000;
        text-decoration:none;
        position:relative;
        left:15px;
        line-height:1.3em;
    }
    #header ul li a:hover {
        background:#369;
        color:#fff;
    }
    #header ul li a.active,
    #header ul li a.active:hover {
        color:#fff;
        background:#000;
        font-weight:bold;
    }
    #header ul li a span {
        display:block;
    }
    /* 'widths' sub menu */
    #layoutdims {
        clear:both;
        background:#eee;
        border-top:4px solid #000;
        margin:0;
        padding:6px 15px !important;
        text-align:right;
    }
    /* column container */
    .colmask {
        clear:both;
        float:left;
        width:100%;            /* width of whole page */
        overflow:hidden;        /* This chops off any overhanging divs */
    }
    /* common column settings */
    .colright,
    .colmid,
    .colleft {
        float:left;
        width:100%;
        position:relative;
    }
    .col1,
    .col2,
    .col3 {
        float:left;
        position:relative;
        padding:0 0 1em 0;
        overflow:hidden;
    }
    /* 2 Column (double page) settings */
    .doublepage {
        background:#eee;        /* right column background colour */
    }
    .doublepage .colleft {
        right:50%;            /* right column width */
        background:#fff;        /* left column background colour */
    }
    .doublepage .col1 {
        width:46%;            /* left column content width (column width minus left and right padding) */
        left:52%;            /* right column width plus left column left padding */
    }
    .doublepage .col2 {
        width:46%;            /* right column content width (column width minus left and right padding) */
        left:56%;            /* (right column width) plus (left column left and right padding) plus (right column left padding) */
    }
    /* Footer styles */
    #footer {
        clear:both;
        float:left;
        width:100%;
        border-top:1px solid #000;
    }
    #footer p {
        padding:10px;
        margin:0;
    }
    /* --> */
    </style>

</head>
<body>
'@

    $ReportFoot =
@'

</body>
</html>
'@

    $ColumnBreaker = @'

        </div><div class="col2">
'@

    # In the following variables I use some generic tokens which get replaced later on 
    # when constructing the reports.
    $TableHeader_MailboxSummary=
@'

<a name="top"></a>
<table id="mailboxsummaryreport" width="100%" class="sortable" id="header">
    
'@

    $TableHeader_MailboxPermissions = 
@'

<div id="header">
    <center><a name="@Name@"><h2>@Name@ ( @Email@ )</h2></a></center>
</div>
<div class="colmask doublepage">
    <div class="colleft">
        <div class="col1">
'@

    $TableHeaderOnly_MailboxPermissions = @'

<div id="header">
    <center><a name="@Name@"><h2>@Name@" ( @Email@ )</h2></a></center>
</div>
'@

    $TableFooter_MailboxPermissions = @'

        </div>
    </div>
</div>
<div id="footer">
<a href="#top">&#9650;</a><hr />
</div>
'@

    # @Width@ = Percentage of screen
    # @Colspan@ = Number of Properties
    # @Title@ = Report Title
    $TableHeader_SubReport = @'

        <table id="mailboxsubreport">
        <tr>
            <th class="title" colspan="@Colspan@">
                @Title@
            </th>
        </tr>
'@

    $TableFooter_SubReport = @'

</table>
'@
    #endregion Report Variables
    
    #endregion Globals
}
PROCESS 
{
    # Start by trying to load the configuration file
    if ($LoadConfiguration)
    {
        if (Load-Config)
        {
            $DAGName = $varScopeDAG
            $ServerName = $varScopeServer
            $DatabaseName = $varScopeMailboxDatabase
            $WholeEnterprise = $varScopeEnterprise
            $SendOnBehalfPermReport = $varSendOnBehalfReport
            $SummaryReport = $varSummaryReport
            $SendAsPermReport = $varSendAsReport
            $FullPermReport = $varFullAccessReport
            $CalendarPermReport = $varCalendarPermReport
            $ExcludeInheritedPerms = !$varIncludeInherited
            $MailboxRuleForwardingReport = $varMailboxRuleForwarding
            $MailboxRuleRedirectingReport = $varMailboxRuleRedirecting
            $ExcludeZeroResults = $varExcludeZeroResults
            $ExcludeUnknown = $varExcludeUnknownUsers
            $FlagWarnings = $varFlagWarnings
            $IgnoredUsersPermissions = $varMailboxReportIgnoredUsers
            $EmailRelay = $varSMTPServer
            $EmailSender = $varEmailSender
            $EmailSubject = $varEmailSubject
            $EmailRecipient = $varEmailRecipient
            $SendMail = $varEmailReport
            $SaveReport = $varSaveReportsLocally
            $ReportName = $varReportFolder + '\' + $varReportName
            $TotalSizeWarning = $varMailboxSizeWarning
            $TotalSizeAlert = $varMailboxSizeAlert
            $DeletedSizeWarning = $varDeletedSizeWarning
            $DeletedSizeAlert = $varDeletedSizeAlert
        }
    }
    $MBoxes = @()
    $MBoxes += @($Mailboxes)
    $USERPERM_IGNORE_LIST = @()
    $USERPERM_IGNORE_LIST += $IgnoredUsersPermissions
    
    $FullReport = ""
    $SummaryReportTable = ""
    $SummaryReportData = @()
    $FinalMailboxDetailReport = ""
    $TotalSubReports = 0        # Used to determine if we are really just running a summary-only report
    $SummaryOnlyReport = ((!$SendAsPermReport) -and 
                          (!$FullPermReport) -and 
                          (!$CalendarPermReport) -and 
                          (!$ForwardingToReport) -and 
                          (!$MailboxRuleForwardingReport) -and 
                          (!$MailboxRuleRedirectingReport))

    $MBoxes = @(Get-MailboxList -DAGName $DAGName `
                                -WholeEnterprise:$WholeEnterprise `
                                -ServerName $ServerName `
                                -DatabaseName $DatabaseName `
                                -Mailboxes $MBoxes)
    
    Foreach ($Mailbox in $MBoxes)
    {
        # All the arrays we may need to fill up based on what report we are generating.
        $SendOnBehalfPerms = @()
        $SendasPerms = @()
        $FullAccessPerms = @()
        $CalendarPerms = @()
        $ForwardingTo = ""
        $MailboxRuleForwarding = @()
        $MailboxRuleRedirecting = @()
        # Will come back to this later
        #$MailboxCalendarDelegates = @()
        $SubReportCount=0   # Used later to determine if detailed report section is generated
                            # (if it is still zero at the end and the $IgnoreZeroResults option
                            # was set then skip generation of the entire mailbox detail report for
                            # the mailbox.
        
        #region Data Gathering
        #region Summary
        if ($SummaryReport) {
            $CurrentMailboxStats = Get-MailboxStatistics $Mailbox.Identity
            
            # You have some creative capabilities here. If you do change things around, make certain
            # the report generation for the summary is also updated to reflect the changes
            # (specifically the color coding and the colspan number)
            $SummaryData = 
            @{
               'Name' = '<a href="#' + $Mailbox.Name + '">' + $Mailbox.Name + '</a>';
               'Last Logon' = $CurrentMailboxStats.LastLogonTime;
               'Last Logon Account' = $CurrentMailboxStats.LastLoggedOnUserAccount;
               'Primary SMTP' = $Mailbox.PrimarySmtpAddress;
               'Server' = $Mailbox.ServerName;
               'Database' = $Mailbox.Database;
               'Total Size (MB)' = $CurrentMailboxStats.TotalItemSize.Value.ToMB();
               'Total Items' = $CurrentMailboxStats.itemcount;
               'Total Deleted Size (MB)' = $CurrentMailboxStats.TotalDeletedItemSize.Value.ToMB();
               'Single Item Recovery' =	$Mailbox.SingleItemRecoveryEnabled;
               'Litigation Hold' = $Mailbox.LitigationHoldEnabled;
               'Retention Hold' = $Mailbox.RetentionHoldEnabled;
               'Audit Enabled' = $Mailbox.AuditEnabled;
            }
            $NewSummaryObject = New-Object PSObject -Property $SummaryData
            $SummaryReportData = $SummaryReportData + $NewSummaryObject
        }
        #endregion Summary
        
        #region Gather-Sendonbehalf
        if ($SendOnBehalfPermReport)
        {
            $sendbehalfperms=@($Mailbox | `
                
                select-object -expand grantsendonbehalfto | `
                select-object -expand rdn | `
                Sort-Object Unescapedname)
            $sendbehalfperms = @($sendbehalfperms | ?{($USERPERM_IGNORE_LIST -notcontains $_.Unescapedname)})

            if ($ExcludeUnknown -and ($sendbehalfperms.Count -ge 1))
            {                
                $sendbehalfperms = ($sendbehalfperms | ?{($_.Unescapedname -notlike "S-1-*")})
            }   
            if ($sendbehalfperms.Count -ge 1)
            {
                $SubReportCount++                    
            }
        }
        #endregion Gather-Sendonbehalf
        
        #region Gather-SendAs
        if ($SendAsPermReport)
        {
            $sendasperms=@(Get-ADPermission $Mailbox.identity | `
                ?{($_.extendedrights -like "*send-as*") -and `
                  ($USERPERM_IGNORE_LIST -notcontains $_.User)} | `
                  Select User,@{N="Inherited";E={$_.isInherited}},Deny)                
            
            if ($ExcludeInheritedPerms -and ($sendasperms.Count -ge 1))
            {
                $sendasperms = @($sendasperms | ?{($_.Inherited -like "false")})
            }
            if ($ExcludeUnknown -and ($sendasperms.Count -ge 1))
            {
                $sendasperms = @($sendasperms | ?{($_.User -notlike "S-1-*")})
            }
            if ($sendasperms.Count -ge 1)
            {
                $SubReportCount++
            }
        }
        #endregion Gather-SendAs
        
        #region Gather-FullAccessPerm
        if ($FullPermReport)
        {
            $fullaccessperms = @($Mailbox | `
                Get-MailboxPermission | `
                ?{($_.AccessRights -like "*fullaccess*") -and `
                  ($USERPERM_IGNORE_LIST -notcontains $_.User)} | `
                  Select User,IsInherited,Deny)            
            if ($ExcludeInheritedPerms -and ($fullaccessperms.Count -ge 1))
            {
                $fullaccessperms = @($fullaccessperms | ?{($_.isinherited -like "false")})
            }
            if ($ExcludeUnknown -and ($fullaccessperms.Count -ge 1))
            {
                $fullaccessperms = @($fullaccessperms | ?{($_.User -notlike "S-1-*")})
            }
            if ($fullaccessperms.Count -ge 1)
            {                
                $SubReportCount++
            }                
        }
        #endregion Gather-FullAccessPerm

        #region Gather-CalendarPerm
        if ($CalendarPermReport)
        {
            $CalendarPerms = @(Get-CalendarPermission $Mailbox.DistinguishedName | `
                               ?{($USERPERM_IGNORE_LIST -notcontains $_.User)} | `
                               select User,@{n="Permission";E={$_.Permission}})
                               
            $CalendarPerms = @($CalendarPerms | Where {$_.User -ne $Mailbox.Name})
            if ($CalendarPerms.Count -ge 1)
            {
                $SubReportCount++
            }
        }
        #endregion Gather-CalendarPerm
        
        #region Gather-ForwardingTo (In AD)
        if ($ForwardingToReport)
        {
            if ($Mailbox.ForwardingAddress -ne $null)
            {
                $ForwardingTo = Get-Mailbox -Identity $Mailbox.ForwardingAddress | Select $_.PrimarySMTPAddress
            }
        }
        #endregion Gather-ForwardingTo (In AD)
        
        #region Gather-MailboxRule Forwarding
        if ($MailboxRuleForwardingReport)
        {
            $MailboxRuleForwarding = @(Get-InboxRule -Mailbox $Mailbox.DistinguishedName | `
                                     where {($_.ForwardTo) -or ($_.ForwardAsAttachmentTo)} | `
                                     select Name,@{n="Forward To";E={$_.ForwardTo}},@{n="Forward As Attachment To";E={$_.ForwardAsAttachmentTo}})
            if ($MailboxRuleForwarding.Count -ge 1)
            {
                $SubReportCount++
            }
        }
        #endregion Gather-MailboxRule Forwarding

        #region Gather-MailboxRule Redirecting
        if ($MailboxRuleRedirectingReport)
        {
            $MailboxRuleRedirecting = @(Get-InboxRule -Mailbox $Mailbox.DistinguishedName | `
                                      where {$_.RedirectTo} | `
                                      select Name,@{n="Redirect To";E={$_.RedirectTo}})
            if ($MailboxRuleRedirecting.Count -ge 1)
            {
                $SubReportCount++
            }
        }
        #endregion Gather-MailboxRule Redirecting
        
#            # Will come back to this later (maybe)
#            #region Gather-Mailbox Calendar Delegates
#            if ($MailboxDelegateReport)
#            {
#                $MailboxCalendarDelegates = @(Get-CalendarProcessing -Identity $Mailbox.DistinguishedName | `
#                                              select identity -expand ResourceDelegates | `
#                                              select @{n="Delegates"; e={$_.Name}})
#                if ($MailboxCalendarDelegates.Count -ge 1)
#                {
#                    $SubReportCount++
#                }
#            }
#            #endregion Gather-Mailbox Calendar Delegates
          
        #endregion Data Gathering

        $MailboxDetailReport = ""
        $PermissionReportWidth = "100"
        #region Mailbox Report Contstruction
        # If we have any sub-report data then create the sub-report html tables and assign
        # them to a column based on how many elements they each have.
        if (!$SummaryOnlyReport)
        {
            if ($SubReportCount -ge 1)
            {
                $LeftColumnRows = 0
                $RightColumnRows = 0
                $LeftColumnTables = ""
                $RightColumnTables = ""
                
                ## Full Mailbox Permission Sub-report
                $FullAccessPermSubReport = ""
                if ($FullAccessPerms.count -ge 1)
                {
                    $CurrentSubReportCount++
                    $FullAccessPermHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', 3 `
                                                            -replace '@Title@', 'Full Access Permission'
                    $FullAccessPermBody = Colorize-Table $FullAccessPerms -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $FullAccessPermBody = Colorize-Table $FullAccessPermBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"
                    
                    if ($FlagWarnings)
                    {
                        $FullAccessPermBody = Colorize-Table $FullAccessPermBody -Column "Deny" -ColumnValue "True" -Attr "style" -AttrValue "color: red;" 
                        $FullAccessPermBody = Colorize-Table $FullAccessPermBody -Column "IsInherited" -ColumnValue "False" -Attr "style" -AttrValue "color: red;"
                    }                
                    
                    $FullAccessPermBody = [Regex]::Match($FullAccessPermBody, "(?s)(<colgroup>).+(?=</table>)").Value
                    $FullAccessPermSubReport = $FullAccessPermHead + $FullAccessPermBody + $TableFooter_SubReport
                    
                    # Determine which column our report goes into
                    $tmpTable = $FullAccessPermSubReport
                    $tmpRowCount = $FullAccessPerms.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }
                }
                
                ## Send As Permission Sub-report
                $SendAsPermSubReport = ""
                if ($sendasperms.Count -ge 1)
                {
                    $CurrentSubReportCount++
                    $SendAsPermHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', '3' `
                                                            -replace '@Title@', 'Send As Permission'
                    $SendAsPermBody = Colorize-Table $SendAsPerms -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $SendAsPermBody = Colorize-Table $SendAsPermBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"

                    if ($FlagWarnings)
                    {
                        $SendAsPermBody = Colorize-Table $SendAsPermBody -Column "Deny" -ColumnValue "True" -Attr "style" -AttrValue "color: red;" 
                        $SendAsPermBody = Colorize-Table $SendAsPermBody -Column "Inherited" -ColumnValue "False" -Attr "style" -AttrValue "color: red;"
                    }   
                    
                    $SendAsPermBody = [Regex]::Match($SendAsPermBody, "(?s)(<colgroup>).+(?=</table>)").Value                
                    $SendAsPermSubReport = $SendAsPermHead + $SendAsPermBody + $TableFooter_SubReport 
                    
                    # Determine which column our report goes into
                    $tmpTable = $SendAsPermSubReport
                    $tmpRowCount = $SendasPerms.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }
                }
                
                ## Send On Behalf Permission Sub-report
                $SendOnBehalfPermSubReport = ""
                if ($SendOnBehalfPerms.Count -ge 1)
                {
                    $CurrentSubReportCount++
                    $SendOnBehalfPermHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', '3' `
                                                            -replace '@Title@', 'Send On Behalf Permission'
                    $SendOnBehalfPermBody = Colorize-Table $SendOnBehalfPerms -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $SendOnBehalfPermBody = Colorize-Table $SendOnBehalfPermBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"

                    if ($FlagWarnings)
                    {
                        $SendOnBehalfPermBody = Colorize-Table $SendOnBehalfPermBody -Column "Deny" -ColumnValue "True" -Attr "style" -AttrValue "color: red;" 
                        $SendOnBehalfPermBody = Colorize-Table $SendOnBehalfPermBody -Column "Inherited" -ColumnValue "False" -Attr "style" -AttrValue "color: red;"
                    }   
                                    
                    $SendOnBehalfPermBody = [Regex]::Match($SendOnBehalfPermBody, "(?s)(<colgroup>).+(?=</table>)").Value
                    $SendOnBehalfPermSubReport = $SendOnBehalfPermHead + $SendOnBehalfPermBody + $TableFooter_SubReport
                    
                    # Determine which column our report goes into
                    $tmpTable = $SendOnBehalfPermSubReport
                    $tmpRowCount = $SendOnBehalfPerms.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }               
                }
                
                ## Calendar Permission Sub-report
                $CalendarPermSubReport = ""
                if ($CalendarPerms.Count -ge 1)
                {
                    $CurrentSubReportCount++
                    $CalendarPermHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', '2' `
                                                            -replace '@Title@', 'Calendar Permissions'
                    # I do this to be able to later insert <br/> tags where I want them...
                    foreach ($tmpCalPerm in $CalendarPerms)
                    {
                        $tmpCalPerm.Permission = [string]$tmpCalPerm.Permission -replace ' ','@br@'
                    }
                    
                    $CalendarPermBody = Colorize-Table $CalendarPerms -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $CalendarPermBody = Colorize-Table $CalendarPermBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"

                    if ($FlagWarnings)
                    {
                        $CalendarPermBody = Colorize-Table $CalendarPermBody -Column "User" -ColumnValue "Anonymous" -WholeRow $true -Attr "class" -AttrValue "info"
                        $CalendarPermBody = Colorize-Table $CalendarPermBody -Column "User" -ColumnValue "Default" -WholeRow $true -Attr "class" -AttrValue "info"

                        # Set cells to call out with warning colorization
                        $CalScriptBlockWarn = {
                                                   ([string]$args[0] -contains 'Editor') -or `
                                                   ([string]$args[0] -contains 'PublishingAuthor') -or `
                                                   ([string]$args[0] -contains 'Author') -or `
                                                   ([string]$args[0] -contains 'NonEditingAuthor') -or `
                                                   ([string]$args[0] -contains 'PublishingEditor') -or `
                                                   ([string]$args[0] -contains 'Reviewer')
                                              }

                        $CalendarPermBody = Colorize-Table $CalendarPermBody -ScriptBlock $CalScriptBlockWarn -Column "Permission" -ColumnValue "Reviewer" -Attr "class" -AttrValue "warn" 
                    }
                    
                    $CalendarPermBody = [Regex]::Match($CalendarPermBody, "(?s)(<colgroup>).+(?=</table>)").Value
                    $CalendarPermBody = $CalendarPermBody -replace '@br@','<br/>'
                    
                    $CalendarPermSubReport = $CalendarPermHead + $CalendarPermBody + $TableFooter_SubReport
                    
                    # Determine which column our report goes into
                    $tmpTable = $CalendarPermSubReport
                    $tmpRowCount = $CalendarPerms.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }
                }
                
                ## Mailbox Rule Forwarding Sub-report
                $MailboxRuleForwardingSubReport = ""
                if ($MailboxRuleForwarding.Count -ge 1)
                {
                    $CurrentSubReportCount++
                    $MailboxRuleForwardingHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', '3' `
                                                            -replace '@Title@', 'Mailbox Rule - Forwarding'
                    $MailboxRuleForwardingBody = Colorize-Table $MailboxRuleForwarding -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $MailboxRuleForwardingBody = Colorize-Table $MailboxRuleForwardingBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"
                    $MailboxRuleForwardingBody = [Regex]::Match($MailboxRuleForwardingBody, "(?s)(<colgroup>).+(?=</table>)").Value
                    
                    $MailboxRuleForwardingSubReport = $MailboxRuleForwardingHead + $MailboxRuleForwardingBody + $TableFooter_SubReport
                    
                    # Determine which column our report goes into
                    $tmpTable = $MailboxRuleForwardingSubReport
                    $tmpRowCount = $MailboxRuleForwarding.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }
                }
                
                ## Mailbox Rule Redirection Sub-report
                $MailboxRuleRedirectingSubReport = ""
                if ($MailboxRuleRedirecting.Count -ge 1)
                {
                    $CurrentSubReportCount++
                    $MailboxRuleRedirectingHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
                                                            -replace '@Colspan@', '2' `
                                                            -replace '@Title@', 'Mailbox Rule - Redirecting'
                    $MailboxRuleRedirectingBody = Colorize-Table $MailboxRuleRedirecting -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
                    $MailboxRuleRedirectingBody = Colorize-Table $MailboxRuleRedirectingBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"
                    $MailboxRuleRedirectingBody = [Regex]::Match($MailboxRuleRedirectingBody, "(?s)(<colgroup>).+(?=</table>)").Value
                    
                    $MailboxRuleRedirectingSubReport = $MailboxRuleRedirectingHead + $MailboxRuleRedirectingBody + $TableFooter_SubReport
                    
                    # Determine which column our report goes into
                    $tmpTable = $MailboxRuleRedirectingSubReport
                    $tmpRowCount = $MailboxRuleRedirecting.count
                    if ($LeftColumnRows -le $RightColumnRows)
                    {
                        if ($LeftColumnRows -gt 0) 
                        {
                            $LeftColumnTables = $LeftColumnTables + "<br/>"
                        }
                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
                        $LeftColumnTables = $LeftColumnTables + $tmpTable
                    }
                    else
                    {
                        if ($RightColumnRows -gt 0) 
                        {
                            $RightColumnTables = $RightColumnTables + "<br/>"
                        }
                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
                        $RightColumnTables = $RightColumnTables + $tmpTable
                    }
                }

# Maybe I'll come back to this later
#               $MailboxCalDelegatesSubReport = ""
#                if ($MailboxCalendarDelegates.Count -ge 1)
#                {
#                    $CurrentSubReportCount++
#                    $MailboxCalDelegatesHead = $TableHeader_SubReport -replace '@Width@', $PermissionReportWidth `
#                                                            -replace '@Colspan@', '1' `
#                                                            -replace '@Title@', 'Mailbox Calendar Delegation'
#                    $MailboxCalDelegatesBody = Colorize-Table $MailboxCalendarDelegates -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
#                    $MailboxCalDelegatesBody = Colorize-Table $MailboxCalDelegatesBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"
#                    $MailboxCalDelegatesBody = [Regex]::Match($MailboxCalDelegatesBody, "(?s)(<colgroup>).+(?=</table>)").Value
#                    
#                    $MailboxCalDelegatesSubReport = $MailboxCalDelegatesHead + $MailboxCalDelegatesBody + $TableFooter_SubReport
#                    
#                    # Determine which column our report goes into
#                    $tmpTable = $MailboxCalDelegatesSubReport
#                    $tmpRowCount = $MailboxCalendarDelegates.count
#                    if ($LeftColumnRows -le $RightColumnRows)
#                    {
#                        if ($LeftColumnRows -gt 0) 
#                        {
#                            $LeftColumnTables = $LeftColumnTables + "<br/>"
#                        }
#                        $LeftColumnRows = $LeftColumnRows + $tmpRowCount + 2
#                        $LeftColumnTables = $LeftColumnTables + $tmpTable
#                    }
#                    else
#                    {
#                        if ($RightColumnRows -gt 0) 
#                        {
#                            $RightColumnTables = $RightColumnTables + "<br/>"
#                        }
#                        $RightColumnRows = $RightColumnRows + $tmpRowCount + 2
#                        $RightColumnTables = $RightColumnTables + $tmpTable
#                    }
#                }
                
                $MailboxDetailReportHead = $TableHeader_MailboxPermissions -replace '@Name@', $Mailbox.Name `
                                    -replace '@Email@', $Mailbox.PrimarySMTPAddress `
                                    
                $MailboxDetailReport = $MailboxDetailReportHead + `
                                       $LeftColumnTables + `
                                       $ColumnBreaker + `
                                       $RightColumnTables + `
                                       $TableFooter_MailboxPermissions                       
            }            
            elseif ($ExcludeZeroResults -eq $false)
            {
                $MailboxDetailReportHead = $TableHeaderOnly_MailboxPermissions -replace '@Name@', $Mailbox.Name `
                                                -replace '@Email@', $Mailbox.PrimarySMTPAddress

                $MailboxDetailReportHead = $MailboxDetailReportHead
                $MailboxDetailReport = $MailboxDetailReportHead
            }
        }                
        $FinalMailboxDetailReport = $FinalMailboxDetailReport + $MailboxDetailReport
        #endregion Mailbox Report Construction
        $TotalSubReports = $TotalSubReports + $SubReportCount
    }
}
END 
{
    if ($SummaryReport -and ($SummaryReportData.count -gt 0))
    {           
        $SummaryReportData = $SummaryReportData | Select 'Name','Last Logon','Last Logon Account','Primary SMTP','Server','Database','Total Size (MB)','Total Items','Total Deleted Size (MB)','Single Item Recovery','Litigation Hold','Retention Hold','Audit Enabled' 
        
        $SummaryBody = Colorize-Table $SummaryReportData -ColorizeMethod "ByEvenRows" -Attr "class" -AttrValue "even"
        $SummaryBody = Colorize-Table $SummaryBody -ColorizeMethod "ByoddRows" -Attr "class" -AttrValue "odd"
        
        if ($FlagWarnings)
        {
            $SummaryBody = Colorize-Table $SummaryBody -Column 'Single Item Recovery' -ColumnValue "True" `
                           -Attr "style" -AttrValue "color: red;"
            $SummaryBody = Colorize-Table $SummaryBody -Column 'Litigation Hold' -ColumnValue "True" `
                           -Attr "style" -AttrValue "color: red;"
            $SummaryBody = Colorize-Table $SummaryBody -Column 'Retention Hold' -ColumnValue "True" `
                           -Attr "style" -AttrValue "color: red;"
            $SummaryBody = Colorize-Table $SummaryBody -Column 'Audit Enabled' -ColumnValue "True" `
                           -Attr "style" -AttrValue "color: red;"

            $SummaryScriptBlockTest = {[int]$args[0] -ge [int]$args[1]}
            $SummaryBody = Colorize-Table $SummaryBody `
                           -ScriptBlock $SummaryScriptBlockTest `
                           -Column 'Total Deleted Size (MB)' `
                           -ColumnValue $DeletedSizeWarning `
                           -Attr "class" `
                           -AttrValue "warn" 
            $SummaryBody = Colorize-Table $SummaryBody `
                           -ScriptBlock $SummaryScriptBlockTest `
                           -Column 'Total Deleted Size (MB)' `
                           -ColumnValue $DeletedSizeAlert `
                           -Attr "class" `
                           -AttrValue "alert"
            $SummaryBody = Colorize-Table $SummaryBody `
                           -ScriptBlock $SummaryScriptBlockTest `
                           -Column 'Total Size (MB)' `
                           -ColumnValue $TotalSizeWarning `
                           -Attr "class" `
                           -AttrValue "warn" 
            $SummaryBody = Colorize-Table $SummaryBody `
                           -ScriptBlock $SummaryScriptBlockTest `
                           -Column 'Total Size (MB)' `
                           -ColumnValue $TotalSizeAlert `
                           -Attr "class" `
                           -AttrValue "alert"
        }
        # Grab everything after the first table tag up to but not including the table closure tag
        # I keep the colgroup elements even though I don't even know what they are for....
        $SummaryBody = [Regex]::Match($SummaryBody, "(?s)(<colgroup>).+(?=</table>)").Value
        
        # Now insert our custom table header and close the table (Did you change the colspan to match
        # the number of elements you selected waaaayyyy earlier in the script?)
        $SummaryReportTable = ($TableHeader_MailboxSummary -replace '@colspan@', '13') + $SummaryBody + "</table><br/>"        
    }
    
    # Finallly create our report
    $FullReport = $ReportHead + `
                  $ReportStyle + `
                  $SummaryReportTable + `
                  $FinalMailboxDetailReport + `
                  $ReportFoot
    
    # XML turns our nice angle brackets which were embedded in to the table for links into proper html
    # this little hack turns them back into angle brackets again...
    $FullReport = $FullReport -replace '&lt;','<'
    $FullReport = $FullReport -replace '&gt;','>'
    
    if ($SaveReport)
    {
        $FullReport | Out-File $ReportName
    }
    if ($SendMail)
    {
        send-mailmessage -from $EmailSender -to $EmailRecipient -subject $EmailSubject `
        -BodyAsHTML -Body $FullReport -priority Normal -smtpServer $EmailRelay            
    }
    if ((!$SendMail) -and (!$SaveReport))
    {
        Return $FullReport
    }
}
}