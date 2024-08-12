<#
    .SYNOPSIS
    
        Email Organization Report is a great tool for IT Professionals who are working with Microsoft Exchange systems.

        This report will get organization wide information about your Email infrastructure, from Exchange servers O.S info, service health, up time details, beside Exchange and database highly aggregated information.

        Not only will you get a nice Dashboard describing your Exchange, you will get aggregated information and charts that you can present to your managers who care about aggregated, not detailed data.
        
        What you will get when running the script, is an overall overview about:

        1.	Exchange Servers in your organization
                [Break down by AD Site]
                [Info about External and Internal Web services names]
                [Mailbox Count per AD Site]
                    •	O.S version
                    •	O.S Service Pack level.
                    •	Exchange Service health.
                    •	Up time in days.
                    •	Exchange Version.
                    •	Exchange Service Level.
                    •	Exchange Rollup Update Information.
                    •	Exchange Role(s).
                    •	Number of mailboxes in case of MBX role.

        2.	Database Full Inventory
                [Break down by DAG]
                [Separate Table for Recovery DBs]
                [Separate Table for Non DAG DBs]
                [Table per DAG]
                    •	DB Name.
                    •	Server Location.
                    •	Mailbox Count.
                    •	Average Mailbox Size.
                    •	Archive Mailbox Count.
                    •	Average Archive Mailbox Size.
                    •	Mount Status.
                    •	DB Size.
                    •	Storage Group Name (Pre E2010).
                    •	White Space.
                    •	Circular Logging.
                    •	DB Disk Free Percentage.
                    •	Log Disk Free Percentage.
                    •	Last Full Backup Date.
                    •	Backed up Since (Days) – with customized thresholds.
                    •	Quota Info: Prohibit Send.
                    •	Quota Info: Prohibit Send and Receive.
                    •	DB Activation Preference Check [Is it mounted on the preferred Server?].
                    •	DB Copy Location and Activation Preference assignment.

        3.	Mailbox Type Aggregated Data:
                    •	User Mailbox Count.
                    •	Shared Mailbox Count.
                    •	Room Mailbox Count.
                    •	Discovery Mailbox Count.

        4.	Exchange Server Aggregated Data
                    •	Total Number of Exchange Servers break down by
                        i.	Version.
                        ii.	Role.

        5.	Mailbox Aggregated Data
                [Overall Statistics for the Organization]
                    •	Total Mailbox Count.
                    •	Total databases Size                   
                    •	Total Archive Count.
                    •	Total Archive Size.
                    •	Average Archive Size.

        6.	Database Activation Preference Map Table
                [Table per DAG]
                    •	List of all DAG DBs
                    •	Information about Database Copies for each DB.
                    •	Information about Activation Preference for each Copy (MBX info).
                    •	Total number of DB Mounted Per DAG Server.
                    •	Total number of DB Copies per DAG Server.
                    •	Ideal number of DBs mounted for each server using Activation Preference.

        7.	Charts and Diagrams
                    •	Chart : DB Vs Mailbox Count.
                    •	Chart : DB Vs Size in GB.
                    •	Chart : DB Vs Backed Up Since (Days).
                    •	Chart : Server Vs DB Count.
                    •	Chart : Server Vs Mailbox Count.

        8.	Script Filter Finally!!
                [This has been for a long time pending feature]
                    •	Filer by comma separated Exchange Servers.
                            [Example: “Srv1 , Srv2, Srv3,……”]
                    •	Filter by Expression.
                            [Example: “*LON”]
                            [This will get all Exchange Servers with names ending with “LON”
                            [Example: “uk*”]
                            [This will get all Exchange Servers with names started with “uk”
                    •	Filter by list of DAG names
                            [Example: “Dag1,Dag2,….”
                            [This will get all Exchange Servers member of those DAGs and will narrow the script scope to those DAGs only.

        9.	Script Log Files
                [Used for logging script actions and steps taken]
                    •	Information Log File.
                [Records script progress, actions, and gives you insights about what data being collected]
                    •	Detail Log File.
                [Record detailed information for every single information detected for Exchange servers and databases]
                    •	Error Log File
                [Records any termination errors along with all information about the error internal exception messages and command failing with line number info]

        10.	Totally new output screen when running the script interactively:
                    •	New Time Watch to record the script execution time.
                    •	Nested on screen progress bar to give you an idea about how long the script will run.
                    •	On screen step by step statistics, so you can get quicker info while the script is running.
                    •	Verbose Mode for advance detailed information on the screen.
                    •	At the end of the script, couple of information will be displayed to list you the log files created with their location, and aggregate information and statistics about your Exchange Environment.
                    •	Send Email option.

        11.	PowerShell Remoting to get WMI Data.
             [New switch in the script to enable the use of PowerShell remoting instead of RPC legacy calls. You can use this switch if you enabled your Exchange servers for PS Remoting. The huge benefit for doing this is simply security and reliability. The script in this mode will try to test WS-MAN connectivity and version for each remote computer and then decide how to get WMI data. If for any reason PS remoting is not available on a remote computer, the script will fall back to legacy normal RPC calls to ensure reliability of getting the data].

        12.	Totally New Chart PowerShell Wrapper
                [The script is armed with totally new reliable Chart engine “code name: Get-CorpChart_Light” that will do smart calculation to change the chart dimensions and size depending on the number of input data. You will no longer get so crowded charts filled with letters that you cannot read. The chart will use a smart algorithm to calculate the best dimension of the graph depending on the number of input data].

        13.	Better Error and Exception Handling.
                [The script uses complicated logic to track exception and errors, classifying them to categories, logging them to different log files using an internal function called “Write-CorpError”. After that, in most cases, the error will be logged, execution will continue without bothering you with on screen bad error messages. If the error affects the execution of the script, a nice and informative message will be displayed on the screen and log files with suggestions if possible to how to solve it].

        14.	Works with all Exchange versions.
             [This was a limitation on the previous script that only work with Exchange 2010]

        15.	Smart Dashboard HTML Table with customized thresholds. 
                [You can configure couple of thresholds for DB Backup days and Server up time data, which is how many days since the server restarted. When the threshold is crossed, a color change will happen for affected data cells, grapping your attention to what matters quickly.].

        16.	Script divided into Modules with lots of regions
                [To better understand the code, the script is divided into the below modules to better organize and browse through the script.
        -	Module 1 : Customization
                [Thresholds for you to customize]
        -	Module 2 : Functions
                [Script functions]
        -	Module 3 : Factory
                [Preparation tasks like creating log files and variable initialization]
        -	Module 4 : Process
                [Getting the data]
        -	Module 5 : Output
                [Outputting the data]
        -	Module 6 : Charts
                [Drawing the data]
        -	Module 7 : Final Tasks
                [Sending email and closing log files]
      

        17.	 Joined effort writing the code.
                [Some of the script internal functionalities, specifically getting structured data part, is written by Steve Goodman, a Microsoft MVP at the time of writing this script. After communicating with Steve, who is a wonderful professional, his internal functions that gets structured internal data, are used as a base for writing the Process Module of the script. Big thanks to Steve on helping writing this module. I also used couple of helper functions from the online community like the Log Functions. Copyrights are mentions in the code.]


        
    --------------
    Script Info
    --------------

        Script Name                :         Email Organization Report (Get-CorpEmailReport)
        Script Version             :         2.4.9
        Author                     :         Ammar Hasayen   
        Blog                       :         http://ammarhasayen.com 
        Description                :         Generate Exchange Organization Email Report
        Twitter                    :         @ammarhasayen
        Email                      :         me@ammarhasayen.com

    --------------
    Copy Rights
    --------------

        Some of the script internal functionalities, specifically getting structured data part, is written by Steve Goodman,
        a Microsoft MVP at the time of writing this script. After communicating with Steve, who is a wonderful professional,
        his internal functions that gets structured internal data, are used as a base for writing the Process Module of the script.
        Big thanks to Steve on helping writing this module. 
        [Steve Goodman's Exchange Environemnt Report (version 1.5.8) Published February 2,2014]
           
        I also used couple of helper functions from the online community like the Log Functions [http://9to5it.com/] and the 
        Write-CorpError (based on script code from PowerShell Deep Dives Book)

        All copy rights are mentions in the code itself.


     --------------
    Versions
    --------------

     Version 2.4.9 published April 2015:
        Fixes included:
        - Fixing loading PowerShell on Exchange 2013 server.
        - Fixing loading PowerShell when the Exchange install directory is not C:\
        - Fixing Exchange 2013 server version reporting


     .LINK
     My Blog
     http://ammarhasayen.com

     .LINK
     Steve Goodman blog @stevegoodman
     http://www.stevieg.org 

     .LINK
     Luca Blog - author of the Start-Log, Write-Log and Finish-Log helper functions in the script
     http://9to5it.com/powershell-logging-function-library/      



     .DESCRIPTION

         - Account Requirements

            - Local Administrator on the Exchange Servers to get WMI data like disk info and service health state
            - Exchange View Only Administrator Role

         - Software Requirements
          
            - Microsoft Chart Controls for Microsoft .NET Framework, to be installed just on the machine running the script.
              This is used only to draw the charts. If it is not installed, the script will simply continue to run normally
              and will skip the chart part.
              Download Microsoft Chart Controls for Microsoft .NET Framework 3.5 (http://www.microsoft.com/en-us/download/details.aspx?id=14422)
            - You must run the script from Exchange Server and not from your admin machine, even if you have the Exchange PowerShell
              commands available via remoting. 

         - Trick:

             In order to get remote data from EDGE server using WMI like disk info and UpTime details, you have to enable (LocalAccountTokenFilterPolicy)
             registry key on the EDGE server and setting it to 1. 
                Key: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System
                Value: LocalAccountTokenFilterPolicy
                Data: 1 (to disable, 0 enables filtering)
                Type: REG_DWORD (32-bit)
             Then you have to create an administrator account on the EDGE with the same accountname and password as the account running this script.
             More info about (LocalAccountTokenFilterPolicy) is on my blog post here http://wp.me/p1eUZH-pZ

             

     .PARAMETER ScriptFilesPath
     Path to store script files like ".\" to indicate current directory or full path like C:\myfiles
	
	.PARAMETER SendMail
	 Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
	
	.PARAMETER MailFrom
	 Email address to send from. Passed directly to Send-MailMessage as -From
	
	.PARAMETER MailTo
	 Email address to send to. Passed directly to Send-MailMessage as -To
	
	.PARAMETER MailServer
	 SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer
	    
	.PARAMETER ViewEntireForest
	 By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.
   
    .PARAMETER OnlyIncludedServers
	 Filter to narrow the scope of script execution. String Array. Type the name of Exchange Servers to include in the script, seperated by comma (eg Ex1, Ex2,..)

    .PARAMETER ServerFilter
	 Filter to narrow the scope of script execution. Use a text based string to filter Exchange Servers by, e.g. NL-* -  Note the use of the wildcard (*) character to allow for multiple matches.

    .PARAMETER InputDAGs
	 Filter to narrow the scope of script execution. String Array. Type name of DAGs to include in the script seperated by comma (eg DAG1,DAG2,...)

     .PARAMETER WMIRemoting
	 Switch to instruct the script to use PowerShell Remoting first to get WMI data from Exchange servers, and then fall back to RPC.
    
	
    
    .EXAMPLE
     Generate the HTML report and supplying the current directry as a script path to create output files
    .\Get-CorpEmailReport.ps1 -ScriptFilesPath .\

    .EXAMPLE
    Generate the HTML report and supplying the custom directory as a script path to create output files
    .\Get-CorpEmailReport.ps1 -ScriptFilesPath C:\MyFiles

    .EXAMPLE
    Generate the HTML report and Filter by servers that start with "NL"
    .\Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -ServerFilter "NL*"

    .EXAMPLE
    Generate the HTML report and Filter by including only Ex1 and Ex2 servers
    .\Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -OnlyIncludedServers Ex1,Ex2

    .EXAMPLE
    Generate the HTML report and Filter by including only Servers that are member of a DAG called "DAG1"
    .\Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -InputDAGs DAG1

    .EXAMPLE
     Generate the HTML report and use PowerShell Remoting for WMI data collection
     \Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -WMIRemoting

    .EXAMPLE
     Generate the HTML report with SMTP Email option
     \Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -SendMail:$true -MailFrom noreply@contoso.com  -MailTo me@contoso.com  -MailServer smtp.contoso.com

     .EXAMPLE
     Generate the HTML report with disabling ViewEntireForest option
     \Get-CorpEmailReport.ps1 -ScriptFilesPath .\  -ViewEntireForest:$false
   #>



#region parameters

    #Quote : Script block derived from Steve Goodman script
    [cmdletbinding(DefaultParameterSetName="ServerFilter")]

    param(
        [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Path to store script files like c:\ ')][string]$ScriptFilesPath,
	    [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Send Mail ($True/$False)')][bool]$SendMail=$false,
	    [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail From')][string]$MailFrom,
	    [parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail To')]$MailTo,
	    [parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Server')][string]$MailServer,	
	    [parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Change view to entire forest')][bool]$ViewEntireForest=$true,
	    [parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Server Name Filter (eg NL-*)',ParameterSetName="ServerFilter")][string]$ServerFilter="*",
        [parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Only include those DAGs (eg DAG1,DAG2...)',ParameterSetName="DAGFilter")][array]$InputDAGs,
        [parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Only include those Servers (eg Server1,Server2...)',ParameterSetName="IncludeServerFilter")][array]$OnlyIncludedServers,
        [switch]$WMIRemoting
        )
     #End Quote : Script block derived from Steve's script

#endregion parameters



<#
 Script Structure :

    - Module 1 : Customization
        Here you can go and play with thresholds that will affect the script action and the output shape
    - Module 2 : Functions
        Here is the section where helpers and functions are placed
    - Module 3 : Factory
        This is the part where the script initialize and start creating log files and initialize variables
    - Module 4 : Process
        Collected data from the Exchange Environment
    - Module 5 : Output
        Outping the data to HTML Tables
    - Module 6 : Charts
        Outping the data to charts (PNG files)
    - Module 7 : Final Tasks
        Send email if configured, and closing log files
 #>


#++++++++++++++++++++++++++++++++++++ Module 1 : Customization ++++++++++++++++++++++++++++++++++++



#region Module 1 : Customization

    #Threshold for days since last database backup (warning threshold) . Color warning will show otherwise
	[int]$BackupWarning = 1000

    #Threshold for days since last database backup (Error threshold) . Color warning will show otherwise
	[int]$BackupError = 2000

    #Threshold for days since Exchange Servers have been restarted (Numbers represents days) . Color warning will show otherwise
    [int]$UptimeErrorThreshold   = 600	                  

    #Coloring
    $greenColor    = "#00FF00"
    $warningColor  = "#FF9900"
    $errorcolor    = "#980000"

#endregion Module 1 : Customization



#++++++++++++++++++++++++++++++++++++   Module 2 : Function    ++++++++++++++++++++++++++++++++++++



#region Module 2 : Functions

    #region helper functions

        function get-timestamp {

           get-date -format 'yyyy-MM-dd HH:mm:ss'

         } # function get-timestamp


        function Log-Start{
            <#
            .SYNOPSIS
            Creates log file
 
            .DESCRIPTION
            Writes initial logging data
 
            .PARAMETER $LogFullPath
            Mandatory. File name and path name to log file. Eaxmple : C:\temp\myfile.log
  
 
            .INPUTS
            Parameters above
 
            .OUTPUTS
            Log file created
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Note:           modified by Ammar Hasayen
 
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
 
            .EXAMPLE
            Log-Start -LogFullPath "C:\Windows\Temp\mylog.log"
            #>
    
            [CmdletBinding()]
  
            Param ([Parameter(Mandatory=$true)][string]$LogFullPath)
   
            Process{    
   
    
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value "Started processing at [$([DateTime]::Now)]."
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value ""
   
  
            #Write to screen for debug mode
            Write-Debug "***************************************************************************************************"
            Write-Debug "Started processing at [$([DateTime]::Now)]."
            Write-Debug "***************************************************************************************************"
            Write-Debug ""    
            }
        } # function Log-Start


        function Log-Write{
            <#
            .SYNOPSIS
            Writes to a log file
 
            .DESCRIPTION
            Appends a new line to the end of the specified log file
  
            .PARAMETER LogFullPath
            Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
            .PARAMETER LineValue
            Mandatory. The string that you want to write to the log
      
            .INPUTS
            Parameters above
 
            .OUTPUTS
            None
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Purpose/Change: Initial function development
            Note:           Modified by Ammar Hasayen
  
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
 
            .EXAMPLE
            Log-Write -LogFullPath "C:\Windows\Temp\Test_Script.log" -LineValue "This is a new line which I am appending to the end of the log file."
            #>
  
          [CmdletBinding()]
  
          Param ([Parameter(Mandatory=$true)][string]$LogFullPath, [Parameter(Mandatory=$true)][string]$LineValue)
  
    
          Process{
    
    
            Add-Content -Path $LogFullPath -Value $LineValue
  
            #Write to screen for debug mode
            Write-Debug $LineValue

          }

    } # function Log-Write


        function Log-Finish{
            <#
            .SYNOPSIS
            Write closing logging data & exit
 
            .DESCRIPTION
            Writes finishing logging data to specified log and then exits the calling script
  
            .PARAMETER LogFullPath
            Mandatory. Full path of the log file you want to write finishing data to. Example: C:\Windows\Temp\Test_Script.log
 
            .INPUTS
            Parameters above
 
            .OUTPUTS
            None
 
            .NOTES
            Version:        1.0
            Author:         Luca Sturlese
            Creation Date:  10/05/12
            Purpose/Change: Initial function development
            Note:           Modified by Ammar Hasayen
    
            Version:        1.1
            Author:         Luca Sturlese
            Creation Date:  19/05/12
            Purpose/Change: Added debug mode support
  
            Version:        1.2
            Author:         Luca Sturlese
            Creation Date:  01/08/12
            Purpose/Change: Added option to not exit calling script if required (via optional parameter)
 
            .EXAMPLE
            Log-Finish -LogFullPath "C:\Windows\Temp\Test_Script.log" 

            #>
  
            [CmdletBinding()]
  
            Param ([Parameter(Mandatory=$true)][string]$LogFullPath)
  
            Process{
            Add-Content -Path $LogFullPath -Value ""
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
            Add-Content -Path $LogFullPath -Value "Finished processing at [$([DateTime]::Now)]."
            Add-Content -Path $LogFullPath -Value "***************************************************************************************************"
  
            #Write to screen for debug mode
            Write-Debug ""
            Write-Debug "***************************************************************************************************"
            Write-Debug "Finished processing at [$([DateTime]::Now)]."
            Write-Debug "***************************************************************************************************"
  
       
             }
        } # function Log-Finish


        function Write-CorpError {
            
            [cmdletbinding()]

            param(
                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Error Variable')]$myError,	
	            [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Additional Info')][string]$Info,
                [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Log file full path')][string]$mypath,
	            [switch]$ViewOnly

                )

                Begin {
       
                    function get-timestamp {

                        get-date -format 'yyyy-MM-dd HH:mm:ss'
                    } 

                } #Begin

                Process {

                    if (!$mypath) {

                        $mypath = " "
                    }

                    if($myError.InvocationInfo.Line) {

                    $ErrorLine = ($myError.InvocationInfo.Line.Trim())

                    } else {

                    $ErrorLine = " "
                    }

                    if($ViewOnly) {

                        Write-warning @"
                        $(get-timestamp)
                        $(get-timestamp): $('-' * 60)
                        $(get-timestamp):   Error Report
                        $(get-timestamp): $('-' * 40)
                        $(get-timestamp):
                        $(get-timestamp): Error in $($myError.InvocationInfo.ScriptName).
                        $(get-timestamp):
                        $(get-timestamp): $('-' * 40)       
                        $(get-timestamp):
                        $(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)
                        $(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)
                        $(get-timestamp): Command: $($myError.invocationInfo.MyCommand)
                        $(get-timestamp): Line: $ErrorLine
                        $(get-timestamp): Error Details: $($myError)
                        $(get-timestamp): Error Details: $($myError.InvocationInfo)
"@

                        if($Info) {
                            Write-Warning -Message "More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Write-Warning -Message "Error Inner Exception: $($myError.Exception.InnerException.Message)"
                        }

                        Write-warning -Message " $('-' * 60)"

                     } #if($ViewOnly) 

                     else {
                     # if not view only 
        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): $('-' * 60)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):  Error Report"        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error in $($myError.InvocationInfo.ScriptName)."        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Command: $($myError.invocationInfo.MyCommand)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line: $ErrorLine"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError.InvocationInfo)"
                        if($Info) {
                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp): More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp) :Error Inner Exception: $($myError.Exception.InnerException.Message)"
            
                        }    

                     }# if not view only

               } # End Process

        } # function Write-CorpError


        function Test-corpIsWsman{

            <#.Synopsis
            Tests if a computer is running WS-MAN and verifies the version of WS-MAN if found.

            .Description
            Test-corpIsWsman takes a computer account, returns information about the support for WSMAN and verifies the version of WS-MAN if found.

            Returns object with two properties :

                - "wsman_supported" is a Boolean that can be 
                    > $true : if the computer is accessible over wsman
                    > $false : if the computer is not accessible over wsman

                - "wsman_version" is a string that can be 
                * "None"      : if we cannot reach the computer using WSMAN
                * "wsman v2" : if we can reach the computer  using WSMAN, and it is version 2.0
                * "wsman v3" : if we can reach the computer  using WSMAN, and it is version 3.0



            Versioning:
                - Version 1.0 written 5  November 2013 : returns an integer that can be 0 "no wsman" , 2 "wsman v2" , 3 "wsman v3"
                - Version 2.0 written 13 November 2013 : returns an object with two properties "wsman_version" and "Wsman_supported" 



            .PARAMETER Computername
            String value representing the computer to test. You can use the following aliases for this parameter
            "Computer","Name","MachineName"

            .Example
            Getting support for WS-MAN on PC1
            PS C:\>Test-corpIsWsman -ComputerName "PC1"

            .Example
            Getting support for WS-MAN on PC1 using -Name parameter alias
            PS C:\>Test-corpIsWsman -Name "PC1"

            .Example
            Getting support for WS-MAN on PC1 using -Computer parameter alias
            PS C:\>Test-corpIsWsman -Computer "PC1"

            .Example
            Running the function without any parameters will default to the localhost as computername
            PS C:\>Test-corpIsWsman

            .Example
            Running the function with error action SilentlyContinue. You will still get an object back if something went wrong like computer does not exist. The function will throw exception always to allow you to catch it if you want, or you can use EA (SilentlyContinue) to get the object result back without noticing the exception.
            PS C:\>Test-corpIsWsman  -Computer "PC1" -EA SilentlyContinue

            .Example
            Get computer names from text file and pipeline the output to Test-corpIsWsman
            PS C:\> get-content computers.txt | Test-corpIsWsman

            .Example
            Get-ADComputer will produce "Name" property, convert it to "ComputerName" so it can be used accross the pipeline
            PS C:\> Get-ADcomputer PC1 |select @{Name="ComputerName";Expression={$_.Name}} |Test-corpIsWsman

            .Notes
            Last Updated             : Nov 13, 2013
            Version                  : 2.0 
            Author                   : Ammar Hasayen (@ammarhasayen)
            based on                 : Jeffery Hicks script(@JeffHicks)

            .Link
            http://ammarhasayen.com
            .Link
            http://jdhitsolutions.com/blog/2013/04/get-ciminstance-from-powershell-2-0

            #>

            [cmdletbinding()]

                Param(

                [Parameter(Position = 0,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]

                [alias("name","machinename","computer")]

                [string]$Computername=$env:computername
                )


                Begin {
        
                    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
          
                    #a regular expression pattern to match the ending
                    #gets last digit, then the dot before it, then the digit before the dot
                    [regex]$rx="\d\.\d$"

                    $objparam = @{ wsman_supported = $false
                                   wsman_version   = "None"
                                 }

                }#Test-corpIsWsman function Begin Section


                Process {
        
                    Try {

                        Write-Verbose -Message "testing WSMAN on $Computername using Test-corpIsWsman Function"

                        $result = Test-WSMan -ComputerName $Computername -ErrorAction Stop

                        Write-Verbose -Message "WSMAN is accessible on $Computername"

                        $objparam["wsman_supported"] = $true 

                    }#try


                    Catch {

                        #Write the error to the pipeline if the computer is offline
                        #or there is some other issue

                        Write-Verbose -Message  "Cannot connect to $Computername using WSMAN."

                        Write-Verbose -Message "$ComputerName cannot be accessed using WSMAN ..."

                        New-Object -TypeName PSObject `
                                   -Property $objparam  

                        write-Error $_.exception.message
            
                    }#catch

        
        
                    if ($result) {
            
                        Write-Verbose -Message "Checking WSMAN version on $Computername"

                        $m = $rx.match($result.productversion).value

                             if ($m -eq '3.0') {
                
                                Write-Verbose -Message "$ComputerName running WSMAN version 3.0 ..."
                    
                                $objparam["wsman_version"] = "wsman v3"
                   
                                New-Object -TypeName PSObject `
                                           -Property $objparam
                             }

                             else {
                
                                Write-Verbose -Message "$ComputerName running WSMAN version 2.0 ..."
               
                                $objparam["wsman_version"] = "wsman v2"
                
                                New-Object -TypeName PSObject `
                                       -Property $objparam                
                            }        
        
                    }#end if($result)



                }#Test-corpIsWsman function "Process Section"


                End {

                    Write-Verbose -Message "End Test-corpIsWsman Function"

                }#Test-corpIsWsman function "End Section"



        }  # function Test-corpIsWsman


        function Get-CorpCimInfo {

            <#
            .Synopsis
            Creates on-the-fly sessions to retrieve WMI data


            .Description
            The Get-CimInstance cmdlet in PowerShell 3.0 can be used to retrieve WMI information from a remote computer using the WSMAN protocol instead of the legacy WMI service that uses DCOM and RPC.
            However, the remote computers must be running PowerShell 3.0 and the latest version of the WSMAN protocol.

            When querying a remote computer,the script will test if the remote computer can be reached using WS-MAN, and will test if it is running PowerShell 3.0 and the latest WS-MAN.
            If this is the case, Get-CIMInstance setups a temporary CIMSession.

            However, if the remote computer is running PowerShell 2.0 and accessible via WSMAN, then PowerShell remoting is used to get data using Get-WMIObject over PowerShell invoke remoting command.
            If the remote computer cannot be contacted using WS-MAN, DCOM is used (CIMSession with a CIMSessionOption to use the DCOM protocol)

            A switch parameter is added to the script called (-LegacyDCOM), which will instruct the script to use (Get-WMIObject) instead of CIMSession with DCOM as a CIMSessionOption. If you add this switch parameter, and the script failback to using DCOM, then a normal Get-WMIObject will be used to get data. The benifit for this is the type of the returned object which is (ManagementObject) in case of Get-Wmiobject. This type of object can be pipelined to Invoke-wmiMethod if needed.This is the only reason to use this switch

            The script has parameter switches to turn off any connectivity method for more customization.Using -DisableWsman3 switch, will disable connecting to remote machines using WSMAN 3.0.etc.

            Managed exceptions will be thrown wisely to allow calling function to catch them and act accordingly. Exceptions will not break the pipeline if you catch exceptions.

            Running the script with Verbose mode will give a lot of information for you to look at.

            This script is essentially a wrapper around Get-CimInstance and PowerShell remoting to make it easier to query data from a mix of computers that have different levels of WS-MAN support.

            Before returning the results back as an object, we will add a property called (CorpComputer) to identify from where the data get generated.

            If (IncludeConnectivityInfo) switch is used, the returned object will also contains a property called (CorpConnectivityMethod) to show which method is used to collect the data

            Both properties can be viewed by pipelining the returned object with (Select *)



            .PARAMETER Class
            Aliase for this parameter is (ClassName). This is the CIM or WMI class to query. You cannot use this paramater with (Query) parameter at same time. (Class) and (Query) are defined in different parameter sets.

            .PARAMETER Computername
            Aliases for this parameter are (Computer,Host,Hostname,Machine,"MachineName"). Default value for this parameter is localhost.

            .PARAMETER Filter
            String to filter CIM or WMI classes.

            .PARAMETER Property
            Reduce the size of returned data by returning only subset of the properties.

            .PARAMETER NameSpace
            CIM/WMI name space to be used. Default is root\cimv2.

            .PARAMETER Query
            Native wmi query to be executed in the remote computer.You can not use this parameter with (Class) parameter.

            .PARAMETER KeyOnly
            Switch parameter to return key values only.Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER $OperationTimeoutSec
            Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER Shallow
            Only applicable during CIM sessions, not PowerShell remoting.

            .PARAMETER Disablewsman3
            Switch to disable connecting using WS-MAN 3.0 using native Get-CIMInstance over WSMAN.

            .PARAMETER Disablewsman2
            Switch to disable connecting using WS-MAN remoting (PowerShell Remoting).

            .PARAMETER DisableDCOM
            Switch to disable connecting using DCOM (RPC).

            .PARAMETER ShowRunTime
            Switch to show script execution time when running in verbose mode.

            .PARAMETER NoProgressBar
            Switch to hide Progress Bar.By default, a progress bar will be shown to indicate the script progress.

            .PARAMETER includeConnectivityInfo.
            #This switch will add a property to the returned object that indicates the type of method used when retrieving information from each remote computer.

            .PARAMETER LegacyDCOM
            A switch parameter that will instruct the script to use (Get-WMIObject) instead of CIMSession with DCOM as a CIMSessionOption. If you add this switch parameter, and the script failback to using DCOM, then a normal Get-WMIObject will be used to get data. The benifit for this is the type of the returned object which is (ManagementObject) in case of Get-Wmiobject. This type of object can be pipelined to Invoke-wmiMethod if needed.This is the only reason to use this switch.




            .Example
            Get computer names from pipeline.
            PS C:\> get-content computers.txt | Get-CorpCimInfo -class win32_logicaldisk -filter "drivetype=3"

            .Example
            Disable the use of DCOM to access information
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableDCOM

            .Example
            Only enable DCOM to access information using Get-CIMSession with DCOM as a CIMSessionOption
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableWsman3 -DisableWsman2

            .Example
            Only enable DCOM to access information and force the script to use the legacy (Get-WMIObject) method
            PS C:\> Get-CorpCimInfo -Class Win32_Bios "Localhost","Host1" -DisableWsman3 -DisableWsman2 -LegacyDCOM

            .Example
            Use Query parameter insted of ClassName
            PS C:\> Get-CorpCimInfo -Query "SELECT * from Win32_Process WHERE name LIKE 'p%'" -Computername "Host1" 

            .Example
            Using Start-Job to pass array of servers to get-corpCimInfo
            [String[]] $list = @("localhost","localhost")

             $job = Start-Job -ScriptBlock { 
                        param ( [String[]] $list )
                        $list | % {Get-CorpCimInfo -Class win32_bios -computerName $_ }
                     } -ArgumentList (,$list)
              Wait-Job -Job $job | Out-Null
              Receive-Job $job


            Versioning:
                - Version 1.0 written  5 November 2013 : flowControl returns string for $protocolToUse
                - Version 2.0 written 13 November 2013 : - FlowControl returns an object for $protocolTouse
                                                         - Better exception handling when getting return values from Test-corpIsWsman
                                             




            .Notes
            Last Updated: Nov 11, 2013
            Version     : 2.0
            Author      : Ammar Hasayen (@ammarhasayen)
            Based on    : Jeffery Hicks script (@JeffHicks)

            .Link
            http://ammarhasayen.com

            .Link
            Get-CimInstance
            New-CimSession
            New-CimsessionOption

            .Inputs
            string

            .Outputs
            CIMInstance

            #>

            [cmdletbinding()]

                Param(

                [Parameter(Mandatory=$true, `
                           HelpMessage="Enter a class name", `
                           ValueFromPipelineByPropertyName=$true, `
                           ParameterSetName="classOption")]
                [alias("ClassName")]
                [ValidateNotNullorEmpty()]
                [string]$Class,

                 #you cannot provide class parameter and query parameter at the same time
                 [Parameter( Mandatory=$true, `
                             HelpMessage="Enter a Wmi Query", `
                             ValueFromPipelineByPropertyName=$true, `
                             ParameterSetName="queryOtpion")]
                 [string]$Query,

                [Parameter(Position=1, `
                           ValueFromPipelineByPropertyName=$true, `
                           ValueFromPipeline=$true, `
                           HelpMessage="Enter one or more computer names separated by commas.") ]
                [ValidateNotNullorEmpty()]
                [alias("Computer","host","hostname","machine","machinename")]
                [string[]]$Computername=$env:computername,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [string]$Filter,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [string[]]$Property,

                [Parameter(ValueFromPipelineByPropertyName=$true)]
                [ValidateNotNullorEmpty()]
                [string]$Namespace="root\cimv2",
   

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [switch]$KeyOnly,

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [uint32]$OperationTimeoutSec,

                #only available via cim session, when PowerShell remoting is used, this parameter
                #will be ignored
                [switch]$Shallow,

                #use this switch to disable Get-CIMSession over WSMAN 3.0 Connectivity
                [switch]$DisableWsman3,

                #use this switch to disable WSMAN remoting (PowerShell Remoting)
                [switch]$DisableWsman2,

                #use this switch to disable DCOM Connectivity 
                [switch]$DisableDcom,

                #use this switch to enabel DCOM legacy Get-WMIObject when failing back to DCOM
                [switch]$LegacyDCOM,

                 #use this switch to show script execution time when running in verbose mode
                [switch]$ShowRunTime,

                [switch]$NoProgressBar,
                #Switch to hide Progress Bar

                [switch]$includeConnectivityInfo
                #This switch will add a property to the returned object that indicates the type of method used when retrieving information from each remote computer

               #[Parameter(Mandatory=$false)]               
               #[System.Management.Automation.Credential()]$Credential
               #uncomment this parameter if you need to supply a credential

                )#end function parameters


                Begin{
            
                        if ($PSBoundParameters.ContainsKey("ShowRunTime")){
                            #Start stop watch
                            $Watch  =  [System.Diagnostics.Stopwatch]::StartNew()
                        }


                        if (!(($PSBoundParameters.ContainsKey("query") ) -OR ($PSBoundParameters.ContainsKey("class"))) ) {
                            Throw " You should provide either Class parameter or Query Parameter"
                        }

                        if (($PSBoundParameters.ContainsKey("query") ) -AND ($PSBoundParameters.ContainsKey("class")) ) {
                            Throw " You cannot provide both Class parameter and Query Parameter"
                        }

            
                        Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  

                        Write-verbose -Message ($PSBoundParameters | out-string)
    
                        #defining the methods to use when connecting to the remote computer
                        #by default all methods will be performed
                        #by default, wsman3 will be tried first, if it fails, then wsman2, then dcom

                        $propertiesMethods = @{wsman3 = $true;
                                               wsman2 = $true;
                                               dcom   = $true
                                              }
                        $Methods = New-Object -TypeName PSObject -Property $propertiesMethods
                                         
                        if ($PSBoundParameters.ContainsKey("disableWsman3"))  {$Methods.wsman3 = $false}
                        if ($PSBoundParameters.ContainsKey("disableWsman2"))  {$Methods.wsman2 = $false}
                        if ($PSBoundParameters.ContainsKey("disabledcom"))    {$Methods.dcom   = $false}

                        #check if wsman3,wsman2 and dcom are all disabled via parameter switches
                        if (!($Methods.dcom  -OR $Methods.wsman2 -OR $Methods.wsman3 )){
                           Throw  "You cannot disable all test types"
                        }


                        #printing out available tests
                        Write-verbose -Message "Tests available for the script : $(if($Methods.wsman3){"WS-MAN3"} if($Methods.wsman2){"WS-MAN2"} if($Methods.DCOM){"DCOM"})" 
             
          
                        function Get-corpProgress {
	            
		                    param($PercentComplete,$status)
		    
		                    Write-Progress -activity "Get-corpCimInfo Script" `
                                          -percentComplete ($PercentComplete) `
                                          -Status $status

	                    }#end Get-corpProgress function

            
                    
            
                        function flowControl {
                        
                        #route the script execution depending on the connectivity methods available 
                       
                            Param (
                                [Parameter(Mandatory=$true)]
                                [string]$Computername,

                                [Parameter(Mandatory=$true)]
                                [object]$wsSupport,
                    
                                [Parameter(Mandatory=$true)]
                                [object]$Methods,

                                [Parameter(Mandatory=$true)]
                                [object]$ProtocolToUse
                             )

                             Begin {   #function flowControl

                             Write-Verbose -Message "+++++++Entering flowControl function"

                             }   #begin function flowControl

                             Process {#function flowControl
                    
                                #Possibility 1:

                                if ($ProtocolToUse.Method -like "Nothing") {
                    
                                    if ($Methods.wsman3 -and ($wsSupport.wsman_version -like "wsman v3")) {
                                        $ProtocolToUse.Method = "WSMAN 3.0"
                                         Return $ProtocolToUse
                            
                                    }

                                    elseif ($Methods.wsman2 -and ($wsSupport.wsman_supported)) {
                                        $ProtocolToUse.Method = "WSMAN 2.0"
                                         Return $ProtocolToUse
                                    } 

                                    elseif ($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Return $ProtocolToUse
                                    }

                                    else {
                                        Write-Verbose -Message "None of the available connectivity methods works with $Computername"  
                                        Write-Error "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }

                                }#end ($ProtocolToUse.Method -like "Nothing")  
                    
                    
                                #Possibility 2:
                                if ($ProtocolToUse.Method -like "WSMAN 3.0") { #means WSMAN 3.0 conenctivity failed
                        
                                    if ($Methods.wsman2 -AND ($wsSupport.wsman_supported) )  {
                                        $ProtocolToUse.Method = "WSMAN 2.0"
                                        Write-Verbose -Message "failing back to $($ProtocolToUse.Method) for $Computername"
                                        Return $ProtocolToUse
                                    } 

                                    elseif($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Write-Verbose -Message "failing back to $($ProtocolToUse.Method) for $Computername"
                                        Return $ProtocolToUse
                                    } 
                          
                                    else {
                                        Write-Verbose -Message "no failback method for $Computername.. +++++++Exiting"
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }



                              }#end if ($ProtocolToUse.Method -like "WSMAN 3.0")



                                #Possibility 3:
                               if ($ProtocolToUse.Method -like "WSMAN 2.0") { #means WSMAN 2.0 conenctivity failed
                        
                                    if ($Methods.dcom) {
                                        $ProtocolToUse.Method = "dcom"
                                        Return $ProtocolToUse
                                    }

                                     else {
                                        Write-Verbose -Message "no failback method for $Computername.. +++++++Exiting"
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $Computername using any configured method"
                                    }

                               }#end if ($ProtocolToUse.Method -like "WSMAN 2.0")


                     


                             }#process function flowControl

                             End { #function flowControl

                                Write-Verbose -Message "Protocol to be used $($ProtocolToUse.Method)"

                                Write-Verbose -Message "+++++++Exiting flowControl function"

                             }#end block of function flowControl 
                

                        }#end Function flowControl

                   }#end Get-CorpCisInfo function begin block
    



                Process {  #process block for Get-corpCisInfo function

                        Write-Verbose -Message "Processing $($computername.count) computer(s)"

                        [int]$counter = 0 #used to show progress only



                        foreach ($computer in $computername) {       
     
                            if (!($PSBoundParameters.ContainsKey("NoProgressBar"))) {

                                $counter += 1

                                Get-corpProgress -PercentComplete (($counter/($Computername.count))*100) `
                                                 -Status "getting information from $computer"
                             }#end if

                
                            #This is a variable to hold the decision on what method to be used to connect to the remote computer. The flowControl function is the only one the can change this object
                            #All routing decisions are taken based on the value of this object
                
                            $protocolToUse = New-Object -TypeName psobject `
                                                        -Property @{Method = "Nothing"}
                 
                            Write-Verbose -Message "Processing $computer"

                
                            #First thing to do is to hand control to the flowControl funtion
                            #to do that, we will evaluate WSMAN support first using a variable $wsSupport
                            #WSMAN support means testing if the remote computer is accessible via WSMAN
                                
                            #clearing error to get only error data from Test-corpIsWsman function
                            $error.clear()
                
                            if($Methods.wsman3 -OR $Methods.wsman2) {

                               $wsSupport= Test-corpIsWsman `
                                                -ComputerName $computer `
                                                -ErrorAction SilentlyContinue
                  
                            }#end if

                            else {

                                 $objparam = @{ wsman_supported = $false
                                                wsman_version   = "None"
                                              }   
                                 $wsSupport = New-Object -TypeName PSObject `
                                   -Property $objparam 
                            }



                            #building flowControl function parameters:
                            $paramflowControl = @{ ComputerName  = $computer;
                                                   Methods       = $Methods;
                                                   ProtocolToUse = $protocolToUse;
                                                   wsSupport     = $wsSupport
                                                 }
                                 
                            #we will lt the flowControl function determin what method to use to connect
                            $protocolToUse = flowControl @paramflowControl

                            #Since exception is not thrown, then there is a method to be evaluated
            
                            #hashtable of parameters for New-CimSession or remoting
                            #adding computername and EA

                            $sessParam = @{Computername=$computer;ErrorAction='Stop'}

                            #credentials?
                            if (($PSBoundParameters.ContainsKey("Credential")))
                                {

                                if ($credential) {
                                    Write-Verbose -Message "Adding alternate credential for CIMSession"
                                    $sessParam.Add("Credential",$Credential)
                                }#end if if ($credential)

                            }#end if


                             #WSMAN 3.0 method
                             if ($protocolToUse.Method -like "WSMAN 3.0") {
                                 
                                 Write-Verbose -Message "trying $($protocolToUse.Method) on $computer"

                                 Try {               
                                     $session = $null
                                     $session = New-CimSession @sessParam
                                     Write-Verbose -Message "Session using $($protocolToUse.Method) is created on $computer"
                                 }#end try

                                 Catch {
                                     Write-Warning "Failed to create a CIM session to $computer using $($protocolToUse.Method)"
                                     Write-Warning $_.Exception.Message
                                     #setting alternative method

                                     #building flowControl function parameters:
                                     $paramflowControl = @{ ComputerName  = $computer;
                                                            Methods       = $Methods;
                                                            ProtocolToUse = $protocolToUse;
                                                            wsSupport     = $wsSupport
                                                           }
                                 
                                     #we will lt the flowControl function determin what method to use to connect
                                     $protocolToUse = flowControl @paramflowControl                           
                                                              
                                 }#end catch

                                 if ($session){

                                     #create the parameters to pass to Get-CIMInstance
                                     $paramHash=@{
                                                  CimSession= $session
                                                  }

                                     $cimParams = "Filter","KeyOnly","Shallow","OperationTimeOutSec","Namespace","Query","Class"

                                     foreach ($param in $cimParams) {

                                        if ($PSBoundParameters.ContainsKey($param)) {

                                             Write-Verbose -Message "Adding $param for CIM command : $computer"
                                             $paramhash.Add($param,$PSBoundParameters.Item($param))
                                         } #if

                                     }#foreach param 
                    
                     

                                     #execute the query
                                     Write-Verbose -Message "Querying $class using $($protocolToUse.Method) on $computer"
                    
                                     Try {
                                        $obj = Get-CimInstance @paramhash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)
                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer
                            
                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                     -MemberType NoteProperty `
                                                     -Name "corpConnectivityMethod" `
                                                     -Value $protocolToUse.Method
                                        }
                                        $obj
                                     }#end try

                                     Catch {
                        
                                         Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)" 

                                         #writing out exception
                                         Write-Warning $_.Exception.Message

                                         #setting alternative method
                                         #building flowControl function parameters:
                                         $paramflowControl = @{ ComputerName  = $computer;
                                                                Methods       = $Methods;
                                                                ProtocolToUse = $protocolToUse;
                                                                wsSupport     = $wsSupport
                                                               }
                                 
                                         #we will lt the flowControl function determin what method to use to connect
                                         $protocolToUse = flowControl @paramflowControl      

                                      }#end catch

                                      Finally {
                                         Write-Verbose "Removing CIM 3.0 Session from $computer"
                                         if ($session) {Remove-CimSession $session}
                                      }#end finally

                                 }#end if($session)
                
                             }#end if($protocolToUse.Method -like "WSMAN 3.0")      

                  
                             #WSMAN 2.0 method
                             if ($($protocolToUse.Method) -like "WSMAN 2.0") {
                
                                 $paramHash=@{}
                                 $wsParams = "Filter","Namespace","Query","Property","class"
                
                                 foreach ($param in $wsParams) {
                    
                                    if ($PSBoundParameters.ContainsKey($param)) {
                        
                                        Write-Verbose -Message "Adding $param for PS remoting command : $computer"
                                        $paramhash.Add($param,$PSBoundParameters.Item($param))

                                    } #end if ($PSBoundParameters.ContainsKey($param))

                                }#foreach ($param in $wsParams       

                                #execute the query
                                Write-Verbose -Message "Querying $class using $($protocolToUse.Method) on $computer"
                
               
                                Try {  

                                    $wssession = $null
                         
                                    $wssession = New-PSSession @sessParam
                                  
                                    $obj = Invoke-Command -Session $wssession `
                                                          -ScriptBlock{param($x) Get-WmiObject @x} `
                                                          -ArgumentList $paramhash `
                                                          -EA Stop

                                    #we add a property called "corpComputer" that will help identify the machine returning this info
                                    #this become useful when doing PowerShell Jobs (start-Job)
                                    $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                    if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                     -MemberType NoteProperty `
                                                     -Name "corpConnectivityMethod" `
                                                     -Value $protocolToUse.Method
                                        }

                                    $obj
                                    Write-Verbose -Message "Done : querying $class using $($protocolToUse.Method) on $computer"

                                }#end try

                
                                Catch {

                                    Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)"

                                    #writing out exception
                                    Write-Warning $_.Exception.Message
                                        
                                    #setting alternative method

                                    #building flowControl function parameters:
                                    $paramflowControl = @{ ComputerName  = $computer;
                                                           Methods       = $Methods;
                                                           ProtocolToUse = $protocolToUse;
                                                           wsSupport     = $wsSupport
                                                          }
                                 
                                     #we will lt the flowControl function determin what method to use to connect
                                     $protocolToUse = flowControl @paramflowControl  

                                }#end catch  

                                Finally {

                                    Write-Verbose -Message "Removing PSSession from $computer"
                                    if($wssession) {Remove-PSSession $wssession}

                                }#end Finally   
                          
                            }#end if ($protocolToUse.Method -like "WSMAN 2.0")



                            #dcom method
                            if($protocolToUse.Method -like "dcom")  {
                
                
                                 write-verbose "trying $($protocolToUse.Method) on $computer"

                                 if ($PSBoundParameters.ContainsKey("LegacyDCOM")) {
                                 #using legacy Get-WMIObject command

                                    write-verbose "using Legacy DCOM"

                                    $paramHash=@{
                                                 ComputerName = $computer
                                                }
                        
                                    $dcomLegacyParams = "Filter","Namespace","Query","Property","class"

                                    foreach ($param in $dcomLegacyParams) {
                       
                                        if ($PSBoundParameters.ContainsKey($param)) {
                       
                                             Write-Verbose -Message "Adding $param for Legacy dcom command : $computer"
                        
                                            $paramhash.Add($param,$PSBoundParameters.Item($param))

                                        }#if

                                    } #foreach param 

                                     #execute the query
                                    Write-Verbose "Querying $class using legacy $($protocolToUse.Method) on $computer"
                                    Try {

                                        $obj = Get-WmiObject @paramHash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)

                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                            -MemberType NoteProperty `
                                                            -Name "corpConnectivityMethod" `
                                                            -Value $protocolToUse.Method
                                         }#end if

                                        $obj #return object
                       
                                    }#end Try

                                     Catch {

                                        Write-Verbose -Message " Error while Querying $class on $computer via legacy $($protocolToUse.Method)" 

                                        #writing out exception
                                        Write-Warning $_.Exception.Message

                                        #Write-Error  exception and existing
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"
                                                  
                        
                                     }#end Catch

                                     Finally {
                                     }#end Finally 

                                 }#end if ($PSBoundParameters.ContainsKey("LegacyDCOM"))

                                 else {
                                 #use Get-CIMInstance with CIMSessionOption (Protocol Dcom)
                                 $opt = New-CimSessionOption -Protocol Dcom
                                 $sessparam.Add("SessionOption",$opt)  


                                 Try {               
                                    $session = $null
                                    $session = New-CimSession @sessParam
                                    write-verbose -Message "Session using $($protocolToUse.Method) is created on $computer"
                                  }#end try

                                 Catch {
                                    Write-Warning "Failed to create a CIM session to $computer using $($protocolToUse.Method)"
                                    Write-Warning "No other actions will be performed"
                                    Write-Warning $_.Exception.Message  
                                    Write-Error "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"  
                                 }#end catch

                                 if($session){


                                    #create the parameters to pass to Get-CIMInstance
                                    $paramHash=@{
                                         CimSession= $session
                                                }


                                    $cimParams = "Filter","KeyOnly","Shallow","OperationTimeOutSec","Namespace","Query","Property","class"

                                    foreach ($param in $cimParams) {
                       
                                        if ($PSBoundParameters.ContainsKey($param)) {
                       
                                             Write-Verbose -Message "Adding $param for CIM command : $computer"
                        
                                            $paramhash.Add($param,$PSBoundParameters.Item($param))

                                        }#if
                                    } #foreach param 
                    
                     

                                    #execute the query
                                    Write-Verbose "Querying $class using $($protocolToUse.Method) on $computer"
                                    Try {

                                        $obj = Get-CimInstance @paramhash -EA Stop
                                        #we add a property called "corpComputer" that will help identify the machine returning this info
                                        #this become useful when doing PowerShell Jobs (start-Job)

                                        $obj | Add-Member -MemberType NoteProperty -Name "corpComputer" -Value $computer

                                        if($PSBoundParameters.ContainsKey("includeConnectivityInfo"))
                                            { $obj | Add-Member `
                                                            -MemberType NoteProperty `
                                                            -Name "corpConnectivityMethod" `
                                                            -Value $protocolToUse.Method
                                         }#end if

                                        $obj #return object
                       
                                    }#end Try

                                     Catch {

                                        Write-Verbose -Message " Error while Querying $class on $computer via $($protocolToUse.Method)" 

                                        #writing out exception
                                        Write-Warning $_.Exception.Message

                                        #Write-Error  exception and existing
                                        Write-Error  "+++++++Exiting script with managed exception for you to catch... cannot connect to $computer using any configured method"
                                                  
                        
                                     }#end Catch

                                     Finally {

                                        Write-Verbose -Message "Removing CIM Session from $computer"
                                        if($session){Remove-CimSession $session}

                                     }#end Finally 

                                }#end if($session)
               
                             }#end else



                             }#end if($ProtocolToUse.Method -like "dcom") for dcom    
  

                             Write-Verbose -Message "Finish processing $computer. Last Method used is $($protocolToUse.Method)"

                       }#end foreach ($computer in $computername)
 
                }#end function process block


                End { 
                     Write-Verbose -Message "Script Get-CorpCimInfo Ends"
         
                     if ($PSBoundParameters.ContainsKey("ShowRunTime")) {
                         #Stop and display stop watch
            
                         Write-Verbose -Message "Script run time (Minutes:seconds:milliseconds): $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString()):$($Watch.Elapsed.MilliSeconds.ToString())"
                     }#end if ($PSBoundParameters.ContainsKey("ShowRunTime"))

                }# end function end block



        } # function Get-CorpCimInfo function


        function _GetExSvrMailboxCount {
    
	        #Quote: Script block derived from Steve Goodman script
            param($Mailboxes,$ExchangeServer,$Databases)
	        # The following *should* work, but it doesn't. Apparently, ServerName is not always returned correctly which may be the cause of
	        # reports of counts being incorrect
	        #([array]($Mailboxes | Where {$_.ServerName -eq $ExchangeServer.Name})).Count
	
	        # ..So as a workaround, I'm going to check what databases are assigned to each server and then get the mailbox counts on a per-
	        # database basis and return the resulting total. As we already have this information resident in memory it should be cheap, just
	        # not as quick.
	        $MailboxCount = 0
	        foreach ($Database in [array]($Databases | Where {$_.Server -eq $ExchangeServer.Name}))
	        {
		        $MailboxCount+=([array]($Mailboxes | Where {$_.Database -eq $Database.Identity})).Count
	        }
	        $MailboxCount
            #End Quote: Script block derived from Steve's script
	
        } # function _GetExSvrMailboxCount


        function Get-CorpNetBiosFromFDQN{    
            <#    
            .SYNOPSIS
                Give me a string and i will trim everything after the first dot.
                This will simply convert FDQN to Netbios name.       

                Version                    :         1.0
                Author                     :         Ammar Hasayen (@ammarhasayen)(http://ammarhasayen.com)  

            .PARAMETER $name
             String name

             .EXAMPLE     
             Output should be "servername"
            .\Get-CorpNetBiosFromFDQN "servername.contoso.com"

            #>  

            [cmdletbinding()]

            param(
                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage='String name')]
                [ValidateNotNullorEmpty()]
                [string]$name
            )

            [string]$Output = ""

            if (!($name.Contains("."))) {
        
                #input string does not have dots
                Write-Output $name
            }
            else {
                    #input string has at least one dot

                    if ($name.IndexOf(".")){

                        #we can get an index for the first dot in the input strings

                        $myindex = $name.IndexOf(".")

                        $output = $name.Substring( 0,$myindex)

                        Write-Output $Output

                    }

                    else {Write-Output $name }
            }


     } # function Get-CorpNetBiosFromFDQN


        function Get-CorpDagMemberList{    
            <#    
            .SYNOPSIS
                Give me DAG name and will return string array of members

                OR will return $null if error happens

                Input should be a string that represent the name of the DAG

                Output is either array string or null

                Version                    :         1.0
                Author                     :         Ammar Hasayen (@ammarhasayen)(http://ammarhasayen.com)        


            .PARAMETER $DatabaseAvailabilityGroup
             String name of the database availability group.

             .EXAMPLE     
            .\Get-CorpDagMemberList -$DatabaseAvailabilityGroup "NYC"

            #>  

            [cmdletbinding()]

            param(

                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage='String name of the database availability group')]
                [ValidateNotNullorEmpty()]
                [string]$DatabaseAvailabilityGroup
            )

            #region defining variables

                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Entering Get-CorpDagMemberList function with DAG name : $($DatabaseAvailabilityGroup.ToUpper())"

                $DatabaseAvailabilityGroup_Members = @()

                $DatabaseAvailabilityGroup_Members_List = @()            

            #endregion #region defining variables

            try {
                $DatabaseAvailabilityGroup_Members = (Get-DatabaseAvailabilityGroup $DatabaseAvailabilityGroup -ErrorAction STOP).servers 

                $DatabaseAvailabilityGroup_Members_List = $DatabaseAvailabilityGroup_Members |foreach{$_.name}

                #region Log Info 
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpDagMemberList] Info : DAG $DatabaseAvailabilityGroup : number of members returned : $($DatabaseAvailabilityGroup_Members_List.count)"    
                foreach ($var in $DatabaseAvailabilityGroup_Members_List) {
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpDagMemberList] Info : DAG Member Name : $var "    


                }
                #endregion Log Info 



            }catch {
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpDagMemberList] Error DAG $DatabaseAvailabilityGroup : fails Get-DatabaseAvailabilityGroup command. Skipping $($DatabaseAvailabilityGroup.ToUpper())"
                $DatabaseAvailabilityGroup_Members_List = $null
            }    
        
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Exiting Get-CorpDagMemberList function with DAG name : $DatabaseAvailabilityGroup"

            Write-Output $DatabaseAvailabilityGroup_Members_List

     } # function Get-CorpDagMemberList


        function _GetExchServiceHealth {

            param ($Server,$UsePSRemote)

            $ServiceDownCount = 0
            $ServicesHealth   = "Fail"
            $wmi = $null
         
            #Check connectivity 

             if(!$UsePSRemote) {
     
                try{
                    $wmi = get-WmiObject win32_service -ComputerName $Server -ErrorAction Stop 
                }catch {
                    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):[_GetExchServiceHealth] Error : Cannot get Service health from $Server using get-WmiObject win32_service ... Skipping"
                }        
     
             }else {

                try{
                    $wmi = Get-CorpCimInfo -class win32_service -ComputerName $Server -ErrorAction Stop 
                }catch{
                 Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):[_GetExchServiceHealth] Error : Cannot get Service health from $Server using Get-CorpCimInfo -class win32_service ... Skipping"
                 }
     
             } # if(!$UsePSRemote)
       
  
          if (!$wmi) {
    
            $ServicesHealth = "N/A"
            $ServiceDownCount = "N/A"
    
          } # we cannot connect to the server
          else {# we can connect to the server and retrieve services

               #Get Exchange services that are automatic and stopped

               $ServiceStatus = $wmi |
                   where {($_.displayName -match "exch*") -and ($_.StartMode -match "Auto") -and($_.state -match "Stopped")}
                   
               # If any stopped services
               if ($serviceStatus -ne $null) {
                   foreach ($service in $serviceStatus) {
	                             
	                    $SvcName = $service.name
	                    $SvcState = $service.state
	                    $ServiceDownCount += 1
                        $ServicesHealth = "Fail"
                   }
               }
               else { # If all services are started
                      
                    $SvcName = "N/A"
	                $SvcState = "N/A"
                    $ServicesHealth = "Pass"
                    $ServiceDownCount = "None"
               }
          }

                @{ServicesHealth  = $ServicesHealth
                  ServiceDownCount = $ServiceDownCount 
                 }
       } # function _GetExchServiceHealth


        function Test-DAGFullMembers {
    
            param($DAG, $ExchangeServersList)
            # give me DAG Object and string array of exchange servers and i will verify that the exchange server string array contains all members of that DAG
        
            $DAGSrv = $DAG.members
            $success = $true

            foreach ($srv in $DAGSrv) {

                if ($ExchangeServersList -notcontains $srv) {
                    $success  = $false

                }
            }

            Write-Output $success
    
       }  # Test-DAGFullMembers


        function Get-CorpUptime {           
    

            [cmdletbinding()]

                Param(
                [Parameter(Position = 0,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]
                [alias("Name","MachineName","Computername","Host","Hostname")]
                [string]$Computer,

                [Parameter(Position = 1,
                           ValueFromPipeline = $true,
                           ValueFromPipelineByPropertyName = $true)]
    
                [bool]$UsePSRemote

                )


                Begin {   
     
                # Begin block for Get-CorpUpTime 

                    $now = [DateTime]::Now  

   
                } # Get-CorpUpTime Begin block



                Process {
                # Process block for Get-CorpUpTime

        
            
                        #Write-Verbose -Message "Computer : $computer - Start Processing"

                        $paramOutput = @{computername = $computer}

                        try {

                            # getting data

                            if ($UsePSRemote) {

                                $operatingSystem = Get-CorpCimInfo `
                                                    -Class Win32_OperatingSystem `
                                                    -Computername $computer  `
									                -Property lastbootuptime `
									                -EA Stop
                            }else {

                                $operatingSystem = Get-WmiObject `
                                                    -Class Win32_OperatingSystem `
                                                    -Computername $computer  `
									                -Property lastbootuptime `
									                -EA Stop

                            }

                            # when the connectivity method is not DCOM using (Get-WmiObject), then the returned data is already in DatTime format, so no need to convert.

                            # when the connectivity method is DCOM using Get-WmiObject), then the returned value is string and needs converting.

                            try {
                                # just in case you are using Get-WmiObject not Get-CorpCimInfo
                                $boottime=[Management.ManagementDateTimeConverter]::ToDateTime($operatingSystem.LastBootUpTime) 
                            }
                            catch{
                                $boottime = $operatingSystem.LastBootUpTime
                           }
                          
                            $uptime = New-TimeSpan -Start $boottime -End $now

                
                            # building output object

                            $paramOutput.Add("upTime",$uptime)

                            $objOutput = New-Object -TypeName psobject  `
                                                    -Property $paramOutput
                 

                            Write-Output -InputObject $objOutput

                        } #end try

                        catch {
                            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [Get-CorpUpTime] Error : $computer - fail to get Up Time information, returning N/A"


                            #Write-Verbose -Message "Computer : $computer - fail to get information, returning n/a"

                            # building output object

                            $paramOutput.Add("upTime","n/a")

                            $objOutput = New-Object -TypeName psobject  `
                                                    -Property $paramOutput


                            Write-Output -InputObject $objOutput



                        }# end catch

                        #Write-Verbose -Message "Computer : $computer - END Processing"

      



            } # Get-CorpUpTime Process block



        End{
        #End block for Get-CorpUpTime
            
    
    
        } # Get-CorpUpTime End block



    } # funtion Get-CorpUpTime


        function _sendEmail {
     
           param($from,$to,$subject,$smtphost,$ScriptFilesPath,$InfoFullPath,$ErrorFullPath) 

               
                $varerror = $ErrorActionPreference
                $ErrorActionPreference = "Stop"

                #region files inventory
                    
                    $htmlFileName = Join-Path $ScriptFilesPath "HTMLReport.html"

                    $files = @()

                    foreach ($file in (Get-ChildItem $ScriptFilesPath)) {
                        if (($file.name -notlike "HTMLReport*") -and ($file.name -notlike "*.log") ) {
                            $files += $file.Name
                        }
                    } 

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Directory to look at attachment $ScriptFilesPath"
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Files to attached count = $($files.count)"
                #endregion files inventory


                #region prepare SMTP Client

                    $msg = new-object Net.Mail.MailMessage            
                    $msg.From = $from
                    $msg.To.Add($to)
                    $msg.Subject = $subject
                    $msg.Body = Get-Content $htmlFileName 
                    $msg.isBodyhtml = $true 

                #endregion prepare SMTP Client

                
                #region add attachments

                    #protection threshold
                    if ($files.count) {
                        if($files.count -lt 7) {
                            $varattach = $files.count
                        }else {
                            $varattach = 7
                        }

                        for ($i=0; $i -lt $varattach; $i++) {

                            try {

                                $file = $files[$i]
                                $file = Join-Path $ScriptFilesPath $file
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - File  = $file "
                                $attachment = new-object Net.Mail.Attachment($file)
                                $msg.Attachments.Add($attachment)

                            }catch {
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - Error attaching File : $file "
                                Write-CorpError -myError $_  -Info "[Module Final Tasks - Fail attaching file" -mypath $ErrorFullPath
                            }
                        }
                    } else {
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - No files to attach "
                        
                    }

                #endregion add attachments


                #region send email
                    
                    try{
                        $smtp = new-object Net.Mail.SmtpClient($smtphost)
                        $smtp.Send($msg)
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Email Sent to $to"
                    }catch {
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Fail to send email... check Error log"
                        Write-CorpError -myError $_  -Info "[Module Final Tasks - Could not send email" -mypath $ErrorFullPath
                        _status "      Could not send email... check Info log for detials" 2 
                    }

                #endregion send email


                $ErrorActionPreference = $varerror 
                
               

        }  # function _sendEmail


        function _screenheadings {

            Cls
            write-host 
            write-host 
            write-host 
            write-host "--------------------------" 
            write-host "Script Info" -foreground Green
            write-host "--------------------------"
            write-host
            write-host " Script Name : Email Organization Report (Get-CorpEmailReport)"  -ForegroundColor White
            write-host " Author      : Ammar Hasayen @ammarhasayen" -ForegroundColor White
            write-host " Copy rights : Included in the script code"
            write-host " Version     : 2.4.9"   -ForegroundColor White
            write-host
            write-host "--------------------------" 
            write-host "Script Release Notes" -foreground Green
            write-host "--------------------------"
            write-host
            write-host "-Account Requirements" -ForegroundColor Yellow 
            write-host "   1. Local Administrator on Exchange Servers to collect WMI data  "
            write-host "   2. (ViewAdministrator) Exchange Role"         
            write-host
            write-host "-Script Requirement" -ForegroundColor Yellow
            Write-Host "   * (Microsoft Chart Controls for Microsoft .NET Framework)"
            Write-Host "     to be installed only on the machine that will run the script."
            Write-Host "   * This is needed in order to generate charts"
            write-host "   * Download here : Microsoft Chart Controls for Microsoft .NET Framework 3.5"
            write-host "   * http://www.microsoft.com/en-us/download/details.aspx?id=14422" 
            write-host
            write-host "-Script may take some time to run, depending on your environment size."
            write-host
            write-host "-Run the script with -Verbose switch to get more logging info"
            write-host "-To Get WMI Information from EDGE Servers, check http://wp.me/p1eUZH-pZ"
            write-host
            write-host "-ALWAYS CHECK FOR NEWER VERSION @" -NoNewline
            Write-Host " http://ammarhasayen.com"  -ForegroundColor Red
            write-host  
            write-host "--------------------------" 
            write-host "Script Start" -foreground Green
            write-host "--------------------------"
            Write-Host
            
    } # function _screenheadings


        function _screenFooter {
            param ($ExchangeEnvironment,$Databases , $Watch)

   

            
                write-host 
                write-host 
                write-host 
                write-host "--------------------------" 
                write-host "Script Log Files" -foreground Green
                write-host "--------------------------"
                write-host
                write-host " Three log files are created :"  
                write-host "    - Info Log" -NoNewline -ForegroundColor Yellow
                Write-host "     : records the flow of the script and any non termination errors"
                write-host "    - Detailed log" -NoNewline -ForegroundColor Yellow
                Write-Host " : records detailed info about Exchange Servers and Database Info"
                write-host "    - Error Log" -NoNewline -ForegroundColor Yellow
                Write-Host "    : Terminating error info."  
                write-host
                write-host "--------------------------" 
                write-host "Script Results" -foreground Green
                write-host "--------------------------"
                write-host
                write-host "- Total Mailbox Count" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($ExchangeEnvironment.TotalMailboxes) Mailboxes" 
                write-host "- Total DB Size" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($ExchangeEnvironment.DBSizes) GB"   
                write-host "- Total Archives Size" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($ExchangeEnvironment.TotalArchivesSize) GB"                  
                write-host 
                write-host " -Total Exchange Servers processed" -NoNewline -ForegroundColor Yellow
                Write-Host " : $(($ExchangeEnvironment.Servers).count)"
                write-host " -Total Exchange DAGs processed" -NoNewline -ForegroundColor Yellow
                Write-Host " : $(($ExchangeEnvironment.DAGs).count)"
                write-host " -Total Exchange DBs processed" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($Databases.count)"
                Write-Host
                write-host "--------------------------" 
                write-host "Script Environment" -foreground Green
                write-host "--------------------------"
                Write-Host
                if($PSVersionTable.PSVersion.major){
                write-host " PowerShell Host Version" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($PSVersionTable.PSVersion.major) " 
                } 
                write-host " Script run time ( Minutes : seconds : milliseconds )" -NoNewline -ForegroundColor Yellow
                Write-Host " : $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString()):$($Watch.Elapsed.MilliSeconds.ToString())"  
                write-host
                write-host "--------------------------" 
                write-host "Script Ends" -foreground Green
                write-host "--------------------------" 
            
        } # function __screenFooter


        function _status {

          param($text,$code)       
        
            if ($code -eq 1) {        
                write-host
                write-host "$text"  -foreground cyan
            }
            if ($code -eq 2) {        
            write-host "$text"  -foreground Magenta  
            } 
            if ($code -eq 3) {        
            write-host "$text"  -foreground White 
            }   

     } # function _status 


        function Get-CorpIntersectDags {
    
             # Give me string array of exchange servers and i will return
             # array of DAG Objects that they belong to

             param ($ExchangeServersList)

             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Entering Get-CorpIntersectDags function"
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpIntersectDags] Getting all DAGs using Get-DatabaseAvailabilityGroup and see if each of[ExchangeServersList] variable belongs to it "
 
             $Output_Obj = @()
             $AllDAGs = @()
             $dag_members_list = @()
 
             $AllDAGs =  [array](Get-DatabaseAvailabilityGroup)  
                   
                    foreach ($dag in $AllDAGs) {
                   
                        #this is string array of members
                        $dag_members_list = @(Get-CorpDagMemberList $dag)

                        #if the dag in the loop has members
                        if ($dag_members_list.count) {

                            #looping through each string value representing member server in the dag in question
                            foreach ($dag_member in $dag_members_list) {
                                
                                #if the dag member exist within the input list of servers
                                if ($ExchangeServersList -contains $dag_member) {
                                    
                                    # we will add the DAG Object to our output result
                                    $Output_Obj +=  $dag

                                } # if ($ExchangeServersList -contains $dag_member) 

                            } #foreach ($dag_member in $dag_members_list)
                           
                    
                        } #if ($dag_members_list)
                   
                   
                    }#foreach 

                    #region Log Info 

                    if($Output_Obj.count -eq 0) {

                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpIntersectDags] the input servers does not belong to a DAG "


                    }else {
                        foreach ($var in $Output_Obj) {

                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpIntersectDags] ExchangeServersList belong to this DAG : $var "

                        }
                    }

                    #endregion Log Info 

                           Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Get-CorpIntersectDags] +++++++Exiting Get-CorpIntersectDags function "

                    Write-Output ($Output_Obj |Get-Unique)

      }  # function Get-CorpIntersectDags   


    #endregion helper functions


    #region non helper functions


        function _parameterGetExch {

            # This function will look at each filter in the function and will return 
            # the Exchange Server Objects matching that filter.

            param($ScriptFilter)        
            
                # preparing the output result which is empty array to hold
                # filtered Get-ExchangeServer output 
                $myExchangeServers = @()
    
                Switch ($ScriptFilter) {
    
                        "ServerFilter" {

                                  # Quote : Script block taken from Steve's script

                                  try {
                                    $myExchangeServers = [array](Get-ExchangeServer $ServerFilter -ErrorAction Stop) 
                                  }catch {
                                    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error :(Get-ExchangeServer) with ServerFilter : ""$($ScriptFilter)"""
                                    throw " [_parameterGetExch] Error :(Get-ExchangeServer) with ServerFilter : ""$($ScriptFilter)"""
                              

                                  }
                          
                                  if (!$myExchangeServers) {
        
	                                    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error :No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""
                                        throw "[_parameterGetExch] Error: No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""
                                   }

                                 # End Quote : Script block taken from Steve's script

                                 break;

                        } # "ServerFilter" 



                        "DAGFilter" {

                            $var = @()
                            $var_dag= @()
    
                            foreach ($InputDAG in $InputDAGs) {

                                   #string array of DAG server members
                           
                                   $var = Get-CorpDagMemberList $InputDAG

                                   if ($var) {

                                   $var_dag += $var

                                   }

                                } #end foreach ($InputDAG in $InputDAGs)

                     
                            foreach ($inputServer in $var_dag) {

                                    try {
                                        $var_Srv = Get-ExchangeServer $inputServer -ErrorAction stop

                                        $myExchangeServers += $var_Srv
   
                                    }catch {
                                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error : Server $inputServer fails Get-ExchangeServer Command. Skipping"
                                        Write-Verbose -Message "[_parameterGetExch] Error :Server $inputServer fails Get-ExchangeServer Command. Skipping"

                                    }


                                } #end foreach ($inputServer in $var_dag)


                                if (!$myExchangeServers) {
        
	                                    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error:No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""

                                        throw "[_parameterGetExch] Error : No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""
                                   }

                                break;

                        } # "DAGFilter"

                        "IncludeServerFilter" {

                       
                                foreach ($inputServer in $OnlyIncludedServers) {

                                    try {
                                        $var = Get-ExchangeServer $inputServer -ErrorAction stop
                                        $myExchangeServers += $var
   
                                    }catch {
                                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error : Server $inputServer fails Get-ExchangeServer Command. Skipping"
                                        Write-Verbose -Message "[_parameterGetExch] Error :Server $inputServer fails Get-ExchangeServer Command. Skipping"

                                    }


                                } #end foreach

                                if (!$myExchangeServers) {
        
	                                    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [_parameterGetExch] Error : No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""

                                        throw "[_parameterGetExch] Error : No Exchange Servers matched by -ServerFilter ""$($ScriptFilter)"""
                                   }


                                break;
                        } # "IncludeServerFilter"           


                    } #end switch
         
            #Get-Unique in case of duplicate user input for the same server.
            Write-Output ($myExchangeServers |Get-Unique)


      } # function _parameterGetExch 


        function _parameterGetDBs {

            #taken into consideration the filter in the script filter,
            #we will return database objects (Get-MailboxDatabase) that are
            #hosted on the filtered Exchange Servers

            param ($ScriptFilter, $E2010, $E2013, $ServerFilter, $ExchangeServersList)
        
            $mydatabases = @()

            Switch ($ScriptFilter) {
    
                        #region ServerFilter
                        "ServerFilter" {
                 
                    
                            # Quote : Script block taken from Steve's script
 
                            if ($E2010) {

                                if ($E2013) {
                        	
                                     $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status)  | Where {$_.Server -like $ServerFilter}
                              
                                } # if ($E2013)

                                elseif ($E2010) {
                        	
                                     $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status)  | Where {$_.Server -like $ServerFilter} 

                                } # elseif

                             } # if ($E2010)
                     
                             else { # if (!$E2010)
                        
                                $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where {$_.Server -like $ServerFilter}

                             }#else 

                            # END Quote : Script block taken from Steve's script

                            break;

                        } # "ServerFilter"

                        #endregion ServerFilter

                        #region DAGFilter,IncludedServerFilter
                        Default {                    
                    
 
                            if ($E2010) {

                                if ($E2013) {
                        	
                                     $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status)  | Where {$ExchangeServersList -contains $_.Server }
                              
                                } # if ($E2013)

                                elseif ($E2010) {
                        	
                                     $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status)  | Where {$ExchangeServersList -contains $_.Server } 

                                } # elseif

                             } # if ($E2010)
                     
                             else { # if ($E2010)
                        
                                $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where {$ExchangeServersList -contains $_.Server }

                             }#else 
                    

                            break;


                        } #default

                        #endregion DAGFilter,IncludedServerFilter

            } #switch

            Write-Output $mydatabases

        } # function _parameterGetDBs


        function _parameterGetMailboxes {
    

            param($Databases)

            $myMailboxes = @()

            foreach ($database in $Databases) {

                $myMailboxes  += @(get-mailbox -ResultSize unlimited -Database $database.identity  -WarningAction silentlycontinue)

            } #foreach
    
            Write-Output $myMailboxes

        } #function _parameterGetMailboxes


        function _parameterGetArchiveMailboxes {
    
            param($DatabasesList )

            $myArchiveMailboxes = @()
                    
            $myArchiveMailboxes = [array](Get-Mailbox -Archive -ResultSize Unlimited) | Where {$DatabasesList  -contains $_.ArchiveDatabase }

            Write-Output $myArchiveMailboxes


        } # function _parameterGetArchiveMailboxes


        function _parameterGetDAG {

            param ($ServerFilter,$InputDAGs,$ExchangeServersList)
    
            
            $myDAGs = @()

            Switch ($ScriptFilter) {
    
                        "ServerFilter" {

                            $myDAGs = [array](Get-DatabaseAvailabilityGroup) | Where {$_.Servers -like $ServerFilter}

                            break;

                        } # "ServerFilter"



                        "DAGFilter" {

                            foreach ($inputDAG in $InputDAGs) {

                                try{
                                    $var = Get-DatabaseAvailabilityGroup $inputDAG -ErrorAction STOP
                                    $myDAGs += $var
                                }
                                Catch {
                                    Write-Verbose -Message "[_parameterGetDAG] Error: DAG $InputDAG : fails Get-DatabaseAvailabilityGroup Command. Skiping it"
                                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [_parameterGetDAG] Error :DAG $InputDAG : fails Get-DatabaseAvailabilityGroup Command. Skiping it"
                                }


                            }#foreach

                            break;

                        } # "DAGFilter"



                        "IncludeServerFilter" {

                            $myDAGs = Get-CorpIntersectDags  $ExchangeServersList

                            break;

                        } # "IncludeServerFilter" 



            } #switch

            Write-Output $myDAGs


        } # function _parameterGetDAG
 

             #Quote: Script block derived from Steve Goodman script

        function _TotalsByVersion {

	        param($ExchangeEnvironment)
	    
            #region log info
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): +++++++Entering _TotalsByVersion function"
            #endregion log info 


            $TotalMailboxesByVersion=@{}
	        if ($ExchangeEnvironment.Sites)
	        {
		        foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
		        {
			        foreach ($Server in $Site.Value)
			        {
				        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
				        {
					        $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)",@{ServerCount=1;MailboxCount=$Server.Mailboxes})
				        } else {
					        $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
					        $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount+=$Server.Mailboxes
				        }
			        }
		        }
	        }
	        if ($ExchangeEnvironment.Pre2007)
	        {
		        foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator())
		        {
			        foreach ($Server in $FakeSite.Value)
			        {
				        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"])
				        {
					        $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)",@{ServerCount=1;MailboxCount=$Server.Mailboxes})
				        } else {
					        $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
					        $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount+=$Server.Mailboxes
				        }
			        }
		        }
	        }
	        $TotalMailboxesByVersion
        } #  function _TotalsByVersion

    
        function _TotalsByRole {

	        param($ExchangeEnvironment)
	    
            #region log info
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): +++++++Entering _TotalsByRole function"
            #endregion log info

            # Add Roles We Always Show
	        $TotalServersByRole=@{"ClientAccess" 	 = 0
						          "HubTransport" 	 = 0
						          "UnifiedMessaging" = 0
						          "Mailbox"			 = 0
						          "Edge" 			 = 0
						          }
	        if ($ExchangeEnvironment.Sites)
	        {
		        foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator())
		        {
			        foreach ($Server in $Site.Value)
			        {
				        foreach ($Role in $Server.Roles)
				        {
					        if ($TotalServersByRole[$Role] -eq $null)
					        {
						        $TotalServersByRole.Add($Role,1)
					        } else {
						        $TotalServersByRole[$Role]++
					        }
				        }
			        }
		        }
	        }
	        if ($ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
	        {
		
		        foreach ($Server in $ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
		        {
			
			        foreach ($Role in $Server.Roles)
			        {
				        if ($TotalServersByRole[$Role] -eq $null)
				        {
					        $TotalServersByRole.Add($Role,1)
				        } else {
					        $TotalServersByRole[$Role]++
				        }
			        }
		        }
	        }
	        $TotalServersByRole

        } # function _TotalsByRole

    
        function _GetDAG {

	        param($DAG)

            #region log info
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):-------------------"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):  -DAG Name : $($DAG.Name.ToUpper())"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):  -MemberCount $($DAG.Servers.Count)"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):-------------------"
        

            #endregion log info

	        @{Name			= $DAG.Name.ToUpper()
	          MemberCount	= $DAG.Servers.Count
	          Members		= [array]($DAG.Servers | % { $_.Name })
	          Databases		= @()
	          }
        } # function _GetDAG

             #End Quote: Script block derived from Steve's script


        function _GetExSvr {

	        # This function is based completely on Steve's Goodman Script with couple of modifications here and there

            param($E2010,$ExchangeServer,$Mailboxes,$Databases,$UsePSRemote)
        
        
            #region log info
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): +++++++Entering _GetExSvrs function with Server : $($ExchangeServer.name) and UsePSRemote value : $UsePSRemote"
            #endregion log info 


	        #region Set Basic Variables

	            $MailboxCount = 0
	            $RollupLevel = 0
	            $RollupVersion = ""
                $Roles =""
                $ExtNames = @()
                $IntNames = @()
                $CASArrayName = "" 
                [string]$ExServicesHealth = "N/A"
                [string]$ExServicesFailCount = "N/A"
                $ExUpTime = $null

            #endregion Set Basic Variables


            #region is WMI reachable?
             
                 if (!$UsePSRemote) {

                    try{
                        $tWMI = Get-WmiObject Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction Stop
                        $tWMI_test = $tWMI
                    }catch {
                        Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) for O.S Info (Get-WmiObject Win32_OperatingSystem) "
                    }

                }# if (!$UsePSRemote)

                else { #if $UseRemote

                    try {
                        $tWMI = Get-CorpCimInfo -Class Win32_OperatingSystem $ExchangeServer.Name -ErrorAction Stop
                        $tWMI_test = $tWMI
                    }catch{
                        Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) for O.S Info (Get-CorpCimInfo -Class Win32_OperatingSystem) "
                    }
                }


            #endregion is WMI reachable?    
    
	
	        #region Getting Exchange WMI Info

            if($tWMI_test) {
            
                #region service health info
                    $var =  _GetExchServiceHealth  $ExchangeServer.name  $UsePSRemote    
                    $ExServicesHealth =$var.ServicesHealth 
                    $ExServicesFailCount =$var.ServiceDownCount 
                #endregion service health info

                #region getting server up time
                    $var = Get-CorpUptime $ExchangeServer.name  $UsePSRemote
                    $ExUpTime = $var.uptime.days
                #endregion getting server up time

                #region getting server O.S info
                    $OSVersion = $tWMI.Caption.Replace("(R)","").Replace("Microsoft ","").Replace("Enterprise","Ent").Replace("Standard","Std").Replace(" Edition","")
		            $OSServicePack = $tWMI.CSDVersion
		            $RealName = $tWMI.CSName.ToUpper()
                #endregion getting server O.S info

                #region getting server disk info
                
                    if (!$UsePSRemote) {

                        try {
                            $tWMI=Get-WmiObject -query "Select * from Win32_Volume" -ComputerName $ExchangeServer.Name -ErrorAction Stop
                        }catch {
                            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) for Disk Info (Get-WmiObject -query ""Select * from Win32_Volume"")"                    
                        } # if (!$UsePSRemote)
            
                     }
            
                    else {
                        try {
                            $tWMI = Get-CorpCimInfo -Query "Select * from Win32_Volume" -Computername $ExchangeServer.Name -ErrorAction Stop
                        }catch {
                            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) for Disk Info (Get-CorpCimInfo -Query ""Select * from Win32_Volume"")"
                        }

                     } #else


                      if ($tWMI) {
                         $Disks=$tWMI | Select Name,Capacity,FreeSpace | Sort-Object -Property Name
                      } 
                      else {
                        $Disks =$null
                      }

                
                #endregion getting server disk info


              } # if($tWMI_test)

            else { #$tWMI_test = $false
                    
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Skipping Exchange Service Health info "
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Skipping Server Up Time info "
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Skipping Server Disk Info "

                $OSVersion = "N/A"
		        $OSServicePack = "N/A"
		        $RealName = $ExchangeServer.Name.ToUpper()
                $Disks=$null

            } #else >> $tWMI_test = $false

            #endregion Getting Exchange WMI Info


	        #region Get Exchange Version

	            if ($ExchangeServer.AdminDisplayVersion.Major -eq 6)
	            {
		            $ExchangeMajorVersion = "$($ExchangeServer.AdminDisplayVersion.Major).$($ExchangeServer.AdminDisplayVersion.Minor)"
		            $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace("Service Pack ","")
	            } else {
		            $ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
		            $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
	            }

            #endregion Get Exchange Version	
            

            #region Exchange 2007+
	             if ($ExchangeMajorVersion -ge 8) {
	            
		                #region Exchange 2007+ Get Roles
		                    $MailboxStatistics=$null
	                        [array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(" ","").Split(",");
		                    if ($Roles -contains "Mailbox") {
		            
			                    $MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
			                    if ($ExchangeServer.Name.ToUpper() -ne $RealName) {
			            
				                    $Roles = [array]($Roles | Where {$_ -ne "Mailbox"})
				                    $Roles += "ClusteredMailbox"
			                    }
			                    # Get Mailbox Statistics the normal way, return in a consitent format
			                    $MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer  | Select DisplayName,@{Name="TotalItemSizeB";Expression={$_.TotalItemSize.Value.ToBytes()}},@{Name="TotalDeletedItemSizeB";Expression={$_.TotalDeletedItemSize.Value.ToBytes()}},Database
	                        }
                        #endregion Exchange 2007+ Get Roles
            


                        #region  Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)

                            if ($Roles -contains "ClientAccess" -and $E2010) {                            
            
                                Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
                                Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
                                Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
                                Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
                                $IntNames+=(Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverInternalURI.Host
                                
                                if ($ExchangeMajorVersion -ge 14){
                                
                                    Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | %{ $ExtNames+=$_.ExternalURL.Host; $IntNames+=$_.InternalURL.Host; }
                                }

                                $IntNames = $IntNames|Sort-Object -Unique
                                $ExtNames = $ExtNames|Sort-Object -Unique
                                $CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name

                                if ($CASArray){
                                
                                    $CASArrayName = $CASArray.Fqdn
                                }
                            }

                         #endregion  Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)

	
            
		                #region Rollup Level / Versions (Thanks to Bhargav Shukla http://bit.ly/msxGIJ)

      
                            if ($ExchangeMajorVersion -ge 14){
		        
			                    $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"
		                    } else {
			                    $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"
		                    }

		                    #We need to do remote registry calls, so we will prepare error handling code blocks
            
                                if ($tWMI_test) {
                                    $var = $ErrorActionPreference
                                    $ErrorActionPreference = "Stop"
                                    $RemoteRegistry =  $null

                                    try {

                                            $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name) 
                                            if ($RemoteRegistry) {		                
			                                         $RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach {"$RegKey\\$_"}
			                                         if ($RUKeys) {	
		                            
				                                         [array]($RUKeys | %{$RemoteRegistry.OpenSubKey($_).getvalue("DisplayName")}) | %{
					                                         if ($_ -like "Update Rollup *") {					                            
						                                            $tRU = $_.Split(" ")[2]
						                                            if ($tRU -like "*-*") { $tRUV=$tRU.Split("-")[1]; $tRU=$tRU.Split("-")[0] } else { $tRUV="" }
						                                            if ($tRU -ge $RollupLevel) { $RollupLevel=$tRU; $RollupVersion=$tRUV }
					                                         }
				                                         }

			                                         } #if ($RUKeys)

                                            } #if ($RemoteRegistry) 

                                            else {#else if ($RemoteRegistry)
			                                     Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) using remote registry (Rollup update info)"
                                                 $RollupLevel = 0
                                                 $RollupVersion = $null
		                                    }#else if ($RemoteRegistry)

                           

                                    }#try
		                            catch{ 
                                         Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) using remote registry (Rollup update info)"
                                    } 
                                    finally {
                                        $ErrorActionPreference =$var
                                    }

                                }# if ($tWMI_test)

                                else {#if (twmi_Test)
                                    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) using remote registry (Rollup update info)"
                                    $RollupLevel = 0
                                    $RollupVersion = $null
                                }


                # Exchange 2013 CU or SP Level
                if ($ExchangeMajorVersion -ge 15) {
		    
			        if ($tWMI_test) {
                    $var = $ErrorActionPreference
                    $ErrorActionPreference = "Stop"
                    $RemoteRegistry =  $null
                

                    try {
                        $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
		                $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name);


                        #region if remoteregistry

                             if ($RemoteRegistry)
		                        {
			                        $ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).getvalue("DisplayName")
                                    if ($ExchangeSPLevel -like "*Service Pack*" -or $ExchangeSPLevel -like "*Cumulative Update*")
                                    {
			                            $ExchangeSPLevel = $ExchangeSPLevel.Replace("Microsoft Exchange Server 2013 ","");
                                        $ExchangeSPLevel = $ExchangeSPLevel.Replace("Service Pack ","SP");
                                        $ExchangeSPLevel = $ExchangeSPLevel.Replace("Cumulative Update ","CU"); 
                                    } else {
                                        $ExchangeSPLevel = 0;
                                    }
                                } else {
                                    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot detect CU/SP via Remote Registry for $($ExchangeServer.Name)"
                                    $ExchangeSPLevel = 0
			      
		                        }

                        #endregion if remoteregistry

                    }catch{ 
                         Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot connect to $($ExchangeServer.name) using remote registry (Rollup update info)"
                         $ExchangeSPLevel = 0

                    } finally {
                        $ErrorActionPreference =$var
                    }

		        
                } # if ($tWMI_test) 
                else {
                    $ExchangeSPLevel = 0
                    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Skipping detecting CU/SP via Remote Registry for $($ExchangeServer.Name) because WMI connectivity is not available"
                }

		
	        } # if ($ExchangeMajorVersion -ge 15)            

            } #  if ($ExchangeMajorVersion -ge 8)

            #endregion  Exchange 2007+           






            
            #region Pre Exchange 2007

	            # Exchange 2003
	            if ($ExchangeMajorVersion -eq 6.5)
	            {
		            # Mailbox Count
		            $MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
		            # Get Role via WMI

                    if ($tWMI_test) {
		                try {$tWMI = Get-WMIObject Exchange_Server -Namespace "root\microsoftexchangev2" -Computername $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'" -ErrorAction STOP}
                        catch{$tWMI = $null}
                    } else {$tWMI = $null}

		            if ($tWMI)
		            {
			            if ($tWMI.IsFrontEndServer) { $Roles=@("FE") } else { $Roles=@("BE") }
		            } else {
                        Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot detect Front End/Back End Server information via WMI for $($ExchangeServer.Name)"
			        
			            $Roles+="Unknown"
		            }
		            # Get Mailbox Statistics using WMI, return in a consistent format
                
                    if ($tWMI_test) {
		                 try {$tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'") -ErrorAction STOP }
                         catch{$tWMI = $null}
                    }else{$tWMI = $null}
		            if ($tWMI)
		            {
			            $MailboxStatistics = $tWMI | Select @{Name="DisplayName";Expression={$_.MailboxDisplayName}},@{Name="TotalItemSizeB";Expression={$_.Size}},@{Name="TotalDeletedItemSizeB";Expression={$_.DeletedMessageSizeExtended }},@{Name="Database";Expression={((get-mailboxdatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").identity)}}
		            } else {
                        Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot retrieve Mailbox Statistics via WMI for $($ExchangeServer.Name)"

			        
			            $MailboxStatistics = $null
		            }
	            }	
	            # Exchange 2000
	            if ($ExchangeMajorVersion -eq "6.0")
	            {
		            # Mailbox Count
		            $MailboxCount = _GetExSvrMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
		            # Get Role via ADSI
		            $tADSI=[ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"
		            if ($tADSI)
		            {
			            if ($tADSI.ServerRole -eq 1) { $Roles=@("FE") } else { $Roles=@("BE") }
		            } else {
			            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): [_GetExSvr] Error : Cannot detect Front End/Back End Server information via ADSI for $($ExchangeServer.Name)"

			            $Roles+="Unknown"
		            }
		            $MailboxStatistics = $null
	            }

            #endregion Handling Pre Exch 2007                   


            #region Mailbox Database copy Info

                $DBs_Mounted         = @()
                $DBs_Mountedcount    = 0 
                $DBcopies            = 0
                $DBCopyCount         = 0

                if ($Roles -contains "Mailbox" -and $E2010) {  
               
                    $srv = $ExchangeServer.Name

                    #Getting DBs mounted on this server
                    $DBs_Mounted        = @($databases | Where {$_.Server -ieq $srv})
                    $DBs_MountedCount    =  $DBs_Mounted.count

                    #Getting DB Copy info
                    try {
                        $MailboxServer   = Get-MailboxServer $Srv -ErrorAction Stop
                        $DBCopy = $MailboxServer | Get-MailboxDatabaseCopyStatus -ErrorAction Stop

                        if ($DBCopy ) {
                            $DBCopyCount   = $DBCopy.Count
                        }
                    }catch {

                    }
                } # if ($Roles -contains "Mailbox" -and $E2010) 

            #endregion Mailbox Database copy Info


            #region  Filling info for chart

                if ($Roles -contains "Mailbox") {
                    $chartExObj = New-Object -TypeName PSObject -Property $chart_srv_objprop
                    $chartExObj.Name = $ExchangeServer.Name.ToUpper()            
                    if ($MailboxCount) {
                        $chartExObj.MailboxCount = [decimal]::round($MailboxCount)
                    } else {
                    $chartExObj.MailboxCount = 0
                    }
                    $chartExObj.DBMountedCount = $DBs_Mountedcount

                    $ExchangeEnvironment.Chart_Srv += $chartExObj
                }
            


            #endregion Filling info for chart
            

            #region Log Info 
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Name : $($ExchangeServer.Name.ToUpper())"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - RealName :$RealName"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExchangeMajorVersion : $ExchangeMajorVersion "
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExchangeSPLevel : $ExchangeSPLevel"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Edition :$($ExchangeServer.Edition)"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Mailboxes (count) : $MailboxCount"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - OSVersion : $OSVersion"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - OSServicePack : $OSServicePack"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Roles : $Roles"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - RollupLevel : $RollupLevel"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - RollupVersion : $RollupVersion"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Site: $($ExchangeServer.Site.Name)"            
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - IntNames : $IntNames"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExtNames : $ExtNames"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - CASArrayName : $CASArrayName"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExServicesHealth : $ExServicesHealth"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExServicesFailCount : $ExServicesFailCount"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ExUpTime  : $ExUpTime "
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DBCopies  : $DBCopyCount "
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DBsMounted  : $DBs_MountedCount "
            #endregion Log Info 
	

	        # Return Hashtable
	        @{Name					= $ExchangeServer.Name.ToUpper()
	         RealName				= $RealName
	         ExchangeMajorVersion 	= $ExchangeMajorVersion
	         ExchangeSPLevel		= $ExchangeSPLevel
	         Edition				= $ExchangeServer.Edition
	         Mailboxes				= $MailboxCount
	         OSVersion				= $OSVersion;
	         OSServicePack			= $OSServicePack
	         Roles					= $Roles
	         RollupLevel			= $RollupLevel
	         RollupVersion			= $RollupVersion
	         Site					= $ExchangeServer.Site.Name
	         MailboxStatistics		= $MailboxStatistics
	         Disks					= $Disks
             IntNames				= $IntNames
             ExtNames				= $ExtNames
             CASArrayName			= $CASArrayName
             ExServicesHealth       = $ExServicesHealth 
             ExServicesFailCount    = $ExServicesFailCount
             ExUpTime               = $ExUpTime    
             DBCopyCount            = $DBCopyCount
             DBMounted              = $DBs_Mountedcount    

	        }
        	
    } #function _GetExSvr


        function _GetDB {

	         # This function is based completely on Steve's Goodman Script with couple of modifications here and there
            param($Database,$ExchangeEnvironment,$Mailboxes,$ArchiveMailboxes,$E2010)
	
	    
            #region log info
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): +++++++Entering _GetDB function with DB : ""$($Database.name)"" and E2010 value : $E2010"
            #endregion log info 


            #region additional checks

    
                #Mount status check

                    #if database is not mounted, then the DatabaseSize (DB Size) and AvailableNewMailboxSpace (WhiteSpace) values will not be available   
                    $Mounted = $Database.Mounted 
            
                #Recovery db?

                    $IsRecovery = $Database.Recovery

                #Activation preference check

                    $DatabasePreCheck = "Not Available"
                    $DB_Act_pref = "Not Available"
                    $DatabasePereferenceServer = "Not Available"
                    $DatabasePereference = "Not Available"
                    $DatabaseNowSRV = "Not Available"

                    if ($Database.ActivationPreference) {
    
                        $DB_Act_pref = $Database.ActivationPreference
                        $DatabasePereferenceServer = $DB_Act_pref |Where {$_.value -eq 1}
                        $DatabasePereference =$DatabasePereferenceServer.key.name
                        $DatabaseNowSRV = $Database.Server.name
                        If ( $DatabasePereference -ne $DatabaseNOWSRV ) {
	                        $DatabasePreCheck = "Fail"
                        }
                        else {
                        $DatabasePreCheck = "Pass"
                        }        
                     } # if ($Database.ActivationPreference)

    
          

               #Active owner

                   if(($Database.Server.Name)) {
                        $ActiveOwner = $Database.Server.Name.ToUpper()
                   }else {
                        $ActiveOwner ="Not Available"	
                   }    
       

               #Getting Send and Receive quota

                
                    if ($Database.ProhibitSendReceiveQuota -ne $null) {

                        $Get_DB_ProhibitSendReceiveQuota_isunlimited = $Database.ProhibitSendReceiveQuota.isunlimited
                

                        if ($Get_DB_ProhibitSendReceiveQuota_isunlimited -eq $true) {
	                         $DatabasePSR = "Unlimited"
                        }
	                    else {
	                         $DatabasePSR = ($Database.ProhibitSendReceiveQuota.Value.ToMB() ) /1024
                       }

                    } # if ($Database.ProhibitSendReceiveQuota)
                    else {
                        $DatabasePSR = "Not Available" 
                    }


                #Getting Send and Prohibit Send quota

                    if ($Database.ProhibitSendQuota.isunlimited -ne $null) {
                        $Get_DB_ProhibitSendQuota_isunlimited  = $Database.ProhibitSendQuota.isunlimited

                        if ($Get_DB_ProhibitSendQuota_isunlimited -eq $true){
	                         $DatabasePS = "Unlimited"
                        }
	                    else {
	                        $DatabasePS = ($Database.ProhibitSendQuota.Value.ToMB() ) /1024
                        }


                    } #if ($Database.ProhibitSendQuota.isunlimited)
                    else {
                        $DatabasePS = "Not Available" 
                    }

            

                #Dumpster info

                    if($Database.RecoverableItemsWarningQuota) {
                       $DatabaseDumpsterWQ = ($Database.RecoverableItemsWarningQuota.Value.ToMB() )
                    } else {
                        $DatabaseDumpsterWQ ="Not Available" 
                    }

                    if ($Database.RecoverableItemsQuota) {
                        $DatabaseDumpsterQ = ($Database.RecoverableItemsQuota.Value.ToMB() )
                    }else {
                        $DatabaseDumpsterQ = "Not Available"
                    }

                #Database Servers

                    $DBHolders =$null 

                    if ($Database.Servers) {
                       [array]$DBHolders =$null 
		                    ( $Database.Servers) |%{$DBHolders  += $_.name}
                    }

                #DAG Membership

                 $DatabaseDagMembership = "Not Available"  
                 if ($Database.MasterServerOrAvailabilityGroup) {
            
                    $DatabaseDagMembershipValue = $Database.MasterServerOrAvailabilityGroup.name
	
	                if ($DatabaseDagMembershipValue -match "DAG") {
		                $DatabaseDagMembership = $DatabaseDagMembershipValue
                     }
                    else {
	                    $DatabaseDagMembership = "Not Available" 
                    }
         
                 }else {
                    $DatabaseDagMembership  = "Not Available"
                 }  


            #endregion additional checks
       
    
            #region Circular Logging
	        if ($Database.CircularLoggingEnabled) { $CircularLoggingEnabled="Yes" } else { $CircularLoggingEnabled = "No" }
	        #endregion Circular Logging


            #region DB Backup
    
            $LastFullBackup = "Not Available"
            $HowOldBackup = "Not Available"
            if ($Database.LastFullBackup) { 
                #string value of last backup
                $LastFullBackup=$Database.LastFullBackup.ToString() 
                #datetime value of last backup
                $lastBackup =  $Database.LastFullBackup

                $currentDate = Get-Date

                $HowOldBackup= $currentDate - $lastBackup     
                $HowOldBackup= $HowOldBackup.days     

        
            } else { 
                $LastFullBackup = "Not Available"
                $HowOldBackup = "Not Available"
            }
    
	        #endregion DB Backup
    

	        #region  Mailbox Average Sizes

	        $MailboxStatistics = [array]($ExchangeEnvironment.Servers[$Database.Server.Name].MailboxStatistics | Where {$_.Database -eq $Database.Identity})
	        $MailboxAverageSize = 0
            $MailboxItemSizeB  = 0
            if ($MailboxStatistics)
	        {
		        [long]$MailboxItemSizeB = 0
		        $MailboxStatistics | %{ $MailboxItemSizeB+=$_.TotalItemSizeB }
		        [long]$MailboxAverageSize = $MailboxItemSizeB / $MailboxStatistics.Count
	        } else {
		        $MailboxAverageSize = 0
                $MailboxItemSizeB  = 0
	        }

	        #endregion  Mailbox Average Sizes


	        #region Free Disk Space Percentage
	
            $FreeLogDiskSpace=$null
	        $FreeDatabaseDiskSpace=$null
            if ($ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
	        {
		        foreach ($Disk in $ExchangeEnvironment.Servers[$Database.Server.Name].Disks)
		        {
			        if ($Database.EdbFilePath.PathName -like "$($Disk.Name)*")
			        {
				        $FreeDatabaseDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
			        }
			        if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14)
			        {
				        if ($Database.LogFolderPath.PathName -like "$($Disk.Name)*")
				        {
					        $FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
				        }
			        } else {
				        $StorageGroupDN = $Database.DistinguishedName.Replace("CN=$($Database.Name),","")
				        $Adsi=[adsi]"LDAP://$($Database.OriginatingServer)/$($StorageGroupDN)"
				        if ($Adsi.msExchESEParamLogFilePath -like "$($Disk.Name)*")
				        {
					        $FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
				        }
			        }
		        }
	        } else {
		        $FreeLogDiskSpace=$null
		        $FreeDatabaseDiskSpace=$null
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Skipping Disk Info because getting WMI info from server failed previously"
            
	        }
	
            #endregion Free Disk Space Percentage


            #region Archive, whitespace and CopyCount
    
	        $CopyCount = 0
            $Copies = @()
            $ArchiveMailboxCount = 0
            [long]$ArchiveItemSizeB = 0
            $StorageGroup = $null
            $Size = 0
            $Whitespace = 0
            $ArchiveAverageSize = 0

            if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14 -and $E2010)
	        {
		        # Exchange 2010 Database Only
		        $CopyCount = [int]$Database.Servers.Count
		        if ($Database.MasterServerOrAvailabilityGroup.Name -ne $Database.Server.Name)
		        {
			        $Copies = [array]($Database.Servers | % { $_.Name })
		        } else {
			        $Copies = @()
		        }
		        # Archive Info
		        $ArchiveMailboxCount = [int]([array]($ArchiveMailboxes | Where {$_.ArchiveDatabase -eq $Database.Name})).Count
                $ArchiveStatistics = [array]($ArchiveMailboxes | Where {$_.ArchiveDatabase -eq $Database.Name} | Get-MailboxStatistics -Archive )
		        if ($ArchiveStatistics)
		        {
			        [long]$ArchiveItemSizeB = 0
			        $ArchiveStatistics | %{ $ArchiveItemSizeB+=$_.TotalItemSize.Value.ToBytes() }
			        [long]$ArchiveAverageSize = $ArchiveItemSizeB / $ArchiveStatistics.Count
		        } else {
			        $ArchiveAverageSize = 0
		        }
		        # DB Size / Whitespace Info
                if($Database.DatabaseSize){
		            [long]$Size = $Database.DatabaseSize.ToBytes()
                }

                if ($Whitespace = $Database.AvailableNewMailboxSpace) {
		            [long]$Whitespace = $Database.AvailableNewMailboxSpace.ToBytes()
                }
		        $StorageGroup = $null
		
	        } else {
		        $ArchiveMailboxCount = 0
		        $CopyCount = 0
		        $Copies = @()
		        # 2003 & 2007, Use WMI (Based on code by Gary Siepser, http://bit.ly/kWWMb3)
		        $Size = [long](get-wmiobject cim_datafile -computername $Database.Server.Name -filter ('name=''' + $Database.edbfilepath.pathname.replace("\","\\") + '''')).filesize
		        if (!$Size)
		        {
			    
                    Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): Cannot detect database size via WMI for $($Database.Server.Name)"
			        [long]$Size = 0
			        [long]$Whitespace = 0
		        } else {
			        [long]$MailboxDeletedItemSizeB = 0
			        if ($MailboxStatistics)
			        {
				        $MailboxStatistics | %{ $MailboxDeletedItemSizeB+=$_.TotalDeletedItemSizeB }
			        }
			        $Whitespace = $Size - $MailboxItemSizeB - $MailboxDeletedItemSizeB
			        if ($Whitespace -lt 0) { $Whitespace = 0 }
		        }
		        $StorageGroup =$Database.DistinguishedName.Split(",")[1].Replace("CN=","")
	         }
              #endregion Archive, whitespace and CopyCount

              #region Mailboxcount

              $dbMailboxcount = [long]([array]($Mailboxes | Where {$_.Database -eq $Database.Identity})).Count

              #endregion Mailboxcount

              #region Filling chart info
            
            
                if (!$IsRecovery) {
                    $chartDBObj = New-Object -TypeName PSObject -Property $chart_db_objprop
            
                    $chartDBObj.Name = $Database.Name.ToUpper()

                    if ($Size) {
                        $chartDBObj.DBSize = [decimal]::round(($Size / 1GB))
                    }else {
                        $chartDBObj.DBSize = 0
                    }
            
                    $chartDBObj.MailboxCount = $dbMailboxcount
            
                    if ($HowOldBackup -notlike "Not Available") {
                        $chartDBObj.Backup = $HowOldBackup
                    }else {
                        $chartDBObj.Backup = -1
                    }

                    $ExchangeEnvironment.Chart_db += $chartDBObj

                }

              #endregion Filing chart info

        
            #region log info

            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Name : $($Database.Name.ToUpper())"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - StorageGroup : $StorageGroup "
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ActiveOwner : $ActiveOwner"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Mailboxcount : $dbMailboxcount"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - MailboxAverageSize	: $($MailboxAverageSize / 1MB) MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ArchiveMailboxCount : $ArchiveMailboxCount"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ArchiveTotalSize : $($ArchiveItemSizeB / 1MB) MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - ArchiveAverageSize : $($ArchiveAverageSize  / 1MB) MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - CircularLoggingEnabled : $CircularLoggingEnabled"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - LastFullBackup	: $LastFullBackup"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Size : $($Size / 1GB) GB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Whitespace : $($Whitespace / 1MB) MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Copies : $($Copies | % { $_})"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - CopyCount : $CopyCount"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - FreeLogDiskSpace : $FreeLogDiskSpace %"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - FreeDatabaseDiskSpace : $FreeDatabaseDiskSpace %"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DBPreCheck : $DatabasePreCheck"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - HowOldBackup : $HowOldBackup Days"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabaseDumpsterWQ : $DatabaseDumpsterWQ MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabaseDumpsterQ : $DatabaseDumpsterQ MB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabasePS (Prohibit Send) : $DatabasePS GB "
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabasePSR (Prohibit Send and Rec) : $DatabasePSR GB"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - Mounted : $Mounted"        
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabasePereferredServer : $DatabasePereference"
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - IsRecovery : $($Database.Recovery)"        
            Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): - DatabaseDagMembership : $DatabaseDagMembership"

            #endregion log info
	
	        @{Name						= $Database.Name
	          StorageGroup				= $StorageGroup
	          ActiveOwner				= $ActiveOwner
	          MailboxCount				= [long]([array]($Mailboxes | Where {$_.Database -eq $Database.Identity})).Count
	          MailboxAverageSize		= $MailboxAverageSize
	          ArchiveMailboxCount		= $ArchiveMailboxCount
              ArchiveTotalSize          = $ArchiveItemSizeB
	          ArchiveAverageSize		= $ArchiveAverageSize
	          CircularLoggingEnabled 	= $CircularLoggingEnabled
	          LastFullBackup			= $LastFullBackup
	          Size						= $Size
	          Whitespace				= $Whitespace
	          Copies					= $Copies
	          CopyCount					= $CopyCount
	          FreeLogDiskSpace			= $FreeLogDiskSpace
	          FreeDatabaseDiskSpace		= $FreeDatabaseDiskSpace
              DBPreCheck                = $DatabasePreCheck
              HowOldBackup              = $HowOldBackup
              DatabaseDumpsterWQ        = $DatabaseDumpsterWQ
              DatabaseDumpsterQ         = $DatabaseDumpsterQ
              DatabasePS                = $DatabasePS 
              DatabasePSR               = $DatabasePSR
              Mounted                   = $Mounted
              DB_Act_pref               = $DB_Act_pref 
	          DatabasePereferredServer  = $DatabasePereference
	          IsRecovery                = $Database.Recovery
              DBHolders			        = $DBHolders
              DatabaseDagMembership     = $DatabaseDagMembership
	          }

    } # function _GetDB


    #endregion non helper functions


    #region GUI functions

   
        function _GetOverview {
            
            # This function is based completely on Steve's Goodman Script with couple of modifications here and there
	        param($Servers,$ExchangeEnvironment,$ExRoleStrings,$Pre2007=$False)
	        if ($Pre2007)
	        {
		        $BGColHeader="#880099"
		        $BGColSubHeader="#8800CC"
		        $Prefix=""
                $IntNamesText=""
                $ExtNamesText=""
                $CASArrayText=""
	        } else {
		        $BGColHeader="#000099"
		        $BGColSubHeader="#0000FF"
		        $Prefix="Site:"
                $IntNamesText=""
                $ExtNamesText=""
                $CASArrayText=""
                $IntNames=@()
                $ExtNames=@()
                $CASArrayName=""
                foreach ($Server in $Servers.Value)
                {
                    $IntNames+=$Server.IntNames
                    $ExtNames+=$Server.ExtNames
                    $CASArrayName=$Server.CASArrayName
            
                }
                $IntNames = $IntNames|Sort -Unique
                $ExtNames = $ExtNames|Sort -Unique
                if($IntNames) {$IntNames = [system.String]::Join(",",$IntNames)}
                if($ExtNames){$ExtNames = [system.String]::Join(",",$ExtNames)}
                if ($IntNames)
                {
                    $IntNamesText="Internal Names: $($IntNames)"
                    $ExtNamesText="External Names: $($ExtNames)<br >"
                }
                if ($CASArrayName)
                {
                    $CASArrayText="CAS Array: $($CASArrayName)"
                }
	        }
	        $Output="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
	        <col width=""20%""><col width=""20%"">
	        <colgroup width=""25%"">";
	
	        $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort Name| %{$Output+="<col width=""3%"">"}

	        #$Output+="</colgroup><col width=""20%""><col  width=""20%"">
            $Output+="</colgroup><col width=""10%""><col  width=""10%""><col  width=""10%""><col  width=""10%"">

	        <tr bgcolor=""$($BGColHeader)""><th><font color=""#ffffff"">$($Prefix) $($Servers.Key)</font></th>
	        <th colspan=""$(($ExchangeEnvironment.TotalServersByRole.Count)+2)"" align=""left""><font color=""#ffffff"">$($ExtNamesText)$($IntNamesText)</font></th>
	        <th align=""center""><font color=""#ffffff"">$($CASArrayText)</font></th><th></th><th></th></tr>"
	        $TotalMailboxes=0
	        $Servers.Value | %{$TotalMailboxes += $_.Mailboxes}
	        $Output+="<tr bgcolor=""$($BGColSubHeader)""><th><font color=""#ffffff"">Mailboxes: $($TotalMailboxes)</font></th><th>"
            $Output+="<font color=""#ffffff"">Exchange Version</font></th>"
	        $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort Name| %{$Output+="<th><font color=""#ffffff"">$($ExRoleStrings[$_.Key].Short)</font></th>"}
	        $Output+="<th><font color=""#ffffff"">OS Version</font></th>
            <th><font color=""#ffffff"">OS Service Pack</font></th>

            <th><font color=""#ffffff"">UP Time (Days)</font></th>

            <th><font color=""#ffffff"">Exch Service Health</font></th> 
    
    
    
            </tr>"
	        $AlternateRow=0


	
	        foreach ($Server in $Servers.Value)
	        {
		        $Output+="<tr "
		        if ($AlternateRow)
		        {
			        $Output+=" style=""background-color:#dddddd"""
			        $AlternateRow=0
		        } else
		        {
			        $AlternateRow=1
		        }
		        $Output+="><td>$($Server.Name)"
		        if ($Server.RealName -ne $Server.Name)
		        {
			        $Output+=" ($($Server.RealName))"
		        }
		        $Output+="</td><td>$($ExVersionStrings["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].Long)"
		        if ($Server.RollupLevel -gt 0)
		        {
			        $Output+=" UR$($Server.RollupLevel)"
			        if ($Server.RollupVersion)
			        {
				        $Output+=" $($Server.RollupVersion)"
			        }
		        }
		        $Output+="</td>"
		        $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort Name| %{ 
			        $Output+="<td"
			        if ($Server.Roles -contains $_.Key)
			        {
				        $Output+=" align=""center"" style=""background-color:#00FF00"""
			        }
			        $Output+=">"
			        if (($_.Key -eq "ClusteredMailbox" -or $_.Key -eq "Mailbox" -or $_.Key -eq "BE") -and $Server.Roles -contains $_.Key) 
			        {
				        $Output+=$Server.Mailboxes
			        } 
		        }
				
		        $Output+="<td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td>";	



                # Up time

                if ($server.ExUpTime -ne $null -or $server.ExUpTime -eq 0 ) {
                        if ($server.ExUpTime -le $UptimeErrorThreshold) {
                            $Output+="<td align=""center"">$($Server.ExUpTime)</td>";
                        }else {
                            $Output+="<td align=""center"" style=""background-color:$($ErrorColor)""><font color=""#FFFFFF"" >$($Server.ExUpTime)</font></td>";
                        }
                
        
                } else {
        
                $Output+="<td align=""center""> N/A</td>";

                }	

                # Service Health


                if ($server.ExServicesHealth -notlike "N/A" ) {
                    $Output+="<td>$($Server.ExServicesHealth)</td>";
        
                } else {        
                $Output+="<td>$($Server.ExServicesHealth)</td></tr>";

                }	

	        }
	        $Output+="<tr></tr>
	        </table><br />"
	        $Output
        } # function _GetOverview


        function _GetDBTable {

            # This function is based completely on Steve's Goodman Script with couple of modifications here and there
	        param($Databases)
	        # Only Show Archive Mailbox Columns, Backup Columns and Circ Logging if at least one DB has an Archive mailbox, backed up or Cir Log enabled.
	        $ShowArchiveDBs=$False
	        $ShowLastFullBackup=$False
	        $ShowCircularLogging=$False
	        $ShowStorageGroups=$False
	        $ShowCopies=$False
	        $ShowFreeDatabaseSpace=$False
	        $ShowFreeLogDiskSpace=$False
	        foreach ($Database in $Databases)
	        {
		        if ($Database.ArchiveMailboxCount -gt 0) 
		        {
			        $ShowArchiveDBs=$True
		        }
		        if ($Database.LastFullBackup -ne "Not Available") 
		        {
			        $ShowLastFullBackup=$True
		        }
		        if ($Database.CircularLoggingEnabled -eq "Yes") 
		        {
			        $ShowCircularLogging=$True
		        }
		        if ($Database.StorageGroup) 
		        {
			        $ShowStorageGroups=$True
		        }
		        if ($Database.CopyCount -gt 0) 
		        {
			        $ShowCopies=$True
		        }
		        if ($Database.FreeDatabaseDiskSpace -ne $null)
		        {
			        $ShowFreeDatabaseSpace=$true
		        }
		        if ($Database.FreeLogDiskSpace -ne $null)
		        {
			        $ShowFreeLogDiskSpace=$true
		        }
	        }
	
	
	        $Output="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
	
	        <tr align=""center"" bgcolor=""#FFD700"">
	        <th>Server</th>"

    

	        if ($ShowStorageGroups)
	        {
		        $Output+="<th>Storage Group</th>"
	        }
	        $Output+="<th>Database</th>
	        <th>Mailboxes</th>
	        <th>Av. Mailbox Size MB</th>"

            #region Adding Mount Status

            $Output+="<th>Mounted</th>"

            #endregion Adding Mount Status


	        if ($ShowArchiveDBs)
	        {
		        $Output+="<th>Archive MBs</th><th>Av. Archive Size MB</th>"
	        }
	        $Output+="<th>DB Size GB</th><th>Whitespace GB</th>"
	        if ($ShowFreeDatabaseSpace)
	        {
		        $Output+="<th>DB Disk Free</th>"
	        }
	        if ($ShowFreeLogDiskSpace)
	        {
		        $Output+="<th>Log Disk Free</th>"
	        }
	        if ($ShowLastFullBackup)
	        {
		        $Output+="<th>Last Full Backup</th>"
                $Output+="<th>Backup Since (days)</th>"
	        }
	        if ($ShowCircularLogging)
	        {
		        $Output+="<th>Circular?</th>"
	        }
    

            #region Adding Quota Info

            #Prohibi Send
            $Output+="<th>Quota PS GB</th>"
            #Prohibit Send and Receive
            $Output+="<th>Quota PSR GB</th>"

            #endregion Adding Quota Info

            #region Adding Preference Check

            $Output+="<th>DB Pref Check</th>"
            #$Output+="<th>DB Pref srv</th>"
            $Output+="<th>DB Act Pref</th>"

            #endregion Adding Preference Check 

	        <#if ($ShowCopies)
	        {
		        $Output+="<th>Copies (n)</th>"
	        }#>

    
	
	        $Output+="</tr>"
	        $AlternateRow=0;
	        foreach ($Database in $Databases)
	        {
		        $Output+="<tr"
		        if ($AlternateRow)
		        {
			        $Output+=" style=""background-color:#dddddd"""
			        $AlternateRow=0
		        } else
		        {
			        $AlternateRow=1
		        }
		
		        $Output+="><td>$($Database.ActiveOwner)</td>"
		        if ($ShowStorageGroups)
		        {
			        $Output+="<td>$($Database.StorageGroup)</td>"
		        }
		        $Output+="<td><Strong>$($Database.Name)</strong></td>
		        <td align=""center"">$($Database.MailboxCount)</td>
		        <td align=""center"">$("{0:N2}" -f ($Database.MailboxAverageSize/1MB))</td>"


                #region Adding Mount Status
        
                if ($Database.Mounted -eq $true) {
                $Output+="<td align=""center"" style=""background-color:$($greenColor)"">$($Database.Mounted)</td>"
                }else{ $Output+="<td align=""center"">$($Database.Mounted)</td>"}

        

                 #endregion Adding Mount Status


		        if ($ShowArchiveDBs)
		        {
			        $Output+="<td align=""center"">$($Database.ArchiveMailboxCount)</td> 
			        <td align=""center"">$("{0:N2}" -f ($Database.ArchiveAverageSize/1MB))</td>";
		        }
		    
                if ($Database.Size) {
                    $Output+="<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) </td>"
                }else {
                    $Output+="<td align=""center"">N/A</td>"
                }

                if ($Database.Whitespace) {
                    $Output+="<td align=""center"">$("{0:N2}" -f ($Database.Whitespace/1GB))</td>";
                }else{
                    $Output+="<td align=""center"">N/A</td>";
                }




		        if ($ShowFreeDatabaseSpace)
		        {
			        $Output+="<td align=""center"">$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
		        }
		        if ($ShowFreeLogDiskSpace)
		        {
			        $Output+="<td align=""center"">$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
		        }
		        if ($ShowLastFullBackup)
		        {
			        if ($Database.LastFullBackup -notlike "Not Available"){
                        $Output+="<td align=""center"">$($Database.LastFullBackup)</td>";
                    }else {
                        $Output+="<td align=""center"">N/A</td>";
                    }


                    if ($Database.HowOldBackup -like "Not Available"){
                        $Output+="<td align=""center"">N/A</td>";
                    }else {


                        #region coloring backup
                            if ( ($Database.HowOldBackup) -ge $BackupError) {
                                $Output+="<td align=""center"" style=""background-color:$($ErrorColor)""><font color=""#FFFFFF"" >$($Database.HowOldBackup)</font></td>";
                            }elseif(($Database.HowOldBackup) -ge $BackupWarning){
                                $Output+="<td align=""center"" style=""background-color:$($warningColor)""><font color=""#FFFFFF"" >$($Database.HowOldBackup)</font></td>";
                            }else{
                                $Output+="<td align=""center"">$($Database.HowOldBackup)</td>";
                            }
                        #endregion coloring backup
                
                        }  
            
		        }


		        if ($ShowCircularLogging)
		        {
			        $Output+="<td align=""center"">$($Database.CircularLoggingEnabled)</td>";
		        }
        
        
                #region Adding Quota Info

            

                    $Output+="<td align=""center"">$("{0:N2}" -f ($Database.DatabasePS))</td>"
                    $Output+="<td align=""center"">$("{0:N2}" -f ($Database.DatabasePSR))</td>"        

                #endregion Adding Quota Info

                #region Adding Preference Check

                    if($Database.DBPreCheck -like "Pass"){
                        $Output+="<td align=""center""><font color=""#339933"" ><strong>$($Database.DBPreCheck)</strong></font></td>"
                    }else {

                    $Output+="<td align=""center""><font color=""$($errorcolor)"" ><strong>$($Database.DBPreCheck)</strong></font></td>"
                    }

            

                    #$Output+="<td align=""center"">$($Database.DatabasePereferredServer)</td>"
                    $Output+="<td align=""center"">$($Database.DB_Act_pref)</td>"

                #endregion Adding Preference Check

        


		        <#if ($ShowCopies)
		        {
			        $Output+="<td>$($Database.Copies|%{$_}) ($($Database.CopyCount))</td>"
		        }#>


		        $Output+="</tr>";
	        }
	        $Output+="</table><br />"
	
	        $Output

      } # function _GetDBTable


        function _GetRecoveryDBTable {

	        param($Databases)
	
     
	        # Drawing Table header
	
	        $Output="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">	
	        <tr align=""center"" bgcolor=""#FFD700"">"
	
	        $Output+="<th>#</th><th>Database Name</th><th>Mounted</th><th>Server</th>"
	
	        $Output+="<th>DB Size</th>"
	
	        $Output+="<th>DB Disk Free</th>"
	        $Output+="<th>Log Disk Free</th>"
		
	        $Output+="</tr>"
	        $AlternateRow=0;
	
	
	        #Writing Table content
	
	        foreach ($Database in $Databases) {

		        $C++

		        $Output+="<tr"

		        if ($AlternateRow) {
		
			        $Output+=" style=""background-color:#dddddd"""
			        $AlternateRow=0
		        } else {
		
			        $AlternateRow=1
		        }
		
		        $Output+=">"
	
	            $Output+="<td  align=""center""><strong>$C</strong></td>"
		        $Output+="<td  align=""center""><strong>$($Database.Name)</strong></td>"
		
		        if($Database.Mounted -eq $True) {
		            $Output+="<td align=""center""><font color=""#008000""><Strong>$($Database.Mounted)</Strong></font></td>"
		            $Output+="<td align=""center""><font color=""#000000"">$($Database.ActiveOwner)</font></td>"
		            $Output+="<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>"
                }
		        else {
		            $Output+="<td align=""center""><font color=""#000000""><Strong>$($Database.Mounted)</Strong></font></td>"
			        $Output+="<td align=""center""><font color=""#00000"">$($Database.ActiveOwner)</font></td>"
                    $Output+="<td align=""center""> -</td>"  
               }	
		 
			        $Output+="<td align=""center"">$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
			        $Output+="<td align=""center"">$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"	
	                $Output+="</tr>";
	        }


	        $Output+="</table><br />"	
	        $Output

    } # function _GetRecoveryDBTable


        function _GetDBPreference_Table_Info {

            param($ExchangeEnvironment,$myDAGObj)

            $my_DAG_dbs = $myDAGObj.databases

            $array_server_vs_db = @()
            $array_server_vs_db_Aggregates = @()
            $hash_server_vs_db_Aggregates = @{}

            #$hash_server_vs_db
        
            foreach ($DAGDB in $my_DAG_dbs) {

                #$hash_server_vs_db.Add($DAGDB.Name, ($DAGDB.DatabasePereferredServer))

                $obj = New-Object -TypeName PSObject 
                $obj | Add-Member -type NoteProperty -name Name  -value $DAGDB.Name
                $obj | Add-Member -type NoteProperty -name count -value $DAGDB.DatabasePereferredServer
                $array_server_vs_db += $obj

            } # foreach ($DAGDb in $my_DAG_dbs)

            $array_server_vs_db_Aggregates = $array_server_vs_db |Group-Object -Property count

            foreach ($server in $array_server_vs_db_Aggregates) {
        
                $hash_server_vs_db_Aggregates.Add($server.name, $server.count) 

            } # foreach ($server in $array_server_vs_db_Aggregates)

            $servers_in_dag = $myDAGObj.members
        
        
            #this loop will counter for a case where there is a server in the DAG with no mounting or prefered DBs on it
            foreach ($server in $servers_in_dag  ) {

                if ( ! ($hash_server_vs_db_Aggregates.ContainsKey($server) )) {
                    $hash_server_vs_db_Aggregates.Add($server,0)
                } # if

            } # foreach ($server in $servers_in_dag  )

        
            <# Output is a hash table containing as a key server name, and as a value, number of databases hosted in it.
            #>
            Write-Output $hash_server_vs_db_Aggregates 


    } # function _GetDBPreference_Table_Info 


        function _GetDAG_DB_Layout  {
        
            param($databases,$DAG,$ExchangeEnvironment, $hash_server_vs_db_Aggregates)

            $WarningColor                      = "#FF9900"
		    $ErrorColor                        = "#980000"
		    $BGColHeader                       = "#000099"
		    $BGColSubHeader                    = "#0000FF"
		    [Array]$Servers_In_DAG             = $DAG.Members

            $Output2 ="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
	        <col width=""5%"">
	        <colgroup width=""25%"">"
	        $Servers_In_DAG | Sort-Object | %{$Output2+="<col width=""3%"">"}
	        $Output2 +="</colgroup>"
	        $ServerCount = $Servers_In_DAG.Count
	
	        $Output2 += "<tr bgcolor=""$($BGColHeader)""><th><font color=""#ffffff"">DatabaseCopies</font></th>	
	        <th colspan=""$($ServerCount)""><font color=""#ffffff"">Mailbox Servers in $($DAG.name)</font></th>	
	        </tr>"
	        $Output2+="<tr bgcolor=""$($BGColSubHeader)""><th></th>"
	        $Servers_In_DAG |Sort-Object | %{$Output2+="<th><font color=""#ffffff"">$($_)</font></th>"}
	
	        $Output2 += "</tr>"

            #region Table Part 1 : Preference info
	            $AlternateRow=0
		        foreach ($Database in $Databases) {
	    
	                $Output2+="<tr "
	                if ($AlternateRow) {
					    
				        $Output2+=" style=""background-color:#dddddd"""
				        $AlternateRow=0
			        } else {
			            $AlternateRow=1
			        }
		
		            $Output2+="><td><strong>$($database.name)</strong></td>"

                    $DatabaseServer   = $Database.ActiveOwner
			        $DatabaseServers  = $Database.DBHolders

		            $Servers_In_DAG|Sort-Object| % { 
					        $ActvPref =$Database.DB_Act_pref
					        $server_in_the_loop = $_
					        $Actv = $ActvPref  |where {$_.key -eq  $server_in_the_loop}
					        $Actv=  $Actv.value
					        $ActvKey= $ActvPref |Where {$_.value -eq 1}
					        $ActvKey = 	 $ActvKey.key.name
									  
									  
					        $Output2+="<td"
							
					        if (  ($DatabaseServers -contains $_) -and ( $_ -like $databaseserver)  ) {
										
						                if ($ActvKey -like $databaseserver ) {
							    $Output2+=" align=""center"" style=""background-color:#F7FB0B""><font color=""#000000f""><strong>$Actv</strong></font> "
                            }
						                else {
							    $Output2+=" align=""center"" style=""background-color:#FB0B1B""><strong><font color=""#ffffff"">$Actv</strong></font> "
                            }
										
					        } elseif ($DatabaseServers -contains $_) {
							        $Output2+=" align=""center"" style=""background-color:#00FF00"">$Actv "							 
					         }else {
							        $Output2+=" align=""center"" style=""background-color:#dddddd"">"	
                            }
								 
			
			          } # $Servers_In_DAG|Sort-Object|
				
		
		
		            $Output2+="</tr >"
		        } # foreach ($Database in $Databases) 
		
	            $Output2+="<tr></tr><tr></tr><tr></tr>"

            #endregion Table Part 1 : Preference info

            #region Table Part 2 : Aggregates info

                #Total Assigned copies	
	            $Output2 += "<tr bgcolor=""#440164""><th><font color=""#ffffff"">Total Copies</font></th>"	
	            $Servers_In_DAG|Sort-Object| 
			        %{ 
		                $srv = $ExchangeEnvironment.Servers[$_]
					    $Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($srv.DBCopyCount)</strong></font></td>"	
						
		             }
	            $Output2 +="</tr>"
            
                #Copies Assigned Ideal	
	            $Output2 += "<tr bgcolor=""#DB08CD""><th><font color=""#ffffff"">Ideal Mounted DB Copies</font></th>"	
	            $Servers_In_DAG|Sort-Object| 
			    %{ 
			        foreach ($srv in $hash_server_vs_db_Aggregates.GetEnumerator()) {
										
					    if ($srv.key  -like $_) {
						    $Output2 += "<td align=""center"" style=""background-color:#FBCCF9""><font color=""#000000""><strong>$($srv.value)</strong></font></td>"
                        }		
				    }
		        }
	            $Output2 +="</tr>"
            
                # Copies Actually Assigned
	
	            $Output2 += "<tr bgcolor=""#440164""><th><font color=""#ffffff"">Actual Mounted DB Copies</font></th>"
	
	            $Servers_In_DAG|Sort-Object| 
		                %{ 
                        $srv = $ExchangeEnvironment.Servers[$_]
					    $Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($srv.DBMounted)</strong></font></td>"	
						
		                }
	            $Output2 +="</tr></table>"            	

            


            #endregion Table Part 2 : Aggregates info

            Write-Output $Output2


        } # function _GetDAG_DB_Layout
   

        function _GetDBPreference_Table_HTML {

            param ($DAG ,$ExchangeEnvironment,$hash_server_vs_db_Aggregates)
            #input is DAG Custom Object        

            # Database Availability Table number 1 [Header]
		    $Output ="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
		    <col width=""20%""><col width=""10%""><col width=""70%"">
		    <tr align=""center"" bgcolor=""#FC8E10""><th><font color=""#ffffff"">Database Availability Group Name</font></th><th><font color=""#ffffff"">Member Count</font></th>
		    <th><font color=""#ffffff"">Database Availability Group Members</font></th></tr>
		    <tr><td>$($DAG.Name)</td><td align=""center"">
		    $($DAG.MemberCount)</td><td>"
		    $DAG.Members | % { $Output+="$($_) " }
		    $Output +="</td></tr></table>"

        
            #Database Availability Table number 2 
            $Output += _GetDAG_DB_Layout -Databases $DAG.Databases -DAG $DAG $ExchangeEnvironment $hash_server_vs_db_Aggregates
              

            Write-Output $output

      } # function _GetDBPreference_Table_HTML


   #endregion GUI functions


    #region chart functions        
        

       function Get-Corpchart_light {

            <#

            .Synopsis
            Draws a chart. Requires both .NET 3.5 and Microsoft Chart Controls for Microsoft .NET Framework 3.5 (http://www.microsoft.com/en-us/download/details.aspx?id=14422)
            Note: To get the more advance PowerShell Chart Script Wrapper, visit http://wp.me/p1eUZH-l5

            .DESCRIPTION
            Draws a chart. Requires both .NET 3.5 and Microsoft Chart Controls for Microsoft .NET Framework 3.5 (http://www.microsoft.com/en-us/download/details.aspx?id=14422).

            Input data:
             array of objects.You have to supply which are the two properties of the object to draw by supplying parameters (-obj_key and -obj-value)

   


            .PARAMETER Data
            Array of objects. Required parameter.

            .PARAMETER Filepath
            File path to save the chart like "c:\Chart1.PNG". Required parameter.

            .PARAMETER Type 
            Chart type. Default is 'column'. Famous types are "Point", "FastPoint", "Bubble", "Line","Spline", "StepLine", "FastLine", "Bar","StackedBar", "StackedBar100", "Column",
                                 "StackedColumn", "StackedColumn100", "Area","SplineArea","StackedArea", "StackedArea100", "Pie", "Doughnut", "Stock", "Candlestick",
                                 "Range","SplineRange", "RangeBar", "RangeColumn", "Radar", "Polar", "ErrorBar", "BoxPlot", "Renko", "ThreeLineBreak", "Kagi", "PointAndFigure", "Funnel",
                                 "Pyramid"

            .PARAMETER Title_text
            Chart title. Default is empty title. Optional parameter.

            .PARAMETER Chartarea_Xtitle.  
            Chart X Axis title. Default is empty title. Optional parameter.

            .PARAMETER Chartarea_Ytitle  
            Chart Y Axis title. Default is empty title. Optional parameter.

            .PARAMETER Xaxis_Interval  
            Enter X Axis interval. Default is 1. Optional parameter.

            .PARAMETER Yaxis_Interval  
            Enter Y Axis interval. Usually you do not need to use this parameter. Optional parameter.

            .PARAMETER Chart_color  
            Enter chart column color. Only in case of 'column' or 'bar' chart types. Optional parameter.

            .PARAMETER Title_color  
            Enter chart title color. Default is 'red'. Optional parameter.

            .PARAMETER CollectedThreshold  
            Enter a threshold that all data values below it will be grouped as one item named 'Others'. This parameter takes an integer from 1 to 100. Optional parameter.

            .PARAMETER Sort  
            Sort the data option. Values are either 'asc' or 'dsc'. Optional parameter.

            .PARAMETER IsvalueShownAsLabel
            Switch parameter to indicate values appear as a label on the chart. Optional parameter.

            .PARAMETER ShowHighLow
            Switch parameter to indicate that the maximum and minimum values are highlighted in the chart.Optional parameter.

            .PARAMETER ShowLegend
            Switch parameter to determine if legend should be added to the chart. Optional parameter.

            .PARAMETER Obj_key
            This parameter is required when the input data type is array of objects. This represents the name of the object properties to be used as X Axis data. Optional parameter.

            .PARAMETER Obj_value
            This parameter is required when the input data type is array of objects. This represents the name of the object properties to be used as Y Axis data. Optional parameter.

            .PARAMETER Append_date_title
            Append the current date to the title. Optional parameter.

            .PARAMETER Fix_label_alignment
            Only applicable if the chart type is Pie or Doughnut. If the data labels in the chart are overlapping, use the switch to fix it. Optional parameter.

            .PARAMETER Show_percentage_pie
            Only applicable if the chart type is Pie or Doughnut. This will show the data labels on the chart as percentages instead of actual data values. Optional parameter.



            .EXAMPLE
            Chart by supplying array of objects as input. We are interested in the Name and Population properties of the input objects.
            In this case, we should also use the -obj_key and -obj_value parameters to tell the function which properties to draw. Default chart type 'column' is used.
            PS C:\> Get-Corpchart_light -data $array_of_city_objects -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" 

            .EXAMPLE
            Specifying chart type as pie chart type.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type pie

            .EXAMPLE
            Specifying chart type as pie chart type. legend is shown.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type pie -showlegend 

            .EXAMPLE
            Specifying chart type as SplineArea chart type.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type SplineArea

            .EXAMPLE
            Specifying chart type as bar chart type, and specifying the title for the chart and x/y axis.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type Bar -title_text "people per country" -chartarea_Xtitle "cities" -chartarea_Ytitle "population"

            .EXAMPLE
            Specifying chart type as column chart type. Applying the -showHighLow switch to highlight the max and min values with different colors.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type column -showHighLow


            .EXAMPLE
            Chart with percentages shown on the pie/doughnut charts.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type Doughnut -Show_percentage_pie

            .EXAMPLE
            Chart with percentages shown on the pie/doughnut charts. Fixing the overlapping labels on the chart.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -charttheme 5 -Show_percentage_pie -fix_label_alignment

            .EXAMPLE
            If the chart type is pie or doughnut, you can specify a threshold (percentage) that all data values below it, will be shown as one data item called (Others).
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -type Doughnut -CollectedThreshold 16

            .EXAMPLE
            Column chart with green columns.
            PS C:\> Get-Corpchart_light -data $cities -obj_key "Name" -obj_value "Population" -filepath "c:\chart.png" -chart_color green


            .Notes
            Last Updated             : April 23, 2014
            Version                  : 1.0 
            Author                   : Ammar Hasayen (Twitter @ammarhasayen)
            Email                    : me@ammarhasayen.com


            Note: To get the more advance PowerShell Chart Script Wrapper, visit http://wp.me/p1eUZH-l5


            .Link
            http://ammarhasayen.com




            #>

                [cmdletbinding()]  


                Param(

                    #region REQUIRED parameters        
        
                        [ValidateNotNull()]
                        [ValidateNotNullorEmpty()]
                        [object]$data,          
        
                        [string]$filepath,
                        [string]$ErrorFullPath,
        
                    #endregion



                    #region OPTIONAL parameters
        
                        # Chart type               
                        [string]$Type = "column",        

        
                        # Chart Titles
                        [string]$title_text = " ",

                        [string]$chartarea_Xtitle = " ",

                        [string]$chartarea_Ytitle = " ",

                        [int]$Xaxis_Interval = 1,

                        [int]$Yaxis_Interval,

                        [string]$chart_color = "MediumSlateBlue",

                        [string]$title_color="red",   

                        # Chart extra customization               
                        [string]$sort, 

                        [switch]$IsvalueShownAsLabel,
    
                        [switch]$showHighLow,
        
                        [switch]$showLegend,

                        [string]$obj_key,

                        [string]$obj_value,
        
                        [switch]$append_date_title,

                        [switch]$fix_label_alignment,
       
                        [switch]$show_percentage_pie,

                        [int]$CollectedThreshold
       
                   #endregion
    
                )


                Begin {  

                        #region variables

                            New-Variable -Name currentDate -Option ReadOnly -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -scope local 

                            New-Variable -Name font_Style1 -Option ReadOnly -Value (new-object system.drawing.font("ARIAL",18,[system.drawing.fontstyle]::bold))  -scope private 

                            New-Variable -Name font_style2 -Option ReadOnly -Value (new-object system.drawing.font("calibri",16,[system.drawing.fontstyle]::italic))  -scope private 

                            #chart background color
                            New-Variable -Name chartarea_backgroundcolor -Option ReadOnly  -Value "white" -scope private

           

            
                            #default chart dimension
                            $propChartDimension  = @{ "width"           = 1500;
                                                      "height"          = 800;
                                                      "left"            = 80;
                                                      "top"             = 100;
                                                      "name"            = "chart";
                                                      "BackColor"       = "white"
                                          } 

                            $ObjChartDimension = New-Object -TypeName psobject -Property $propChartDimension



                            #hashtable for theme IDs
                            $theme_charttype = @{ 1   = "column"  ;
                                                  2   = "column"  ;
                                                  3   = "bar"     ;
                                                  4   = "Doughnut";
                                                  5   = "Doughnut";
                                                  6   = "pie"     ;
                                                  7   = "Doughnut";
                                                }

                            #hashtable to mark dimensions that will be scaled dynamically according to the number of input objects
                            $dynamicdimension = @{"column"="width";
                                                  "bar"   ="height"
                                                 }
                        #endregion variables



                        #region get class

                            # loading the data visualization .net class
                            [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

                        #endregion get class



                        #region internal functions

                            Function Get-CorpCalculateDim {

                                # Adjust width (in case of charts with 'column' type
                                # or height (in case of charts with 'bar' type according to the number of items
                                # Criteria :For each 15 item, dimention should expand by 1000

                                param ($count)

                                    [int]$v = ($count / 15)

                                    # if input items are less than 15 items (which gives 0 when doing int division by 15)
                                    if($v -eq 0) {$v=1}
                                    # if input items are more than 15 items
                                    else { $v=$v+1}

                                    return ($v*1000)

                            } #  Function Get-CorpCalculateDim


                        #endregion internal functions
            
            
                } # Function Get-Corpchart_light BEGIN Section


                Process {             
            
                        # Function Get-Corpchart_light PROCESS Section

                        #region input validation
            
                            # you should use the -obj_key and -obj_value if the input data is array of objects
                            if (($data.gettype()).name -notlike "hashtable") {

                                  if ( -not (   ($PSBoundParameters.ContainsKey("obj_key") ) -and (($PSBoundParameters.ContainsKey("obj_value")) ))) {

                                     throw " Since your data is not a hashtable, then it shall be array of objects. In this case, use -obj_key and -obj_value parameter to inform the function about which properties to draw."

                                  } # end if


                            } # end if

            
                            # do not use the -obj_key and -obj_value parameters if the input data is hashtable
                            if (($data.gettype()).name -like "hashtable") {

                              if (  (   ($PSBoundParameters.ContainsKey("obj_key") ) -or (($PSBoundParameters.ContainsKey("obj_value")) ))) {

                                 throw "Input data is hashtable. No need to specify -obj_key or -obj_value parameter"

                              } # end if


                        } # end if   



                        #endregion            


                        #region create chart  
                            $chart  = new-object System.Windows.Forms.DataVisualization.Charting.Chart
                        #endregion

            
                        #region chart data

                            [void]$chart.Series.Add("Data") 

            
                            $array_keys   = @()
                            $array_values = @()

                            foreach ($object in $data) {

                                $array_keys += $object.$obj_key
                                $array_values += $object.$obj_value   

                            } # end foreach
            
                            $chart.Series["Data"].Points.DataBindXY($array_keys, $array_values)
            
                        #endregion 


                        #region chart type
                            $chart_type   =  $Type
                            $chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$chart_type
            
                        #endregion


                        #region chart look and size : setting default chart dimensions 
            
                                $chart.width     = $ObjChartDimension.width
                                $chart.Height    = $ObjChartDimension.height
                                $chart.Left      = $ObjChartDimension.left
                                $chart.top       = $ObjChartDimension.top
                                $chart.Name      = $ObjChartDimension.name
                                $chart.BackColor = $ObjChartDimension.BackColor

                            # if the chart type is one that needs dynanmic dimensions according to the $dynamicdimension hashtable
                            # then we need to pull the dimension to be dynamically calculated
                            # example, if you are going to draw column chart type, then we will expand the width of the chart
                            # according to the number of items in the input data to give some room.
                            # while if you are going to draw a bar chart type, then we will expand the height accordingly.
                            if ($dynamicdimension.ContainsKey($chart_type)) {

                                # the variable item represents which dimension (according to the chart type) to  be dynamically calculated.
                                # for example, in case of column chart type, #item will be the width dimension.
                                $item = $dynamicdimension[$chart_type]
                                # the function Get-CorpCalculateDim will return the value of that item by giving it the number of items in the data input.
                                $chart.$item =  Get-CorpCalculateDim ($data.count)

                            } # if ($dynamicdimension.ContainsKey($chart_type))             

                        #endregion   


                        #region chart label

                        # if the $IsvalueShownAsLabel switch is specified, we will enable the label on the chart
                        $chart.Series["Data"].IsvalueShownAsLabel = $IsvalueShownAsLabel

                        #endregion


                        #region chart maxmin

                            # there is an option where the highest and lowest values in the input data can be highlighted by different colors
                            # if you specify the -showHighLow switch, the chart will do the highlighting

                            if ( $PSBoundParameters.ContainsKey("showHighLow") ) {

                                 #Find point with max value and change the colour of that value to red
                                 $maxValuePoint = $Chart.Series["Data"].Points.FindMaxByValue() 
                                 $maxValuePoint.Color = [System.Drawing.Color]::Red 
 
                                 #Find point with min value and change the colour of that value to green
                                 $minValuePoint = $Chart.Series["Data"].Points.FindMinByValue() 
                                 $minValuePoint.Color = [System.Drawing.Color]::Green
                            }

                        #endregion


                        #region Title

                            # putting the title of the chart

                            $title =New-Object System.Windows.Forms.DataVisualization.Charting.title
                            $chart.titles.add($title)            
                            $chart.titles[0].font         = $font_style1
                            $chart.titles[0].forecolor    = $title_color
                            $chart.Titles[0].Alignment    = "topLeft"

                            if ($PSBoundParameters.ContainsKey("append_date_title") ) {

                                $chart.titles[0].text   = ($title_text + "`n " + $currentDate )


                            }
                            else {
                
                                $chart.titles[0].text   = $title_text

                            }

                        #endregion


                        #region legend

                            # putting the legend of the chart if the -showlegend switch is used

                            if ( $PSBoundParameters.ContainsKey("showlegend") ) {

                                $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
                                $legend.BorderColor     = "Black"
                                $legend.Docking         = "Top"
                                $legend.Alignment       = "Center"
                                $legend.LegendStyle     = "Row"
                                $legend.MaximumAutoSize  = 100
                                $legend.BackColor       = [System.Drawing.Color]::Transparent  
                                $legend.shadowoffset= 1  
                    
                                $chart.Legends.Add($legend)

                            } # if ( $PSBoundParameters.ContainsKey("showlegend") )


                        #endregion


                        #region chart area

                            # chart area is where the X axis and Y axis titles and font style is to be configured.

                            $chartarea                = new-object system.windows.forms.datavisualization.charting.chartarea
                            $chartarea.backcolor      = $chartarea_backgroundcolor
                            $ChartArea.AxisX.Title    = $chartarea_Xtitle
                            $ChartArea.AxisX.TitleFont= $font_style2
                            $chartArea.AxisY.Title    = $chartarea_Ytitle             
                            $ChartArea.AxisY.TitleFont= $font_style2
                            $ChartArea.AxisX.Interval = $XAxis_Interval 
            
                            if ($PSBoundParameters.ContainsKey("YAxis_Interval") ) {
                                $ChartArea.AxisY.Interval = $YAxis_Interval           
                            }

                            $chart.ChartAreas.Add($chartarea)

                        #endregion                   

            
                        #region more configurations. 
            
                                #region Pie and Doughnuts configuration

                                    if (($chart_type -like "pie") -or ($chart_type -like "Doughnut") ) {
                                    # this applies only if the chart type is Pie or Doughnut.
                
                                        #region CollectedThreshold settings
                
                                        # sometimes, there is so much data to draw, you can specify a threshold (value between 1 and 100) 
                                        # which represents a percentage of the input data value, that is when any data item value is below it
                                        # the chart will group them under (Other) as one item with green color.
                                        if ( $PSBoundParameters.ContainsKey("CollectedThreshold") ) {

                                               $chart.Series["Data"]["CollectedThreshold"]           = $CollectedThreshold   
                                               $chart.Series["Data"]["CollectedLabel"]               = "Other"
                                               $chart.Series["Data"]["CollectedThresholdUsePercent"] = $true
                                               $chart.Series["Data"]["CollectedLegendText"]          = "Other"
                                               $chart.Series["Data"]["CollectedColor"]               = "green"

                                        } # if ( $PSBoundParameters.ContainsKey("CollectedThreshold") )

                                        #endregion

                                        #region fix alignment for labels

                                        # sometime the labels on the chart overlap above each other's making ugly look
                                        # the trick is make that chart as 3D with zero inclination
                                        # this can be done if you specify the -fix_label_alignment switch
                                        if ( $PSBoundParameters.ContainsKey("fix_label_alignment") ) {

                                            $chartArea.Area3DStyle.Enable3D = $true

                                            # if there is no inclination configured, that is if the chart is configured as 3D.
                                            # this validation is important to prevent overwriting an already configured inclination for 3D charts.
                                            if(-Not ($chartArea.Area3DStyle.Inclination)) { $chartArea.Area3DStyle.Inclination = 0 }
                

                                        } # if ( $PSBoundParameters.ContainsKey("fix_label_alignment") )
                
                                 #endregion

                            #region show data as percentage

                            # sometimes, it is better to show the data values as percentages instead of actual values.
                            # this can be done by using the -show_percentage_pie.
                            # this applies to both pie and doughnut chart types.
                            if ( $PSBoundParameters.ContainsKey("show_percentage_pie") ) {
                                   
                                #we will set the label to VLAX which is the X axis value then the percent with two decimals of the value (Y axis)
                                $chart.Series["Data"].Label = "#VALX (#PERCENT{P2})"

                                # on the legend, we will put the X axis value (VLAX).
                                $chart.Series["Data"].LegendText = "#VALX"                    

                            } # if ( $PSBoundParameters.ContainsKey("show_percentage_pie") )


                            #endregion

                        } # if (($chart_type -like "pie") -or ($chart_type -like "Doughnut") )

                        #endregion

                                #region Column and Bar configuration

                                    if (($chart_type -like "column") -or ($chart_type -like "bar") ) {
                                    # this applies only if the chart type is column or bar.
            
                                         # the X axis and Y axis line colors are set to DarkBlue.
                                         $chartarea.AxisX.LineColor =[System.Drawing.Color]::DarkBlue 
                                         $chartarea.AxisY.LineColor =[System.Drawing.Color]::DarkBlue                 

                                         # the title of the X axis and Y axis font color is set to DarkBlue.
                                         $ChartArea.AxisX.TitleForeColor =[System.Drawing.Color]::DarkBlue
                                         $ChartArea.AxisY.TitleForeColor =[System.Drawing.Color]::DarkBlue 
                 
                                         # configuring the internal chart grid.

                                         # enable customization of the grid
                                         $chartarea.AxisX.IsInterlaced = $true 
                                         # the grid line color
                                         $chartarea.AxisX.InterlacedColor = [System.Drawing.Color]::AliceBlue 
                                         # grid line type
                                         $chartarea.AxisX.ScaleBreakStyle.BreakLineStyle = "Straight"
                                         # grid area alternate color for both axises
                                         $chartarea.AxisX.MajorGrid.LineColor =[System.Drawing.Color]::LightSteelBlue 
                                         $chartarea.AxisY.MajorGrid.LineColor =[System.Drawing.Color]::LightSteelBlue 

                                         # configuring the chart column internal color
                                         $chart.Series["Data"].Color = $chart_color


                                    } # if (($chart_type -like "column") -or ($chart_type -like "bar") )

                                #endregion            

                                #region Configuration that applies to all chart types

                                         $chart.BorderlineWidth = 1
                                         $chart.BorderColor = [System.Drawing.Color]::black
                                         $chart.BorderDashStyle = "Solid" # values can be "Dash","DashDot","DashDotDot","Dot","NotSet","Solid"
                                         $chart.BorderSkin.SkinStyle = "Emboss"
                                #endregion

                                #region data sorting

                                 # showing the data sorted is a welcome thing. If you specify the -sort parameter, we will sort the data before drawing it.
             
                                 if ($sort -like "asc") {
                                    $Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Ascending, "Y") 
                                 }
              
                                 elseif ($sort -like "dsc") {

                                    $Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Descending, "Y") 
                                 }

                                 #endregion

                        #endregion
             


                } # Function Get-Corpchart_light PROCESS Section


                End {
                         $var = $ErrorActionPreference
                         $ErrorActionPreference = "Stop"
                         try{
                            $chart.SaveImage($filepath, "PNG")
                        }catch {
                            Write-CorpError -myError $_ -mypath $ErrorFullPath -Info "[Module 6 Charts : Failed to save chart at $filepath"
                        }
                        finally {
                        $ErrorActionPreference = $var
                        }


                } # Function Get-Corpchart_light END Section


        } # Function Get-Corpchart_light


    #endregion chart functions

#endregion Module 2 : Functions

#endregion Module 2 : Functions

#++++++++++++++++++++++++++++++++++++     Module 3 : Factory    ++++++++++++++++++++++++++++++++++++



#region Module 3 : Factory

    # This module is all about preparing the script environment by creating couple of log files and 
    # validating the PowerShell environment beside creating some global variables.

     
    #region create directory
    
    try{
        $ScriptFilesPath = Convert-Path $ScriptFilesPath -ErrorAction Stop
    }catch {
        Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Validating log files path] Sorry, please check the sript path name again"
        Exit
        throw " [Module Factory - Creating files] Validating log files path] Sorry, please check the sript path name again "
    }

    $ScriptFilesPath = Join-Path $ScriptFilesPath "EmailReportFiles"

    if(Test-Path $ScriptFilesPath ) {
        try {
                Remove-Item $ScriptFilesPath -Force -Recurse -ErrorAction Stop

            }catch {

                Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Deleting old working directory] Could not delete directory $ScriptFilesPath"
                Exit
                throw "[Module Factory - Deleting old working directory] Could not delete directory $ScriptFilesPath "

            }
    }
    
    
    if(!(Test-Path $ScriptFilesPath )) {
        try {
            New-Item -ItemType directory -Path $ScriptFilesPath -ErrorAction Stop
        }catch{
            Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating working directory] Could not delete directory $ScriptFilesPath"
            Exit
            throw "[Module Factory - Creating working directory] Could not delete directory $ScriptFilesPath "

        }
    }  
    
    
    #endregion create directory
                
    #region 3.1 : Create files
            
        _status "    1.1 Creating Files" 1
            
        #region ErrorLog            
           
            try{                
                $ErrorLogFile = "ExReportError.log"    
                $ErrorFullPath = Join-Path $ScriptFilesPath  $ErrorLogFile -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Validating log files] Sorry, please check the sript path name again"
                Exit
                throw " [Module Factory - Creating files] Sorry, please check the sript path name again "

            }
             
            #Check if file exists and delete if it does

            If((Test-Path -Path $ErrorFullPath )){
                try {
                Remove-Item -Path $ErrorFullPath  -Force -ErrorAction Stop
                }catch {
                    Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Deleting old log files]"
                    Exit
                    throw " [Module Factory - Creating files] Sorry, but the script could not delete log file on $ErrorFullPath "
                }
            }
    
            #Create Error file

            try {
                New-Item -Path $ScriptFilesPath -Name $ErrorLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating log files]"
                 Exit
                throw "  [Module Factory - Creating files] Sorry, but the script could not create log file on $ScriptFilesPath [Module 3 Factory - Creating files] "
            }

            #initiate error log file

            Log-Start -LogFullPath $ErrorFullPath

            Write-verbose -Message "[Module Factory] : Error Log File created $ErrorFullPath"

        #endregion ErrorLog

        #region InfoLog

            try{
                $InfoLogFile = "ExReportInfo.log"    
                $InfoFullPath =  Join-Path $ScriptFilesPath  $InfoLogFile -ErrorAction Stop
            }catch{
                Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Validating log files] Sorry, please check the sript path name again"
                Exit
                throw " [Module Factory - Creating files] Sorry, please check the sript path name again "

            }

            #Check if file exists and delete if it does
            If((Test-Path -Path $InfoFullPath )){
                try {
                    Remove-Item -Path $InfoFullPath  -Force -ErrorAction Stop
                }catch {
                    Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Deleting old log files]"
                    Exit
                    throw "[Module Factory] : Sorry, but the script could not delete log file on $InfoFullPath [Module 3 Factory - Creating files] "
                }
            }
    
            #Create Error file
            try {
                New-Item -Path $ScriptFilesPath -Name $InfoLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating log files]"
                 Exit                 
                 throw "[Module Factory] : Sorry, but the script could not delete log file on $ScriptFilesPath [Module 3 Factory - Creating files] "
            }

           
            #initiate error log file

             Log-Start -LogFullPath $InfoFullPath

             Write-verbose -Message "[Module Factory] : Error Log File created $InfoLogFile"
            

            

        #endregion Info Log

        #region DetailedLog            
           
            $DetailedLogFile = "ExReportDetailed.log"    
            try {
                $DetailedFullPath =  Join-Path $ScriptFilesPath  $DetailedLogFile -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Validating log files] Sorry, please check the sript path name again"
                Exit

            }
             
            #Check if file exists and delete if it does

            If((Test-Path -Path $DetailedFullPath )){
                try {
                Remove-Item -Path $DetailedFullPath  -Force -ErrorAction Stop
                }catch {
                     Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Deleting old log files]"
                     Exit
                     throw " [Module Factory - Creating files] Sorry, but the script could not delete log file on $DetailedFullPath "
                }
            }
    
            #Create Error file

            try {
                New-Item -Path $ScriptFilesPath -Name $DetailedLogFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating log files]"
                 Exit                
                throw "  [Module Factory - Creating files] Sorry, but the script could not create log file on $ScriptFilesPath [Module 3 Factory - Creating files] "
            }

            #initiate error log file

            Log-Start -LogFullPath $DetailedFullPath

            Write-verbose -Message "[Module Factory] : Error Log File created $DetailedFullPath"

        #endregion DetailedLog
    

    #endregion 3.1 : Create files

    #region initial log info    
        
            
            
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Info log file $($InfoFullPath) is created successfully"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Error log file $($ErrorFullPath) is created successfully"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Detailed log file $($DetailedFullPath) is created successfully"


        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Starting $($MyInvocation.Mycommand)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): ($PSBoundParameters | out-string)"
        
        if($PSVersionTable.PSVersion.major){
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): PowerShell Host Version :$($PSVersionTable.PSVersion.major)"
        }

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 1 Customization] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 3 Function] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  
 
    #endregion initial log info

    #region Section 3.2 : Initial preperation

        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 1 of 6 : Preperation Tasks" -percentComplete 10
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module Factory]"

        #region 3.1.1 Start StopWatch

            #Start stop watch
            # This is to report the time it takes to run the script.
            $Watch  =  [System.Diagnostics.Stopwatch]::StartNew()
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Starting script stopwatch"

        #endregion 3.1.2 Start StopWatch


        #region 3.1.3 : Screen Headings

            Write-Verbose -Message "Info : Starting $($MyInvocation.Mycommand)"  

            Write-verbose -Message ($PSBoundParameters | out-string)

            _screenheadings

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Clearing the console screen and displaying script headings"

            _status " 1. Preparation Tasks" 1 

        #endregion  3.1.3 : Screen Heading        


        #region 3.1.4 : PowerShell Environment 
            
            
            _status "    1.1 Checking PowerShell environment" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Checking Script PowerShell Environment"

            # Quote : Script block taken from Steve's script

                    # Check Powershell Version

                    if ((Get-Host).Version.Major -eq 1)
                    {
	                    Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [Module Factory] Powershell Version 1 not supported : if ((Get-Host).Version.Major -eq 1)"
                        
                        throw "Powershell Version 1 not supported";
                    }

                    # 1.1 Check Exchange Management Shell, attempt to load
                    
                        # Sometime it is tricky to load Exchange Management Shell specially if Exchange was installed on a drive other than the C drive.
                        #So we will get the Exchange Installation Path
                        [string]$Exch_InstallPath = $env:exchangeinstallpath
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):    - Exchange Install Directory : $Exch_InstallPath "
                        $Exch_InstallDrive = $Exch_InstallPath.Substring(0,3)                         
                        $loadscript1 = Join-Path $Exch_InstallDrive "Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"  
                        $loadscript2 = Join-Path $Exch_InstallDrive "Program Files\Microsoft\Exchange Server\bin\Exchange.ps1"   
                        $loadscript3 = Join-Path $Exch_InstallPath  "bin\RemoteExchange.ps1"


                        if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
                           
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):    - Exchange Commands not available.. trying to load Exchange PowerShell"

                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"
	                        if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"){	  
                                
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"                              
                                . 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
		                        Connect-ExchangeServer -auto
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Works!"

	                        }  
                            
                            elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1") {
                               Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1"
		                        Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		                        .'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Works!"
	                        }
                            
                            elseif (Test-Path $loadscript1 ) {	
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load $loadscript1"                        
		                        . $loadscript1
		                        Connect-ExchangeServer -auto
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Works!"
	                        }
                            
                            elseif (Test-Path $loadscript2) {
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load $loadscript2"
		                        Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		                        . $loadscript2
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Works!"
	                        }
                            
                            elseif (Test-Path $loadscript3 ) { #Exchange 2013	                        
		                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Trying to load $loadscript3"
                                . $loadscript3
		                        Connect-ExchangeServer -auto
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):       -- Works!"
	                        }    

                            else {
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [Module Factory] Exchange Management Shell cannot be loaded"
		                    
                                throw "Exchange Management Shell cannot be loaded"                            
	                        }

                        }


                     # Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
                        if ($SendMail)
                        {
	                        if (!$MailFrom -or !$MailTo -or !$MailServer)
	                        {
		                        Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [Module Factory] If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer"

                                throw "If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer"
	                        }
                        }

                      # Check Exchange Management Shell Version
                        if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
                        {
	                        $E2010 = $false;
	                        if (Get-ExchangeServer | Where {$_.AdminDisplayVersion.Major -gt 14})
	                        {
		                        Write-Warning "Exchange 2010 or higher detected. You'll get better results if you run this script from an Exchange 2010/2013 management shell"
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module Factory] Exchange 2010 or higher detected. You'll get better results if you run this script from an Exchange 2010/2013 management shell"

	                        }
                        }else{
    
                            $E2010 = $true

                            $varPS = $ErrorActionPreference
                            $ErrorActionPreference = "Stop"
                            try {
                                $localserver = get-exchangeserver $Env:computername -ErrorAction Stop
                                $localversion = $localserver.admindisplayversion.major
                                if ($localversion -eq 15) { $E2013 = $true }
                            }catch {
                            Write-Warning -Message " [Module Factory] You are not running the script from an Exchange Server"
                            Write-Warning -Message " [Module Factory] The script logic cannot determine if PS version is E2013 or E2010"
                            Write-Warning -Message " [Module Factory] Knowing this info is so important to determine the command set to use"
                            Write-Warning -Message " [Module Factory] Command failing is (Get-ExchangeServer `$Env:computername) "
                            Write-Warning -Message " [Module Factory] `$Env:computername in this case evaluates to $($Env:computername)"
                            Write-Warning -Message " [Module Factory] Please run the script from an Exchange Server"
                            Write-Warning -Message " [Module Factory] Existing script"
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): You are not running the script from an Exchange Server"
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): The script logic cannot determine if PS version is E2013 or E2010"
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Knowing this info is so important to determine the command set to use"
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Command failing is (Get-ExchangeServer `$Env:computername) "
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): `$Env:computername in this case evaluates to $($Env:computername)"
                            Log-Write -LogFullPath $ErrorFullPath-LineValue "$(get-timestamp): Please run the script from Exchange Server"
                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Exiting script"
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Exiting script... Please check ErrorLog file for detail."
                            Write-Host -ForegroundColor Red "Terminating Script.. Pls check Error Log file for details"
                            Write-Host -ForegroundColor Red "Error Log File is : $ErrorFullPath "
                            Exit
                            Throw "[Module Factory] You are not running the script from an Exchange Server so the code decided to Exit"
                            }finally {

                            $ErrorActionPreference = $varPS
                            }

                            if(!$localversion) {
                                Write-Warning -Message " [Module Factory] You are not running the script from an Exchange Server"
                                Write-Warning -Message " [Module Factory] The script logic cannot determine if PS version is E2013 or E2010"
                                Write-Warning -Message " [Module Factory] Knowing this info is so important to determine the command set to use"
                                Write-Warning -Message " [Module Factory] Command failing is (Get-ExchangeServer `$Env:computername) "
                                Write-Warning -Message " [Module Factory] `$Env:computername in this case evaluates to $($Env:computername)"
                                Write-Warning -Message " [Module Factory] Please run the script from an Exchange Server"
                                Write-Warning -Message " [Module Factory] Existing script"
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): You are not running the script from an Exchange Server"
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): The script logic cannot determine if PS version is E2013 or E2010"
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Knowing this info is so important to determine the command set to use"
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Command failing is (Get-ExchangeServer `$Env:computername) "
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): `$Env:computername in this case evaluates to $($Env:computername) "
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Please run the script from Exchange Server"
                                Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): Exiting"
                                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Exiting script... Please check ErrorLog file for detail."
                                Write-Host -ForegroundColor Red "Terminating Script.. Pls check Error Log file for details"
                                Write-Host -ForegroundColor Red "Error Log File is : $ErrorFullPath "
                                Exit
                                Throw "[Module Factory] You are not running the script from an Exchange Server so the code decided to Exit"




                            }

                        }

                      #  Check view entire forest if set (by default, true)
                        if ($E2010)
                        {
	                        Set-ADServerSettings -ViewEntireForest:$ViewEntireForest 
                        } else {
	                        $global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module Factory] ViewEntireForest value is $ViewEntireForest"

                        }

            #End Quote : Script block taken from Steve's script


                        #region log info

                        Write-Verbose -Message "Info: Value of E2010 Variable is $E2010"
                        Write-Verbose -Message "Info: Value of E2013 Variable is $(if($E2013){$true}else{$false})"

                        
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Info: Value of E2010 Variable is $E2010"
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Info: Value of E2013 Variable is $(if($E2013){$true}else{$false})"
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Finish validating PowerShell environment"

                        #endregion log info


                        #region final checks

                        if (!$E2010 -and ($ScriptFilter -like "DAGFilter") ) {

                            Log-Write -LogFullPath $ErrorFullPath -LineValue "$(get-timestamp): [Module Factory] Error: You cannot use DAGFilter while running old version of PowerShell"
                            throw "[Module Factory] Error: You cannot use DAGFilter while running old version of PowerShell"

                        }

                        #endregion final checks




        #endregion 3.1.4 : PowerShell Environment

       
    #endregion Section 3.1 : Initial preperation

    #region Section 3.2 : Global variables
        
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Creating Global Variables"
        
        _status "    1.2 Creating global variables" 2
       
        #region Script data structure Hash Table

            #Quote : Script block taken from Steve's script
            
                $ExchangeEnvironment = @{Sites					= @{}
						                 Pre2007				= @{}
						                 Servers				= @{}
						                 DAGs					= @()
						                 NonDAGDatabases		= @()
                                         RecoveryDatabases		= @()
                                         NonRecoveryDatabases   = @()
                                         Chart_db               = @()
                                         Chart_srv              = @()
						                }
              # hash table to represent an object that will use for charts

              $chart_db_objprop = @{ Name = ""
                                     MailboxCount = ""
                                     DBSize  = ""  
                                     Backup = "" 
                                  }
              $chart_srv_objprop = @{ Name = ""
                                      DBMountedCount = ""
                                      MailboxCount = ""   
                                  }


                 #Exchange Major Version String Mapping
                $ExMajorVersionStrings = @{"6.0" = @{Long="Exchange 2000";Short="E2000"}
				   		                    "6.5" = @{Long="Exchange 2003";Short="E2003"}
				   		                    "8"   = @{Long="Exchange 2007";Short="E2007"}
                                            "14"  = @{Long="Exchange 2010";Short="E2010"}
						                    "15"  = @{Long="Exchange 2013";Short="E2013"}
                                            "16"  = @{Long="Exchange 2013";Short="E2013"}}

                #Exchange Service Pack String Mapping
                $ExSPLevelStrings = @{"0" = "RTM"
					                    "1" = "SP1"
				                        "2" = "SP2"
			                            "3" = "SP3"
				                        "4" = "SP4"
                                        "CU1" = "CU1"
                                        "CU2" = "CU2"
                                        "CU3" = "CU3"
                                        "CU4" = "CU4"
                                        "CU5" = "CU5"
                                        "CU6" = "CU6"
                                        "CU7" = "CU7"
                                        "CU8" = "CU8"
                                        "CU9" = "CU9"
                                        "SP1" = "SP1"
                                        "SP2" = "SP2"}

                 #Exchange Role String Mapping
                $ExRoleStrings = @{"ClusteredMailbox" = @{Short="ClusMBX";Long="CCR/SCC Clustered Mailbox"}
				                    "Mailbox"		  = @{Short="MBX";Long="Mailbox"}
				                    "ClientAccess"	  = @{Short="CAS";Long="Client Access"}
				                    "HubTransport"	  = @{Short="HUB";Long="Hub Transport"}
				                    "UnifiedMessaging" = @{Short="UM";Long="Unified Messaging"}
				                    "Edge"			  = @{Short="EDGE";Long="Edge Transport"}
				                    "FE"			  = @{Short="FE";Long="Front End"}
				                    "BE"			  = @{Short="BE";Long="Back End"}
				                    "Unknown"	  = @{Short="Unknown";Long="Unknown"}}

                #Populate Full Mapping using above info
                $ExVersionStrings = @{}
                foreach ($Major in $ExMajorVersionStrings.GetEnumerator()) {
        
	                foreach ($Minor in $ExSPLevelStrings.GetEnumerator()) {
	        
		                $ExVersionStrings.Add("$($Major.Key).$($Minor.Key)",@{Long="$($Major.Value.Long) $($Minor.Value)";Short="$($Major.Value.Short)$($Minor.Value)"})
	                }
                }
    
             #End Quote : Script block taken from Steve's script

             #endregion Script data structure Hash Table

        #region other variables

            #holds results from Get-ExchangeServer
            $ExchangeServers = @()
            #string array holding the name of the Exchange Servers after filter
            $ExchangeServersList = @()
            #holds result of get-mailboxdatabase after filter
            $Databases = @() 
            #string array holding the name of the databases after filter
            $DatabasesList = @()           
            #holds Get-RemoteMailbox
            $RemoteMailboxes 
            #holds get-mailboxes after filter
            $Mailboxes = @()
            #holds get-mailbox -archive after filter
            $ArchiveMailboxes = @()
            #holds get-databaseavailabilitygroup after filter
            $DAGs = @()
            #holds archive statistics
            $ArchiveMailboxStats = @()

                if ($PSBoundParameters.ContainsKey("WMIRemoting")) {
                $UsePSRemote = $true
                }else {
                $UsePSRemote = $false}

        #endregion other variables

        #region aggregates

            #Object to store aggregates like total archive and mailbox size
            [decimal]$archive_total_size = 0
            [int]$archive_total_count = 0
            [decimal]$archive_average_size = 0

            [int]$mailbox_total_count = 0
            [int]$mailbox_total_local_count = 0
            [int]$mailbox_total_Remote_count = 0
            $DBSizes
            

            $GlobalAggregates = New-Object PSObject 
            $GlobalAggregates  | add-member Noteproperty archive_total_size $archive_total_size
            $GlobalAggregates  | add-member Noteproperty archive_total_count $archive_total_count
            $GlobalAggregates  | add-member Noteproperty archive_average_size $archive_average_size
            $GlobalAggregates  | add-member Noteproperty mailbox_total_count $mailbox_total_count
            $GlobalAggregates  | add-member Noteproperty mailbox_total_local_count $mailbox_total_local_count
            $GlobalAggregates  | add-member Noteproperty mailbox_total_Remote_count $mailbox_total_Remote_count
            $GlobalAggregates  | add-member Noteproperty DBSizes $DBSizes

        #endregion aggregates

        #region getting script parameter set

            switch ($PsCmdlet.ParameterSetName) {
                "ServerFilter"        { $ScriptFilter = "ServerFilter"; break} 
                "DAGFilter"           { $ScriptFilter = "DAGFilter"; break} 
                "IncludeServerFilter" { $ScriptFilter = "IncludeServerFilter"; break}
            } 
            Write-Verbose -Message "Info : Script Parameter Set Name : $ScriptFilter"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Script Parameter Set Name is $ScriptFilter"


        #endregion getting script parameter set
        

    #endregion Section 3.2 : Global variables

    #region Log Info 
             
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module 3 Factory] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  

    #endregion Log Info      
                   

#endregion Module 3 : Factory



#++++++++++++++++++++++++++++++++++++     Module 4 : Process    ++++++++++++++++++++++++++++++++++++




#region Module 4 : Process
 

    #region Log Info 

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 4 Process] : Starting"    
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 

    #endregion Log Info 


    #region Section 4.1 Getting general Information [unstructured data]

        _status " 2. Getting Global Information" 1
        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 2 of 6 : Getting Global Information" -percentComplete 40

        #region Get-ExchangeServer 
            
            _status "     2.1 Get-ExchangeServer Info (with filter if any)" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
            
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Get-ExchangeServer Info (with filter if any)"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Entering _parameterGetExch function"

            $ExchangeServers = [array](_parameterGetExch $ScriptFilter )

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Exiting _parameterGetExch function"

            _status "         - Exchange Servers detected : $($ExchangeServers.count)" 3

            $ExchangeServersList = @($ExchangeServers | foreach{ $_.name})

            #region Log Info 

                
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): ---- Script Filter Result----"
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Number of Exchange Servers detected : $($ExchangeServers.count) "
                

                foreach ($var in $ExchangeServersList) {


                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Server : $var"

                }
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): ---- Script Filter Result----" 
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Exiting _parameterGetExch function"
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
            
            #endregion Log Info 
     

        #endregion Get-ExchangeServer



        #region Get-Mailboxdatabase- Status 

            _status "     2.2 Get-MailboxDatabase -Status Info (with filter if any)" 2  

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Get-MailboxDatabase  Info (with filter if any)"

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Entering _parameterGetDBs function"
            
                       
            $Databases =  [array](_parameterGetDBs $ScriptFilter $E2010 $E2013 $ServerFilter $ExchangeServersList) 
            _status "         - Databases : $($Databases.count)" 3
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Number of database detected : $($Databases.count) "

            if ($Databases.count) {
                $DatabasesList = @($Databases | foreach{ $_.name})

                foreach ($var in $DatabasesList) {

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): DB : $var"

                }
            }

            #log info
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Exiting _parameterGetDBs function"
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
              

        #endregion Get-Mailboxdatabase - Status



        #region Get-RemoteMailbox    
    
            _status "     2.3 Get-RemoteMailbox Info" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Get-RemoteMailbox Info"

            if ($E2010) {
                $RemoteMailboxes = @() 
                $RemoteMailboxes = [array](Get-RemoteMailbox  -ResultSize Unlimited)
                
                    _status "         - Remote Mailboxes : $($RemoteMailboxes.count)" 3
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Number of remote mailboxes detected : $($RemoteMailboxes.count)"
                
            }
            else {
                Write-Verbose -Message " Info : Since the Exchange PowerShell Environment is less than 2010, RemoteMailboxes will be assigned 0"
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Since the Exchange PowerShell Environment is less than 2010, RemoteMailboxes will be assigned 0"
                $RemoteMailboxes = $null
                _status "         - Remote Mailboxes : 0" 3
            }             

            $ExchangeEnvironment.Add("RemoteMailboxes",$RemoteMailboxes.Count)    
            $GlobalAggregates.mailbox_total_Remote_count = $RemoteMailboxes.Count  
            
            
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"   

        #endregion Get-RemoteMailbox

            

        #region Get-Mailbox

             _status "     2.4 Get-Mailbox Info (with filter if any)" 2

             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Get-Mailbox Info (with filter if any)"

              if ($Databases.count -gt 0) {                  
                  
                  $Mailboxes =  [array](_parameterGetMailboxes $Databases) 
                  
                  Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Number of Mailboxes detected : $($Mailboxes.count)"         
              
               }else {
                   Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): since no databases are detected, Mailbox count = $($Mailboxes.count)"
               }

              
             _status "         - Mailboxes : $($Mailboxes.count)" 3
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"

        #endregion Get-Mailbox

              

        #region Get-Mailbox -Archive
            _status "     2.5 Get-Mailbox -Archive Info (with filter if any)" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section Get-Mailbox -Archive Info (with filter if any)"

            if ($E2010) {

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Since E2010 variable is true..continue"

                    if ($Databases.count -gt 0) {
                        $ArchiveMailboxes =  [array](_parameterGetArchiveMailboxes $DatabasesList)
                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Since databases.count is not 0..continue"
                    }
                       

            }else {
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Since E2010 variable is false : ArchiveMailboxes and ArchiveMailboxStats = null"
                    $ArchiveMailboxes = $null
                    $ArchiveMailboxStats = $null
            }

            _status "         - Archive mailboxes : $($ArchiveMailboxes.count)" 3

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Archive Mailboxes count = $($ArchiveMailboxes.count) "
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"


        #endregion Get-Mailbox -Archive



        #region Get-DatabaseAvailabilityGroup

            _status "     2.6 Getting DAG Info  (with filter if any)" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting DAG Info  (with filter if any)"

            if ($E2010) {
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Since E2010 variable is true..continue"

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Entering _parameterGetDAG function"

                    $DAGs =  [array](_parameterGetDAG $ServerFilter $InputDAGs $ExchangeServersList)

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): +++++++Exiting_parameterGetDAG function"

                    $Dag_Count = $DAGs.count

                    _status "         - DAGs : $Dag_Count" 3

                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): DAG count = $($DAGs.count)"
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"
            }else {
                $DAGs = $null
                _status "         - DAGs : 0" 3
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Since E2010 variable is false..DAG Count and Info are Null"
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp)"
            }               

    #endregion Get-DatabaseAvailabilityGroup



        #region warning screen output if $E2010 is false        

            if(!$E2010) {
                _status " Warning: Since the PowerSehll Host is not at least EX2010, the following assumptions are taken" 1
                _status "    - DAGs : 0" 3
                _status "    - Archive Mailboxes : 0" 3
                _status "    - Remote Mailboxes : 0" 3
            }

        #endregion warning screen output if $E2010 is false



        #region aggregates

             #Quote : Script block baseon on Steve's script         
             $ExchangeEnvironment.Add("TotalMailboxes",$Mailboxes.Count + $ExchangeEnvironment.RemoteMailboxes);
             $GlobalAggregates.mailbox_total_local_count = $Mailboxes.Count 
             $GlobalAggregates.mailbox_total_count = $GlobalAggregates.mailbox_total_local_count + $GlobalAggregates.mailbox_total_Remote_count
             #End Quote :  Script block baseon on Steve's script
             Write-Verbose -Message ""
             Write-Verbose -Message " Info : Total Mailboxes = Local Mailboxes + Remote Mailboxes = $($ExchangeEnvironment.TotalMailboxes)"
             Write-Verbose -Message ""
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section Aggregates"
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Total Mailboxes = Local Mailboxes + Remote Mailboxes = $($ExchangeEnvironment.TotalMailboxes)"

         #endregion aggregates
     


     #endregion #region Section 4.1 Getting general Information [unstructured data]        
        

    #region Section 4.2 Getting Structured data 

        _status "3. Processing Info and creating structured data from the info collected so far" 1
        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 3 of 6 : Processing Info and creating structured data from the info collected so far" -percentComplete 60


        #region 4.2.1 getting Exchange Server structured data

            _status "     3.1 Getting Detailed information for each Exchange Server (count:$($ExchangeServers.Count) )" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting Exchange Server structured data (_GetExSvr function)" 

            # Quote : Script block taken from Steve's script

                for ($i=0; $i -lt $ExchangeServers.Count; $i++)
                {
                    $displaycount = $i +1
                    $displaycountmax = $ExchangeServers.Count
                   # Write-Progress -id 2 -Activity "Getting Exchange Server Info"  -Status " Processing $($ExchangeServers[$i])"   -percentComplete (($displaycount/$displaycountmax)*100) -ParentId 1

	                _status "         $displaycount/$displaycountmax :  Getting Server WMI Info: $($ExchangeServers[$i])" 3

	                # Get Exchange Info
	                $ExSvr = _GetExSvr -E2010 $E2010 -ExchangeServer $ExchangeServers[$i] -Mailboxes $Mailboxes -Databases $Databases $UsePSRemote
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [_GetExSvr function]) Processing $($ExchangeServers[$i])  "

	                # Add to site or pre-Exchange 2007 list
	                if ($ExSvr.Site)
	                {
		                # Exchange 2007 or higher
		                if (!$ExchangeEnvironment.Sites[$ExSvr.Site])
		                {
			                $ExchangeEnvironment.Sites.Add($ExSvr.Site,@($ExSvr))
		                } else {
			                $ExchangeEnvironment.Sites[$ExSvr.Site]+=$ExSvr
		                }
	                } else {
		                # Exchange 2003 or lower
		                if (!$ExchangeEnvironment.Pre2007["Pre 2007 Servers"])
		                {
			                $ExchangeEnvironment.Pre2007.Add("Pre 2007 Servers",@($ExSvr))
		                } else {
			                $ExchangeEnvironment.Pre2007["Pre 2007 Servers"]+=$ExSvr
		                }
	                }
	                # Add to Servers List
	                $ExchangeEnvironment.Servers.Add($ExSvr.Name,$ExSvr)
                }

                #Write-Progress -Completed -Activity "Getting Exchange Server Info" -id 2 -ParentId 1


            # END Quote : Script block taken from Steve's script


        #endregion 4.2.1 getting Exchange Server structured data


        #region Section 4.2.2 Total Servers and mailboxes by role and feature   
        
             _status "     3.2 Getting Total servers by role and mailboxes by version" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting Total servers by role and mailboxes by version"

             # Quote : Script block taken from Steve's script
     
            $ExchangeEnvironment.Add("TotalMailboxesByVersion",(_TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment))

            $ExchangeEnvironment.Add("TotalServersByRole",(_TotalsByRole -ExchangeEnvironment $ExchangeEnvironment))

             # END Quote : Script block taken from Steve's script

         #endregion Section 4.2.2 Total Servers and mailboxes by role and feature

        
        #region Section 4.2.3 DAG Information

        # Quote : Script block taken from Steve's script
        
            
            _status "     3.3 Getting DAG Info" 2

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting DAG Info (_GetDAG function)"

            if ($DAGs) {
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp):"
                Log-Write -LogFullPath $DetailedFullPath -LineValue "$(get-timestamp): +++++++Entering _GetDAG function"
        
	            foreach($DAG in $DAGs)
	            {
		            $ExchangeEnvironment.DAGs+=(_GetDAG -DAG $DAG)
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [_GetDAG function] : Processing $($DAG.Name.ToUpper()) "
                    _status "         - Getting DAG Info: $($DAG.Name.ToUpper())" 3
	            }
            }

        # END Quote : Script block taken from Steve's script

    #endregion Section 4.2.3 DAG Information


        #region Section 4.3.3 DB Information

            
            <# Three types od databases we are concerned about
                - Recovery Databases [stored in  $ExchangeEnvironment.RecoveryDatabases]
                - Non Recovery Databases , which can be 
                    - DAG DB   [stored in $ExchangeEnvironment.DAGs[$j].Databases]
                    - Non DAG DB  [stored $ExchangeEnvironment.NonDAGDatabases] 
            #>

            
            
            #region loginfo
            _status "     3.4 Getting DB Info (count:$($Databases.Count))" 2
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting DB Info (_GetDB function)"
            #endregion log info

            #region variables
                $AverageMailboxSize_Eachdb = 0
                
            #endregion variables

            #region lopping through each database object from Get-MailboxDatabase -status variable named $Databases

                
               
                for ($i=0; $i -lt $Databases.Count; $i++) {
	
                    $displaycount = $i +1
                    $displaycountmax = $Databases.Count
                    #Write-Progress -id 3 -activity "Getting Database Info"  -Status " Processing $($Databases[$i].Name.ToUpper())"   -percentComplete (($displaycount/$displaycountmax)*100) -ParentId 1
                   


                    
                    # we will loop through databases and get a databaseOBJ containing our customized props

                    #this variable is true if the DB belongs to a DAG
                    # we use this variable when we query our DAG databases and compare it with the database in the loop.
                    $DAGDB = $false 
                    
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): ---------- "
                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Processing $($Databases[$i].Name.ToUpper()) via _GetDB function "
                    Write-Host "$($Databases[$i].Name).." -NoNewline
                   
                    $DatabaseObj = _GetDB -Database $Databases[$i] -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $E2010
	
                    

                    #region if the DB is recovery database
        
    
                        If (($DatabaseObj.IsRecovery) -eq $true) {                           

                            $ExchangeEnvironment.RecoveryDatabases += $DatabaseObj
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - $($Databases[$i].Name.ToUpper()) is a recovery DB "
                        } 

                    #endregion if the DB is recovery database 


                    #region if the DB is not recovery database

                        else { # in case the database is not recovery

                            $ExchangeEnvironment.NonRecoveryDatabases += $DatabaseObj

	                        # Quote : Script block taken from Steve's script

                            
                           

                            for ($j=0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++) {

		                        if ($ExchangeEnvironment.DAGs[$j].Members -contains $DatabaseObj.ActiveOwner) {
		
			                        $DAGDB=$true

			                        $ExchangeEnvironment.DAGs[$j].Databases += $DatabaseObj

                                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - $($Databases[$i].Name.ToUpper()) is member of DAG : $($ExchangeEnvironment.DAGs[$j].Name.ToUpper()) "
		                        }

	                        } # or ($j=0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++)


	                        if (!$DAGDB) {
	
		                        $ExchangeEnvironment.NonDAGDatabases += $DatabaseObj
                                
                                 Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):     - $($Databases[$i].Name.ToUpper()) is a Non Recovery, Non DAG Database"

	                        }

                          #END Quote


                     #region Aggregate for each DB that is not recovery

                
                            $GlobalAggregates.archive_total_size += $databaseobj.ArchiveTotalSize
               
                            $GlobalAggregates.archive_total_count += $databaseobj.ArchiveMailboxCount
                
                            $GlobalAggregates.DBSizes += $DatabaseObj.Size
                            

                            #$AverageMailboxSize_eachdb += $databaseObj.MailboxAverageSize


                    #endregion Aggregate for each DB that is not recovery


	                } # in case the database is not recovery

                #endregionif the DB is not recovery database

                
	
             } # for ($i=0; $i -lt $Databases.Count; $i++) 

                #Write-Progress -Completed -Activity "Getting Database Info" -Id 3 -ParentId 1

         #endregion lopping through each database object from Get-MailboxDatabase -status variable named $Databases

      #endregion Section 4.3.3 DB Information




      #region Section 4.3.4 DB Aggregates

        
        #region log info
            Write-host " " 
            _status "     3.5 Calculating Aggregate information" 2
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Section : Getting Total Aggregates"
       #endregion log info
        
       

        if ($GlobalAggregates.archive_total_count) {
            $GlobalAggregates.archive_average_size = "{0:N2}" -f (($GlobalAggregates.archive_total_size /$GlobalAggregates.archive_total_count) /1MB)
        }
        if ($GlobalAggregates.archive_total_size) {
            $GlobalAggregates.archive_total_size = "{0:N2}" -f ($GlobalAggregates.archive_total_size/1GB)  
        }

        if ($GlobalAggregates.DBSizes) {
        $GlobalAggregates.DBSizes = "{0:N2}" -f ($GlobalAggregates.DBSizes/1GB)  
        }

              
        $ExchangeEnvironment.Add("TotalArchivesCount",$GlobalAggregates.archive_total_count)
        $ExchangeEnvironment.Add("TotalArchivesSize",$GlobalAggregates.archive_total_size)
        $ExchangeEnvironment.Add("Average_Archive_size",$GlobalAggregates.archive_average_size) 
        $ExchangeEnvironment.Add("DBSizes",$GlobalAggregates.DBSizes)        
         

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Total Archives Count: $($ExchangeEnvironment.TotalArchivesCount) "
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Total Archives Size: $($ExchangeEnvironment.TotalArchivesSize) GB"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Average Archive size: $($ExchangeEnvironment.Average_Archive_size) MB"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Total DB Sizes: $($ExchangeEnvironment.DBSizes) GB"
        
         


    
    #endregion Section 4.3.4 DB Aggregates


  #endregion Section 4.3 Getting Structured data


    #region Log Info 
             
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module 4 Process] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  

    #endregion Log Info 


    


#endregion Module 4 : Process




#++++++++++++++++++++++++++++++++++++      Module 5 : Output     ++++++++++++++++++++++++++++++++++++




#region Module 5 : Output
 

    #region Log Info 

        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 4 of 6 : Generating HTML" -percentComplete 90

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 5 Output] : Starting"    
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 

         _status " 4. Generating Output HTML File" 1

    #endregion Log Info 


    #region Section 5.1 Main HTML Output Overall statistics tables
        
         _status "     4.1 Generating Output HTML" 2
         Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Generating Main HTML Output"

        
        #region draw HTML headings
            
            $Output="<html>
            <body>
            <font size=""1"" face=""Arial,sans-serif"">
            <h3 align=""center"">Exchange Environment Report</h3>
            <h5 align=""center"">Generated $((Get-Date).ToString())</h5>
            </font>"

        #endregion draw HTML headings


        #region draw HTML Table for Exchange Server version and roles

            $Output +="  <table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">
            <tr bgcolor=""#009900"">
            <th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count)""><font color=""#ffffff"">Total Servers:</font></th>"
            if ($ExchangeEnvironment.RemoteMailboxes)
                {
                $Output+="<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count+2)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
                } else {
                $Output+="<th colspan=""$($ExchangeEnvironment.TotalMailboxesByVersion.Count+1)""><font color=""#ffffff"">Total Mailboxes:</font></th>"
                }
            $Output+="<th colspan=""$($ExchangeEnvironment.TotalServersByRole.Count)""><font color=""#ffffff"">Total Roles:</font></th></tr>
            <tr bgcolor=""#00CC00"">"

           
            $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort Name| %{$Output+="<th>$($ExVersionStrings[$_.Key].Short)</th>"}
            $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort Name| %{$Output+="<th>$($ExVersionStrings[$_.Key].Short)</th>"}
            if ($ExchangeEnvironment.RemoteMailboxes)
            {
                $Output+="<th>Office 365</th>"
            }
            $Output+="<th>Org</th>"
            $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort Name| %{$Output+="<th>$($ExRoleStrings[$_.Key].Short)</th>"}
            $Output+="<tr>"
            $Output+="<tr align=""center"" bgcolor=""#dddddd"">"
            $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort Name| %{$Output+="<td>$($_.Value.ServerCount)</td>" }
            $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort Name| %{$Output+="<td>$($_.Value.MailboxCount)</td>" }
            if ($RemoteMailboxes)
            {
                $Output+="<th>$($ExchangeEnvironment.RemoteMailboxes)</th>"
            }
            $Output+="<td>$($ExchangeEnvironment.TotalMailboxes)</td>"
            $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort Name| %{$Output+="<td>$($_.Value)</td>"}
            $Output+="</tr><tr><tr></table><br>"

        #endregion draw HTML Table for Exchange Server version and roles


        #region draw table for mailbox types

            $MailboxesTypes= $Mailboxes  |Group-Object -Property Recipienttypedetails

            $Output+="<font size=""1"" face=""Arial,sans-serif""></font>
            <table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">
            <tr bgcolor=""#8E1275"">"
            $Output+= "<th colspan=""$($MailboxesTypes.Count)""><font color=""#ffffff"">Mailbox Types</font></th>"

            $Output+="<tr bgcolor=""#E46ACB"">"
            
            $MailboxesTypes |Sort -Descending Count| %{$Output+="<th><font color=""#ffffff"">$($_.Name)</font></th>"}
            $Output+="<tr>"
            $Output+="<tr align=""center"" bgcolor=""#dddddd"">"
            $MailboxesTypes |Sort -Descending Count| %{$Output+="<td>$($_.count)</td>" }
            $Output+="</tr><tr><tr></table><br />"

        #endregion draw table for mailbox types


        #region draw table for overall mailbox and archive statistics

            $Output+="<table border=""0"" cellpadding=""4"" style=""font-size:8pt;font-family:Arial,sans-serif"">
            <tr bgcolor=""#9D9D00"">"
            $Output+= "<th colspan= ""5"" ><font color=""#ffffff"">General Statistics</font></th>"
            $Output+="<tr bgcolor=""#DFE32D"">"
            $Output+="<th><font color=""#000000"">Total Mailbox Count</font></th>"
            $Output+="<th><font color=""#000000"">Total DB Size (GB)</font></th>"            
            $Output+="<th bgcolor=""#FFA500""><font color=""#000000"">Total Archive Count</font></th>"
            $Output+="<th bgcolor=""#FFA500""><font color=""#000000"">Total Archive Sizes (GB)</font></th>"
            $Output+="<th bgcolor=""#FFA500""><font color=""#000000"">Average Archive Size (GB)</font></th>"
            $Output+="<tr>"
            $Output+="<tr align=""center"" bgcolor=""#dddddd"">"
            $Output+="<td><font color=""#000000"">$($ExchangeEnvironment.TotalMailboxes) Mailboxes</font></td>" 
            $Output+="<td><font color=""#000000"">$($ExchangeEnvironment.DBSizes) GB</font></td>"             
            $Output+="<td ><font color=""#000000"">$($ExchangeEnvironment.TotalArchivesCount) Archives</font></td>" 
            $Output+="<td ><font color=""#000000"">$($ExchangeEnvironment.TotalArchivesSize) GB</font></td>" 
            $Output+="<td  ><font color=""#000000"">$($ExchangeEnvironment.Average_Archive_size) GB</font></td>" 
            $Output+="</tr><tr><tr></table><br />"


        #endregion draw table for overall mailbox and archive statistics


       
    #endregion Section 5.2 Main HTML Output Overall statistics tables


    #region Section 5.2 Sites and Servers


        foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {
            
	       $Output+=_GetOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings
        }
        
        foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator()) {
            
	       $Output+=_GetOverview -Servers $FakeSite -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings -Pre2007:$true
        }



    #endregion Section 5.2 Sites and Servers

     
    #region Section 5.3 DAG/DB info     
    
        foreach ($DAG in $ExchangeEnvironment.DAGs) {

	        if ($DAG.MemberCount -gt 0) {
	        
		        # Database Availability Group Header
		        $Output+="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
		        <col width=""20%""><col width=""10%""><col width=""70%"">
		        <tr align=""center"" bgcolor=""#FF8000 ""><th>Database Availability Group Name</th><th>Member Count</th>
		        <th>Database Availability Group Members</th></tr>
		        <tr><td>$($DAG.Name)</td><td align=""center"">
		        $($DAG.MemberCount)</td><td>"
		        $DAG.Members | % { $Output+="$($_) " }
		        $Output+="</td></tr></table>"
		
		        # Get Table HTML
		        $Output+=_GetDBTable -Databases $DAG.Databases
	        }
	
        } 




    #endregion Section 5.3 DAG/DB Info  


    #region Section 5.4 Non DAG Databases

        if ($ExchangeEnvironment.NonDAGDatabases.Count) {
	
	        $Output+="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
    	    <tr bgcolor=""#FF8000""><th>Mailbox Databases (Non-DAG)</th></table>"
	        $Output+=_GetDBTable -Databases $ExchangeEnvironment.NonDAGDatabases
         }

    #endregion Section 5.5 Non DAG Databases


    #region Section 5.5 Recovery DBs

        if ($ExchangeEnvironment.RecoveryDatabases.Count) {

            $Output+="<table border=""0"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
    	    <tr bgcolor=""#000000""><th><font color=""#ffffff""><strong>Recovery Databases </strong></font></th></table>"
	        $Output+=_GetDBTable -Databases $ExchangeEnvironment.RecoveryDatabases



        }

    #endregion Section 5.5 Recovery DBs


    #region Generating output HTML

        $Output+="</body></html>";

        #creating HTML file

        $HTMLReport = Join-Path $ScriptFilesPath  "HTMLReport.Html"

        
        $var = $ErrorActionPreference                           
        $ErrorActionPreference = "stop"

        try{
            $Output | Out-File $HTMLReport
        } catch {
            Write-CorpError -myError $_ -Info "Module 5 Output - Fail saving HTM report at $HTMLReport" -mypath $ErrorFullPath
        }finally {
        $ErrorActionPreference = $var
        }

        

    #endregion Generating Output HTML


    #region Section 5.6 DB Preference HTML

        _status "     4.2 Generating DB Activation Output HTML" 2  
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Generating DB Activation HTML Output" 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Conditions to generate is : E2010 variable is true and (ExchangeEnvironment.DAGs.count) variable is true " 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - In simple words, if you are running new PowerShell host, and there are DAGs detected using your filter " 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - E2010 Value is $E2010 " 
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - ExchangeEnvironment.DAGs.count Value is $($ExchangeEnvironment.DAGs.count) " 


        if ( $E2010 -and ($ExchangeEnvironment.DAGs.count) ) { 
        
               
                Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  -  Conditions are met ! Proceed "
                #Open an HTML Document

                $Output2 ="<html>
                <body>
                <font size=""1"" face=""Arial,sans-serif"">
                <h3 align=""center"">DAG Database Copies layout</h3>
                <h5 align=""center"">Generated $((Get-Date).ToString())</h5>
                </font>"
            
            
            
                foreach ($dag in $ExchangeEnvironment.DAGs) {
                
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Inspecting $($dag.name.toupper())"
                            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):   - Does this DAG has members and those members have databases? "

                            if($dag.members.count -and $dag.databases.count) {

                                 Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):   - Does this DAG has members and those members have databases? Answer is Yes "

                                    #Testing if we have all DAG members after the filter
            
                                    $test = Test-DAGFullMembers $DAG $ExchangeServersList

                                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Are all member servers of $($dag.name.toupper()) are available to us after applying user filter? "
                                    Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  -Entering (Test-DAGFullMembers function) "

                                    

                                    if ($test) { 
                                        
                                        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Answer is $test . Processing"

                                        $hash_server_vs_db_Aggregates = @{}
                                        #return a hash table with servers in the DAG and count of DBs prefered on it
                                        $hash_server_vs_db_Aggregates  = _GetDBPreference_Table_Info $ExchangeEnvironment $dag


                                        # Writing Two Tables
                                        $Output2 += _GetDBPreference_Table_HTML $DAG $ExchangeEnvironment $hash_server_vs_db_Aggregates
                                         Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Done creating DAG Activation diagram for DAG : $($dag.name.toupper()) "

                                        # Append Space between each DAG Table
                                        $Output2 += "<br /><br /><br /><br /><br />"
                            
                                        } # if ($test)
                                        else {
                                         Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Answer is $test . Skipping getting DB Activation Table for DAG : $($dag.name.toupper()) "

                                        }
                            } 
                            else {
                            # if($dag.members.count -and $dag.databases.count)
                             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):   - Does this DAG has members and those members have databases? Answer is No "
                             
                            }
                    

                            
                            


                } # foreach ($dag in $ExchangeEnvironment.DAGs)
            

                # Close HTML Document
                $Output2 += "</body></html>"
            
                #Output the HTML document
                $HTMLReport = Join-Path $ScriptFilesPath  "DB DAG Layout.Html" 
                
                $var = $ErrorActionPreference                           
                $ErrorActionPreference = "stop"

                try{
                    $Output2 | Out-File $HTMLReport
                } catch {
                    Write-CorpError -myError $_ -Info "Module 5 Output - Fail saving HTM report at $HTMLReport" -mypath $ErrorFullPath
                }finally {
                $ErrorActionPreference = $var
                }

            
           


        } # if ( $E2010 -and ($ExchangeEnvironment.DAGDatabases.count) ) 
        else {
            #skipping HTML report
            Write-Warning -Message " Skipping creating DB Activation HTML"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Skipping creating DB Activation HTML " 
        }
    


    #endregion Section 5.6 DB Preference HTML


    #region Log Info 
             
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module 5 Output] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  

    #endregion Log Info 
   

#endregion Module 5 : Output




#++++++++++++++++++++++++++++++++++++      Module 6 : Charts   ++++++++++++++++++++++++++++++++++++




#region Module 6 : Charts


    #region Log Info 

        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 5 of 6 : Generating Charts" -percentComplete 90
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 6 Charts] : Starting"    
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 
        _status " 5. Generating Charts" 1

    #endregion Log Info       
     
     
    #region Section 6.1 Chart Test
         $chart_Test = $true
         
         Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): "

        $var = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try {

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Testing if (Microsoft Chart Controls for Microsoft .NET Framework) is installed on this machine"
            [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
            $testchart  = new-object System.Windows.Forms.DataVisualization.Charting.Chart
             Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Test result is Success."

        } Catch {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Test result is FAIL. Please make sure (Microsoft Chart Controls for Microsoft .NET Framework) is installed on this machine "  
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Download Microsoft Chart Controls for Microsoft .NET Framework 3.5 (http://www.microsoft.com/en-us/download/details.aspx?id=14422)" 
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  - Skipping Chart Module" 
            $chart_Test = $false 
            _status "      Skipping creating charts... check Info log for detials" 2 
        }
        Finally {
        $ErrorActionPreference = $var

        }


    #endregion Section 6.1 Chart Test

    if ($chart_Test) {
    #region Section 6.2 Chart DB Vs Mailbox Count

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Processing : DB vs Mailbox count chart"
        if (($ExchangeEnvironment.chart_db).count) {
            $var = $ExchangeEnvironment.chart_db
            $chart= Join-Path $ScriptFilesPath "DB_MailboxCount.png"                            
            Get-Corpchart_light -data $var -obj_key Name -obj_value MailboxCount -filepath $chart -Type column -title_text "DB vs Mailboxcount" -chartarea_Xtitle "DB" -chartarea_Ytitle "Mailbox count" -sort asc -append_date_title -IsvalueShownAsLabel -chart_color "magenta" -ErrorFullPath $ErrorFullPath

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Completed : DB vs Mailbox count chart"
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Skipping : No Data available"

        }   

    #endregion Section 6.2 Chart DB Vs Mailbox Count


    #region Section 6.3 Chart DB Vs Backup

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Processing : DB vs Backup days"
        if (($ExchangeEnvironment.chart_db).count) {
            $var = $ExchangeEnvironment.chart_db
            $chart= Join-Path $ScriptFilesPath "DB_Backup.png"                            
            Get-Corpchart_light -data $var -obj_key Name -obj_value Backup -filepath $chart -Type column -title_text "DB vs Backup days" -chartarea_Xtitle "DB" -chartarea_Ytitle "Backup Since (days) [-1 = never backed up]" -sort asc -append_date_title -IsvalueShownAsLabel -chart_color "blue" -ErrorFullPath $ErrorFullPath

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Completed : DB vs Backup days chart"
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Skipping : No Data available"

        }   

    #endregion Section 6.3 Chart DB Vs Backup


    #region Section 6.3 Chart DB Vs Size

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Processing : DB vs Size in GB chart"
        if (($ExchangeEnvironment.chart_db).count) {
            $var = $ExchangeEnvironment.chart_db
            $chart= Join-Path $ScriptFilesPath "DB_Size.png"                            
            Get-Corpchart_light -data $var -obj_key Name -obj_value DBSize -filepath $chart -Type column -title_text "DB vs Size in GB" -chartarea_Xtitle "DB" -chartarea_Ytitle "Size (GB)" -sort asc -append_date_title -IsvalueShownAsLabel -chart_color "DarkRed" -ErrorFullPath $ErrorFullPath

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Completed : DB vs Size in GB chart"
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Skipping : No Data available"

        }   

    #endregion Section 6.3 Chart DB Vs Size


    #region Section 6.4 Chart Mailbox Server Vs Mailbox Count

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Processing : Mailbox Server Vs Mailbox CountB chart"
        if (($ExchangeEnvironment.Chart_Srv).count) {
            $var = $ExchangeEnvironment.Chart_Srv
            $chart= Join-Path $ScriptFilesPath "Srv_MailboxCount.png"                            
            Get-Corpchart_light -data $var -obj_key Name -obj_value MailboxCount -filepath $chart -Type column -title_text "Mailbox Server Vs Mailbox Count" -chartarea_Xtitle "MBX" -chartarea_Ytitle "Mailbox Count" -sort asc -append_date_title -IsvalueShownAsLabel -chart_color "blue" -ErrorFullPath $ErrorFullPath

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Completed : Mailbox Server Vs Mailbox Count chart"
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Skipping : No Data available"

        }   

    #endregion Section 6.4 Chart Mailbox Server Vs Mailbox Count


    #region Section 6.4 Chart Mailbox Server Vs Mounted DBs

        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Processing : Mailbox Server Vs Mounted DBs chart"
        if (($ExchangeEnvironment.Chart_Srv).count) {
            $var = $ExchangeEnvironment.Chart_Srv
            $chart= Join-Path $ScriptFilesPath "Srv_DBMountedCount.png"                            
            Get-Corpchart_light -data $var -obj_key Name -obj_value DBMountedCount -filepath $chart -Type column -title_text "Mailbox Server Vs DBs Mounted" -chartarea_Xtitle "MBX" -chartarea_Ytitle "DBs Mounted" -sort asc -append_date_title -IsvalueShownAsLabel -chart_color "blue" -ErrorFullPath $ErrorFullPath

            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Completed : Mailbox Server Vs Mounted DBs chart"
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):  Skipping : No Data available"

        }   

    #endregion Section 6.4 Chart Mailbox Server Vs Mounted DBs

    } # if ($chart_Test)
    #region Log Info 
             
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module 6 Charts] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"  

    #endregion Log Info 


#endregion Module 6 : Charts




#++++++++++++++++++++++++++++++++++++      Module 7 : Final Tasks  ++++++++++++++++++++++++++++++++++++



#region Module 7 : Final Tasks

      #region Log Info 

        Write-Progress -id 1 -activity "Get-ExchangeOrgReport" -status "Phase 6 of 6 : Final Tasks" -percentComplete 99
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):[Module 7 Final Tasks] : Starting"    
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):" 

        _status " 6. Final Tasks" 1

      #endregion Log Info 

        
      #region Send Email
        
         if ($SendMail) {
            _status "     6.1 Sending Email" 2 
            $subject = "Mail Health Report _$(Get-Date -f 'yyyy-MM-dd')"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Sending Email :"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - To : $MailTo"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - From : $MailFrom "
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - SMTP Host : $mailsmtphost"
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): - Subject : $subject"        
            _sendEmail $MailFrom $MailTo $subject $MailServer $ScriptFilesPath $InfoFullPath $ErrorFullPath
        }else {
            Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Skipping Send Email as (SendEmail) parameter was not supplied"  

        }
         
      #endregion Send Email       
       
       
      #region Log Info 
             
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): [Module 7 Final Tasks] : Pass"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp):"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): $('-' * 60)"
        Log-Write -LogFullPath $InfoFullPath -LineValue "$(get-timestamp): Script run time (Minutes:seconds:milliseconds): $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString()):$($Watch.Elapsed.MilliSeconds.ToString())"   

        _screenFooter $ExchangeEnvironment $Databases $Watch

     #endregion Log Info 


      #region closing log files

        Log-Finish -LogFullPath $InfoFullPath
        Log-Finish -LogFullPath $ErrorFullPath
        Log-Finish -LogFullPath $DetailedFullPath
     
     #endregion closing log files


#endregion Module 7 : Final Tasks



#++++++++++++++++++++++++++++++++++++             END            ++++++++++++++++++++++++++++++++++++

