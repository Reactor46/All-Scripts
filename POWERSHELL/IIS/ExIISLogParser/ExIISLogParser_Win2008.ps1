<#
    .SYNOPSIS
    Parses IIS Log files for records relating to ActiveSync, Exchange Web Services and WebDAV
   
    Steve Goodman
    .DESCRIPTION
    Looks at logs files and produces a CSV file with summary data listing:
    * Activesync clients, whether it is proxied by another CAS and ActiveSync device type
    * EWS clients with columns for PC Outlook, Mac Outlook and Entourage 2008 EWS Edition and other clients
    * WebDAV clients including client versions.
   
    .PARAMETER LogFilePath
    Path to base directory of IIS Log files, e.g. "C:\WINDOWS\system32\LogFiles\W3SVC1"
   
    .PARAMETER Days
    How many days log files to look back by
   
    .PARAMETER OutputCSVFile
    File to write CSV output to
    
    .PARAMETER SaveStateFile
    File to save or load internal state to. Useful when looking at multiple CAS servers or to update output later on based on only more recent logfiles
   
    .EXAMPLE
    Parses log files from the default log directory "C:\WINDOWS\system32\LogFiles\W3SVC1" to "C:\output.csv"
    .\ExIISLogParser.ps1
    
    .EXAMPLE
    Parses the last 30 days of log files from the current directory  to cas_results.csv in the current directory, and saves the state to state.xml in the current directory
    .\ExIISLogParser.ps1 -LogFilePath ".\" -Days 30 -OutputCSVFile ".\cas_results.csv" -SaveStateFile ".\state.xml"
   
    #>
param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Full path to log file directory")][string]$LogFilePath = "C:\WINDOWS\system32\LogFiles\W3SVC1",
    [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Last write date of logs to search back by")][int]$Days=0,
    [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="CSV file for output")][string]$OutputCSVFile="C:\output.csv",
    [parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Script state XML file")][string]$SaveStateFile
    )
if (!(Test-Path $LogFilePath))
{
    throw "LogFilePath does not exist"
}
if ((Test-Path $OutputCSVFile))
{
    throw "OutputCSVFile already exists"
}
[hashtable]$Users = @{}
if ($SaveStateFile)
{
    if ((Test-Path $SaveStateFile))
    {
        $Users = Import-Clixml -Path $SaveStateFile
    }
}
$EarliestLogDate = (Get-Date).Subtract([timespan]"$($Days).00:00").Date
[array]$Files = Get-Item -Path "$($LogFilePath)\*.log" | Where {$_.LastWriteTime -gt $EarliestLogDate}
for ($i = 0; $i -lt $Files.Count; $i++)
{
    Write-Progress -id 1 -activity "Overall Progress" -status "File $($i) of $($Files.Count)" -percentComplete (($i/$Files.Count*100)*0.9); 
    Write-Progress -id 2 -activity "Log $($Files[$i].Name)" -status "Loading" -percentComplete 0
    $Log = Get-Content $Files[$i] | Where {$_ -like "*/Microsoft-Server-ActiveSync/*" -or $_ -like "*/EWS/*" -or $_ -like "*/exchange/*"}
    for ($j = 0; $j -lt $Log.Count; $j++)
        {
        Write-Progress -id 2 -activity "Current Log $($Files[$i].Name)" -status "Working" -percentComplete ($j/$Log.Count*100);
        # Clear Username
        $Username=$null
        
        # Split up log line
        $arrLog = $Log[$j].Split(" ");
        
        # Extract username first
        if ($arrLog[4] -eq "/Microsoft-Server-ActiveSync/Proxy")
        {
            foreach ($QSPart in $arrLog[5].Split("&"))
            {
                if ($QSPart -like "User=*") 
                { 
                    $Username =($QSPart.Split("="))[1]
                }
            }
            
        } else {
            # Not proxied, just get the username provided direct
            $Username = $arrLog[7]   
        }
        
        # Only process if it's an authenticated user
        if ($Username)
        {
            # Take off domain or UPN suffix
            if ($Username -like "*@*") 
            { 
                $Username = ($Username.Split("@"))[0] 
            } elseif ($Username -like "*\*") { 
                $Username = ($Username.Split("\"))[1];
            } elseif ($Username -like "*%40*") {
                $Username = ($Username.Split("%40"))[0] 
            } elseif ($Username -like "*/*") {
                $Username = ($Username.Split("/"))[1] 
            }
            
            # Make username lower case
            $Username = $Username.ToLower()
            
            # Update data for user
            switch -wildcard ($arrLog[4])
            {
                "/EWS/*" 
                {
                    # Get EWS Client
                    $EWSClient = $arrLog[9].Replace("+"," ")
                    # Check if already found or create new hashtable item
                    if (!$Users.Contains($Username))
                    {
                        [hashtable]$obj = @{ActiveSyncUser=0; ActiveSyncProxyUser=0; ActiveSyncClients=@{}; ActiveSyncLastAccess=""; EWSUser=0; EWSClients=@{}; EWSLastAccess=""; WebDavUser=0; WebDavClients=@{}; WebDavLastAccess=""}
                        $Users.Add($Username,$obj)
                    }
                    # Set variables
                    $Users[$Username]["EWSUser"]=1
                    
                    if (!$Users[$Username]["EWSClients"].Contains($EWSClient))
                    {
                        $Users[$Username]["EWSClients"].Add($EWSClient,1)
                    }
                    $Users[$Username]["EWSLastAccess"]=$arrLog[0]
                    break
                }
                "/Microsoft-Server-ActiveSync/*"
                {
                    # Check if already found or create new hashtable item
                    if (!$Users.Contains($Username))
                    {
                        [hashtable]$obj = @{ActiveSyncUser=0; ActiveSyncProxyUser=0; ActiveSyncClients=@{}; ActiveSyncLastAccess=""; EWSUser=0; EWSClients=@{}; EWSLastAccess=""; WebDavUser=0; WebDavClients=@{}; WebDavLastAccess=""}
                        $Users.Add($Username,$obj)
                    }
                    # Set variables
                    
                    # Is a ActiveSync user?
                    $Users[$Username]["ActiveSyncUser"]=1
                    # Is a proxy user?
                    if ($arrLog[4] -eq "/Microsoft-Server-ActiveSync/Proxy")
                    {
                        $Users[$Username]["ActiveSyncProxyUser"]=1
                    }
                    # Client Info
                    foreach ($QSPart in $arrLog[5].Split("&"))
                    {
                        if ($QSPart -like "DeviceType=*") 
                        { 
                            $ASClient = ($QSPart.Split("="))[1]
                            if (!$Users[$Username]["ActiveSyncClients"].Contains($ASClient))
                            {
                                $Users[$Username]["ActiveSyncClients"].Add($ASClient,1)
                            }
                        }
                    }
                    # Last Access Date
                    $Users[$Username]["ActiveSyncLastAccess"]=$arrLog[0]
                    break
                }
                "/exchange/*"
                {
                    # Check if already found or create new hashtable item
                    if (!$Users.Contains($Username))
                    {
                        [hashtable]$obj = @{ActiveSyncUser=0; ActiveSyncProxyUser=0; ActiveSyncClients=@{}; ActiveSyncLastAccess=""; EWSUser=0; EWSClients=@{}; EWSLastAccess=""; WebDavUser=0; WebDavClients=@{}; WebDavLastAccess=""}
                        $Users.Add($Username,$obj)
                    }
                    # Set variables
                    $Users[$Username]["WebDavUser"]=1
                    $WDClient = $arrLog[9].Replace("+"," ")
                    if (!$Users[$Username]["WebDavClients"].Contains($WDClient))
                    {
                        $Users[$Username]["WebDavClients"].Add($WDClient,1)
                    }
                    $Users[$Username]["WebDavLastAccess"]=$arrLog[0]
                    break
                }
            }
        }
    }
}
Write-Progress -id 1 -activity "Overall Progress" -status "Preparing Output" -percentComplete 95; 
[array]$Output=$null
$Users.GetEnumerator() | Foreach {
    $OutputItem = New-Object Object
    $OutputItem | Add-Member NoteProperty Username  $_.Key
    $OutputItem | Add-Member NoteProperty ActiveSyncUser $_.Value["ActiveSyncUser"]
    $OutputItem | Add-Member NoteProperty ActiveSyncProxyUser $_.Value["ActiveSyncProxyUser"]
    $ActiveSyncClients = $null
    $_.Value["ActiveSyncClients"].GetEnumerator() | % { $ActiveSyncClients += "$($_.Key); "}
    $OutputItem | Add-Member NoteProperty ActiveSyncClients $ActiveSyncClients
    $OutputItem | Add-Member NoteProperty ActiveSyncLastAccess $_.Value["ActiveSyncLastAccess"]
    $OutputItem | Add-Member NoteProperty EWSUser $_.Value["EWSUser"]
    $EWSPCOutlook = ""
    $EWSMacMail = ""
    $EWSMacOutlook = ""
    $EWSEntourage = ""
    $EWSOther = ""
    $_.Value["EWSClients"].GetEnumerator() | foreach { 
        $EWSClient = $_.Key
        switch -wildcard ($EWSClient)
        {
            "Microsoft Office*"
            {
                $EWSPCOutlook=$EWSClient
                break
            } 
            "Mac*ExchangeWebServices*"
            {
                $EWSMacMail=$EWSClient
                break
            } 
            "MacOutlook*"
            {
                $EWSMacOutlook=$EWSClient
                break
            } 
            "Entourage*"
            {
                $EWSEntourage=$EWSClient
                break
            }
            default
            {
               $EWSOther+="$($EWSClient); "
            }
            
        }
    }
    $OutputItem | Add-Member NoteProperty EWSPCOutlook $EWSPCOutlook
    $OutputItem | Add-Member NoteProperty EWSMacMail $EWSMacMail
    $OutputItem | Add-Member NoteProperty EWSMacOutlook $EWSMacOutlook
    $OutputItem | Add-Member NoteProperty EWSEntourage $EWSEntourage
    $OutputItem | Add-Member NoteProperty EWSOther $EWSOther
    $OutputItem | Add-Member NoteProperty EWSLastAccess $_.Value["EWSLastAccess"]
    $OutputItem | Add-Member NoteProperty WebDavUser $_.Value["WebDavUser"]
    $WebDavClients=$null
    $_.Value["WebDavClients"].GetEnumerator() | % { $WebDavClients += "$($_.Key); "}
    $OutputItem | Add-Member NoteProperty WebDavClients $WebDavClients
    $OutputItem | Add-Member NoteProperty WebDavLastAccess $_.Value["WebDavLastAccess"]
    $Output += $OutputItem
}

if ($SaveStateFile)
{
    $Users = Export-Clixml -Path $SaveStateFile
}

$Output[1..($Output.Count)] | Select Username,ActiveSyncUser,ActiveSyncProxyUser,ActiveSyncClients,ActiveSyncLastAccess,EWSUser,EWSPCOutlook,EWSMacMail,EWSMacOutlook,EWSEntourage,EWSOther,EWSLastAccess,WebDavUser,WebDavClients,WebDavLastAccess | Export-Csv -Path $OutputCSVFile -NoClobber -NoTypeInformation