<#
	.SYNOPSIS
		Configure LDAP Query logging for one or all domain controllers in domain

	.DESCRIPTION
        Script Asks for option to Enable, Disable,Check or Create Report for selected diagnostics
        Initially designed for Option #15 -  Field Engineering
        See https://support.microsoft.com/en-us/help/314980 for Details
        Report based on parsing Winlog events from all DomainControlles for 1644 Event

	.PARAMETER  ReportForLastnHours  
		When creating report file - get events from selected DC's for last 'n' Hours

	.PARAMETER  DiagnosticsName   
		NTDS Diagnostics Field to Configure. Default is "15 Field Engineering" - LDAP Query Logging

	.PARAMETER  DiagnosticsRegPath   
		NTDS Diagnostics Registry Path. Default is "15 Field Engineering" - LDAP Query Logging

	.EXAMPLE
		PS C:\> .\NTDS_Debug_Operations.ps1
            Select Debug Action
            Select Server where to Enable or Disable LDAP Quiery logging'.
            - Ensure you have Domain Admin Permissions.
            Select Cancel to exit when done.
            [S] Check Current Status  [E] Enable LDAP Query Logging  [D] Disable LDAP Query Logging  [R] Create Report  [C] Cancel  [?] Help (default is "S"):

	.OUTPUTS
		NTDS_%DiagnosticsName%_Report_%Date%.csv File in Selected Directory

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

	.LINK
		https://support.microsoft.com/en-us/help/314980/how-to-configure-active-directory-and-lds-diagnostic-event-logging

#>
param(
   [int]$ReportForLastnHours  = 1,                                                         # n Hours to fetch events for from Domain Controllers
[string]$DiagnosticsName      = "15 Field Engineering",                                    # diagnistics name according to https://support.microsoft.com/en-us/help/314980
[string]$DiagnosticsRegPath   = "SYSTEM\\CurrentControlSet\\services\\NTDS\\Diagnostics\\",# diagnistics path according to https://support.microsoft.com/en-us/help/314980
  [bool]$ThresholdsDefaults   = $true,                                                     # if set to $true - custom values not set for Inefficient / Expensive LDAP call thresholds and using defaults from next two options
[string]$ThresholdExpensive   = "10000",                                                   # means if an LDAP call visit 10,000 or more entries then it will be consider as an expensive call
[string]$ThresholdInefficient = "1000"                                                     # means if a query visit less than 1000 entries then it will not be consider inefficient query even though if it return no entry. 
                                                                                           # (Inefficient LDAP calls are the searches those return less than 10% of visited entries)
)
#-----------------------------------------------------------------------------------------------------------------------
Function CheckIfAdmin{
Begin
    {
        $ScriptName = $MyInvocation.MyCommand.ToString()
        $ScriptPath = $MyInvocation.MyCommand.Path
        $Username = $env:USERDOMAIN + "\" + $env:USERNAME
 
        $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = new-object System.Security.principal.windowsprincipal($CurrentUser)
        }
Process
    {      
        if (!$principal.IsInRole("Administrators")) 
        {
            Write-Host = 'You need to run this from an elevated prompt'
            exit
        }
    }
}
#-----------------------------------------------------------------------------------------------------------------------
Function mcConvertClient ($Client)
{ # Extract IP or Port from client string [0] based on [1]
	$mcReturn='Unknown'
		[regex]$regexIPV6 = '(?<IP>\[[A-Fa-f0-9:%]{1,}\])\:(?<Port>([0-9]+))'
		[regex]$regexIPV4 =	'((?<IP>(\d{1,3}\.){3}\d{1,3})\:(?<Port>[0-9]+))|(?<IP>(\d{1,3}\.){3}\d{1,3})'
		[regex]$KnownClient = '(?<IP>([G-Z])\w+)'
	    switch -regex ($Client[0])
		{ # $client[1] is either IP or Port
	        $regexIPV6 { $mcReturn = $matches.($client[1]) }
	        $regexIPV4 { $mcReturn = $matches.($client[1]) }
		    $KnownClient { $mcReturn = $matches.($client[1]) }
		}
	$mcReturn
}
#-----------------------------------------------------------------------------------------------------------------------
function Get-DomainController
{ 
#Requires -Version 2.0             
[CmdletBinding()]             
 Param              
   ( 
    #$Name, 
    #$Server, 
    #$Site, 
    [String]$Domain, 
    #$Forest 
    [Switch]$CurrentForest  
    )#End Param 
 
Begin             
{             
                
}#Begin           
Process             
{ 
 
if ($CurrentForest -or $Domain) 
   { 
    try 
        { 
            $Forest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()     
        } 
    catch 
        { 
            "Cannot connect to current forest." 
        } 
    if ($Domain) 
       { 
        # User specified domain OR Match 
        $Forest.domains | Where-Object {$_.Name -eq $Domain} |  
            ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
       } 
    else 
       { 
        # All domains in forest 
        $Forest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
       } 
   } 
else 
   { 
    # Current domain only 
    [system.directoryservices.activedirectory.domain]::GetCurrentDomain() | 
        ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
   } 
 
}#Process 
End 
{ 
 
}#End 
 
}
#-----------------------------------------------------------------------------------------------------------------------
Function ConfigureNtdsDiagnostics
{
param(
[switch]$Enable,
[switch]$Disable,
[switch]$CheckOnly,
[string[]]$Servers,
[string]$RegPath  = $Script:DiagnosticsRegPath,
[string]$RegName  = $Script:DiagnosticsName
)


if ($Enable) {[string]$RegValue = "5"} 
if ($Disable){[string]$RegValue = "0"}

if (!$Enable -and !$Disable -and !$CheckOnly){
Write-Host "Please Set: `n -Enable `n -Disable `n -CheckOnly" -ForegroundColor DarkRed -BackgroundColor White
exit
}

function fn_GetDiagValue
{
param([string]$Value)
    if ($Value -ne "")
    {
        switch($Value)
        {
        "0" { "0 Disabled"    }
        "1" { "1 (Minimal)"   }
        "2" { "2 (Basic)"     }
        "3" { "3 (Extensive)" }
        "4" { "4 (Verbose)"   }
        "5" { "5 (All events)"}
        }
    }else{"Unknown"}
}


if (!($Servers)){  [string[]]$Servers = Get-DomainController  }
[array]$response = @()

    foreach ($Server in $Servers)
    { 
    $obj = "" | Select Server,Diagnostics,Value,InefficientThreshold,ExpensiveThreshold
    $obj.server = $Server
    $obj.Diagnostics = $RegName

        If (test-connection -ComputerName $Server -Count 1 -Quiet)
        { 
        $reg          = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server) 
        $regKey       = $reg.OpenSubKey($RegPath,$true) 
        $CurrentValue = $regKey.GetValue($RegName)
    
        $regKey_Thresholds    = $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters",$true) 
        $regValue_InefficientThreshold = $regValue_ExpensiveThreshold = $null
                 
        $regValue_InefficientThreshold  = $regKey_Thresholds.GetValue("Inefficient Search Results Threshold")
        $regValue_ExpensiveThreshold    = $regKey_Thresholds.GetValue("Expensive Search Results Threshold")

            if ($CheckOnly){ $obj.Value  = fn_GetDiagValue -value $CurrentValue ; $obj.ExpensiveThreshold = if(!($regValue_ExpensiveThreshold)){"Default"}Else{$regValue_ExpensiveThreshold} ;$obj.InefficientThreshold = if(!($regValue_InefficientThreshold)){"Default"}else{$regValue_InefficientThreshold} }
            else{

                if ($CurrentValue -ne $RegValue){
                Write-Warning "$Server - $regName : $CurrentValue -> $RegValue"
                $regKey.SetValue($RegName,$RegValue,[Microsoft.Win32.RegistryValueKind]::DWord) 
                }else{ Write-Warning "$Server - Value already set:  $CurrentValue"}
                $CurrentValue = $regKey.GetValue($RegName)
                $obj.Value    = fn_GetDiagValue -value $CurrentValue 
                }


                if(($ThresholdsDefaults -eq $false) -and ($ThresholdExpensive -and $ThresholdInefficient))
                {
                
                    if ($Enable)
                    {
                        if (!($regValue_InefficientThreshold -eq $ThresholdInefficient)){
                        Write-Warning "$Server - Inefficient Search Results Threshold : $regValue_InefficientThreshold -> $ThresholdInefficient"
                        $regKey_Thresholds.SetValue("Inefficient Search Results Threshold",$ThresholdInefficient,[Microsoft.Win32.RegistryValueKind]::DWord) 
                        }else{ Write-Warning "$Server - Inefficient Threshold already set to $ThresholdInefficient"}

                        if (!($regValue_ExpensiveThreshold -eq $ThresholdExpensive)){
                        Write-Warning "$Server - Expensive Search Results Threshold   : $regValue_ExpensiveThreshold -> $ThresholdExpensive"
                        $regKey_Thresholds.SetValue("Expensive Search Results Threshold",$ThresholdExpensive,[Microsoft.Win32.RegistryValueKind]::DWord) 
                        }else{ Write-Warning "$Server - Expensive Threshold already set to $ThresholdExpensive"}
                    }

                    
                }
                if ($Disable)
                    {
                        if ($regValue_InefficientThreshold)
                        {
                            Write-Warning "$Server - Inefficient Search Results Threshold : $regValue_InefficientThreshold -> Deleted"
                            $regKey_Thresholds.DeleteValue("Inefficient Search Results Threshold") 
                        }
                        if ($regValue_ExpensiveThreshold)
                        {
                            Write-Warning "$Server - Expensive Search Results Threshold : $regValue_ExpensiveThreshold -> Deleted"
                            $regKey_Thresholds.DeleteValue("Expensive Search Results Threshold") 
                        }
                    }
        }
        else 
        { 
        Write-warning "$Server unreachable" 
        $obj.Value  = "Server Unreachable"
        } 
$response += $obj
    } 
return $response
}
#-----------------------------------------------------------------------------------------------------------------------
function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory="C:\Scripts\", [switch]$NoNewFolderButton)
{
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
 
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}
#-----------------------------------------------------------------------------------------------------------------------
Function fn_CreateReport1644
{
param($Hours = $Script:ReportForLastnHours)

#if (!($Servers)){$Servers = Get-DomainController}


$date=([datetime]::Now).AddHours(-$hours)
[array]$DebugEvents = @()
Write-Host "Results will be saved  to: $outFolderPath" -ForegroundColor DarkGray
Write-Host "Reading EventLogs from $($Servers.count) domain controllers for 1644 Events for last $hours hours" -ForegroundColor DarkGray

$title   = "Select Server to get Diagnitics results"
$Message = "Select Servers where you previously enabled LDAP Quiery logging'.`r`n- Ensure you have Domain Admin Permissions. `nSelect Cancel to exit when done."

$opt0       = New-Object System.Management.Automation.Host.ChoiceDescription "&Current Server ($ENV:COMPUTERNAME)"          ,"Current Server ($ENV:COMPUTERNAME)"
$opt1       = New-Object System.Management.Automation.Host.ChoiceDescription "&All DC's"                                    ,"All DC's"
$opt2       = New-Object System.Management.Automation.Host.ChoiceDescription "&Select from DC's list"                       ,"Select from DC's list"
$Cancel     = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($opt0,$opt1,$opt2,$Cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 0)

[string[]]$Servers = 
switch($result)
{
    0{ $ENV:COMPUTERNAME}  
    1{ Get-DomainController}
    2{ Get-DomainController | Out-GridView -Title "Select Server(s) from Domain Controllers List" -OutputMode Multiple }
    3{Exit}       
}
$Servers
$outFolderPath   = Read-FolderBrowserDialog -Message "Select folder for saving LDAP Query Logging Report"
$outFileFullName = $outFolderPath+"\NTDS-$DiagnosticsName-"+$(Get-Date -f "dd.MM.yyyy_HH.mm")+'.csv'
Write-Host "Results will be saved  to: $outFileFullName " -ForegroundColor DarkGray
[array]$DebugEvents = @()
$Servers | % {Write-Host "Processing: " $_ -ForegroundColor DarkGreen ; $DebugEvents += Get-WinEvent -FilterHashtable @{Logname="Directory Service";StartTime=$date;Id=1644;} -ComputerName $_  -ErrorAction SilentlyContinue}
[array]$Report = @()
     
		If ($DebugEvents -ne $null)
		{
			$outFileFullName = $outFolderPath+"\NTDS-$($DiagnosticsName -replace "\s","_")-"+$(Get-Date -f "dd.MM.yyyy_HH.mm")+'.csv'
		    Write-Host ("Event 1644 found, generating", $outFileFullName)
				
			ForEach ($mcEvent in $DebugEvents)
			{
            $objSID = New-Object System.Security.Principal.SecurityIdentifier ($mcEvent.UserId.Value)
            $objUser = $objSID.Translate([System.Security.Principal.NTAccount])

                $mc1644 = New-Object System.Object
				$mc1644 | Add-Member -MemberType NoteProperty -Name LDAPServer				            -force -Value $mcEvent.MachineName
				$mc1644 | Add-Member -MemberType NoteProperty -Name TimeGenerated	              		-force -Value $mcEvent.TimeCreated
				$mc1644 | Add-Member -MemberType NoteProperty -Name UserName	              	     	-force -Value $objUser.Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name ClientIP 				            -force -Value (mcConvertClient($mcEvent.Properties[4].Value,'IP'))
                $mc1644 | Add-Member -MemberType NoteProperty -Name ClientName 		                    -force -Value $(try{(([System.Net.Dns]::GetHostByAddress($mc1644.ClientIP)).HostName)}catch{$_})
				$mc1644 | Add-Member -MemberType NoteProperty -Name ClientPort 			               	-force -Value (mcConvertClient($mcEvent.Properties[4].Value,'Port'))
				$mc1644 | Add-Member -MemberType NoteProperty -Name StartingNode		              	-force -Value $mcEvent.Properties[0].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name Filter				                -force -Value $mcEvent.Properties[1].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name SearchScope 			            -force -Value $mcEvent.Properties[5].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name AttributeSelection 		         	-force -Value $mcEvent.Properties[6].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name ServerControls			            -force -Value $mcEvent.Properties[7].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name VisitedEntries 			            -force -Value $mcEvent.Properties[2].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name ReturnedEntries 		         	-force -Value $mcEvent.Properties[3].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name UsedIndexes				            -force -Value $mcEvent.Properties[8].Value # KB 2800945 or later has extra data fields.
				$mc1644 | Add-Member -MemberType NoteProperty -Name PagesReferenced			            -force -Value $mcEvent.Properties[9].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name PagesReadFromDisk 		         	-force -Value $mcEvent.Properties[10].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name PagesPreReadFromDisk		        -force -Value $mcEvent.Properties[11].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name CleanPagesModified       			-force -Value $mcEvent.Properties[12].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name DirtyPagesModified		         	-force -Value $mcEvent.Properties[13].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name SearchTimeMS			            -force -Value $mcEvent.Properties[14].Value
				$mc1644 | Add-Member -MemberType NoteProperty -Name AttributesPreventingOptimization	-force -Value $mcEvent.Properties[15].Value
            $Report += $mc1644
	}
		
}
		else {
			Write-Host ('  No events with EventID = 1644 found.')
		}
$Report  | export-csv -path $outFileFullName -NoTypeInformation  -Delimiter ";"  -enc UTF8
}
#-----------------------------------------------------------------------------------------------------------------------

$title   = "Select Debug Action"
$Message = "Select Server where to Enable or Disable LDAP Quiery logging'.`r`n- Ensure you have Domain Admin Permissions. `nSelect Cancel to exit when done."

$select_check       = New-Object System.Management.Automation.Host.ChoiceDescription "Check Current &Status"          ,"Check Current Status"
$select_ON          = New-Object System.Management.Automation.Host.ChoiceDescription "&Enable LDAP Query Logging"     ,"Enable LDAP Query Logging"
$select_OFF         = New-Object System.Management.Automation.Host.ChoiceDescription "&Disable LDAP Query Logging"    ,"Disable LDAP Query Logging"
$select_Report      = New-Object System.Management.Automation.Host.ChoiceDescription "Create &Report"                 ,"Create Report."
$Cancel             = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($select_check,$select_ON,$select_OFF,$select_Report,$Cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 0)
switch($result)
{
    0{ ConfigureNtdsDiagnostics -CheckOnly |ft -AutoSize }  
    1{ 
            $title   = "Select Where to Enable LDAP Query Logging"
            $Message = "Select Server where to Enable LDAP Quiery logging'.`r`n- Ensure you have Domain Admin Permissions. `nSelect Cancel to exit when done."
            $opt0       = New-Object System.Management.Automation.Host.ChoiceDescription "Current &Server ($ENV:COMPUTERNAME)"          ,"Current Server ($ENV:COMPUTERNAME)"
            $opt1       = New-Object System.Management.Automation.Host.ChoiceDescription "&All DC's"                                    ,"All DC's"
            $opt2       = New-Object System.Management.Automation.Host.ChoiceDescription "Select from &DC's list"                       ,"Select from DC's list"
            $Cancel     = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($opt0,$opt1,$opt2,$Cancel)
            $result = $host.ui.PromptForChoice($title, $message, $options, 0)
            [string[]]$Servers = 
            switch($result)
            {
                0{ ConfigureNtdsDiagnostics -Servers $ENV:COMPUTERNAME -Enable  |ft -AutoSize }  
                1{ ConfigureNtdsDiagnostics -Enable |ft -AutoSize }
                2{ $Servers = Get-DomainController | Out-GridView -Title "Select Server(s) from Domain Controllers List" -OutputMode Multiple
                   ConfigureNtdsDiagnostics -Enable -Servers $Servers |ft -AutoSize }
                3{Exit}       
            }
    }
    2{
            $title   = "Select Where to Disable LDAP Query Logging"
            $Message = "Select Server where to Disable LDAP Quiery logging'.`r`n- Ensure you have Domain Admin Permissions. `nSelect Cancel to exit when done."
            $opt0       = New-Object System.Management.Automation.Host.ChoiceDescription "Current &Server ($ENV:COMPUTERNAME)"          ,"Current Server ($ENV:COMPUTERNAME)"
            $opt1       = New-Object System.Management.Automation.Host.ChoiceDescription "&All DC's"                                    ,"All DC's"
            $opt2       = New-Object System.Management.Automation.Host.ChoiceDescription "Select from &DC's list"                       ,"Select from DC's list"
            $Cancel     = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Cancel."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($opt0,$opt1,$opt2,$Cancel)
            $result = $host.ui.PromptForChoice($title, $message, $options, 0)
            [string[]]$Servers = 
            switch($result)
            {
                0{ ConfigureNtdsDiagnostics -Servers $ENV:COMPUTERNAME -Disable  |ft -AutoSize }  
                1{ ConfigureNtdsDiagnostics -Disable |ft -AutoSize }
                2{ $Servers = Get-DomainController | Out-GridView -Title "Select Server(s) from Domain Controllers List" -OutputMode Multiple
                   ConfigureNtdsDiagnostics -Disable -Servers $Servers |ft -AutoSize }
                3{Exit}       
            }
      }
    3{ fn_CreateReport1644 }
    4{Exit}       
}
#
# END
#-----------------------------------------------------------------------------------------------------------------------