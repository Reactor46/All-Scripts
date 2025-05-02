##################################### 
#            Functions                # 
##################################### 
 
function WMIDateStringToDate($Bootup) {    
    [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootup)    
} 
 
function Test-Port{    
[cmdletbinding(    
    DefaultParameterSetName = '',    
    ConfirmImpact = 'low'    
)]    
    Param(    
        [Parameter(    
            Mandatory = $True,    
            Position = 0,    
            ParameterSetName = '',    
            ValueFromPipeline = $True)]    
            [array]$computer,    
        [Parameter(    
            Position = 1,    
            Mandatory = $True,    
            ParameterSetName = '')]    
            [array]$port,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$TCPtimeout=1000,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$UDPtimeout=1000,               
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$TCP,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$UDP                                      
        )    
    Begin {    
        If (!$tcp -AND !$udp) {$tcp = $True}    
        #Typically you never do this, but in this case I felt it was for the benefit of the function    
        #as any errors will be noted in the output of the report            
        $ErrorActionPreference = "SilentlyContinue"    
        $report = @()    
    }    
    Process {       
        ForEach ($c in $computer) {    
            ForEach ($p in $port) {    
                If ($tcp) {      
                    #Create temporary holder     
                    $temp = "" | Select-Object Server, Port, TypePort, Open, Notes    
                    #Create object for connecting to port on computer    
                    $tcpobject = new-Object system.Net.Sockets.TcpClient    
                    #Connect to remote machine's port                  
                    $connect = $tcpobject.BeginConnect($c,$p,$null,$null)    
                    #Configure a timeout before quitting    
                    $wait = $connect.AsyncWaitHandle.WaitOne($TCPtimeout,$false)    
                    #If timeout    
                    If(!$wait) {    
                        #Close connection    
                        $tcpobject.Close()    
                        Write-Verbose "Connection Timeout"    
                        #Build report    
                        $temp.Server = $c    
                        $temp.Port = $p    
                        $temp.TypePort = "TCP"    
                        $temp.Open = "False"    
                        $temp.Notes = "Connection to Port Timed Out"    
                    } Else {    
                        $error.Clear()    
                        $tcpobject.EndConnect($connect) | out-Null    
                        #If error    
                        If($error[0]){    
                            #Begin making error more readable in report    
                            [string]$string = ($error[0].exception).message    
                            $message = (($string.split(":")[1]).replace('"',"")).TrimStart()    
                            $failed = $true    
                        }    
                        #Close connection        
                        $tcpobject.Close()    
                        #If unable to query port to due failure    
                        If($failed){    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "False"    
                            $temp.Notes = "$message"    
                        } Else{    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "True"      
                            $temp.Notes = ""    
                        }    
                    }       
                    #Reset failed value    
                    $failed = $Null        
                    #Merge temp array with report                
                    $report += $temp    
                }        
                If ($udp) {    
                    #Create temporary holder     
                    $temp = "" | Select-Object Server, Port, TypePort, Open, Notes                                       
                    #Create object for connecting to port on computer    
                    $udpobject = new-Object system.Net.Sockets.Udpclient  
                    #Set a timeout on receiving message   
                    $udpobject.client.ReceiveTimeout = $UDPTimeout   
                    #Connect to remote machine's port                  
                    Write-Verbose "Making UDP connection to remote server"   
                    $udpobject.Connect("$c",$p)   
                    #Sends a message to the host to which you have connected.   
                    Write-Verbose "Sending message to remote host"   
                    $a = new-object system.text.asciiencoding   
                    $byte = $a.GetBytes("$(Get-Date)")   
                    [void]$udpobject.Send($byte,$byte.length)   
                    #IPEndPoint object will allow us to read datagrams sent from any source.    
                    Write-Verbose "Creating remote endpoint"   
                    $remoteendpoint = New-Object system.net.ipendpoint([system.net.ipaddress]::Any,0)   
                    Try {   
                        #Blocks until a message returns on this socket from a remote host.   
                        Write-Verbose "Waiting for message return"   
                        $receivebytes = $udpobject.Receive([ref]$remoteendpoint)   
                        [string]$returndata = $a.GetString($receivebytes)  
                        If ($returndata) {  
                           Write-Verbose "Connection Successful"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "True"    
                            $temp.Notes = $returndata     
                            $udpobject.close()     
                        }                         
                    } Catch {   
                        If ($Error[0].ToString() -match "\bRespond after a period of time\b") {   
                            #Close connection    
                            $udpobject.Close()    
                            #Make sure that the host is online and not a false positive that it is open   
                            If (Test-Connection -comp $c -count 1 -quiet) {   
                                Write-Verbose "Connection Open"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "True"    
                                $temp.Notes = ""   
                            } Else {   
                                <#   
                                It is possible that the host is not online or that the host is online,    
                                but ICMP is blocked by a firewall and this port is actually open.   
                                #>   
                                Write-Verbose "Host maybe unavailable"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "False"    
                                $temp.Notes = "Unable to verify if port is open or if host is unavailable."                                   
                            }                           
                        } ElseIf ($Error[0].ToString() -match "forcibly closed by the remote host" ) {   
                            #Close connection    
                            $udpobject.Close()    
                            Write-Verbose "Connection Timeout"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "False"    
                            $temp.Notes = "Connection to Port Timed Out"                           
                        } Else {                        
                            $udpobject.close()   
                        }   
                    }       
                    #Merge temp array with report                
                    $report += $temp    
                }                                    
            }    
        }                    
    }    
    End {    
        #Generate Report    
        $ADHealthReport 
    }  
} 
 

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
$osType=Get-WmiObject Win32_OperatingSystem -ComputerName $computername| Select-Object -ExpandProperty OSArchitecture 
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

function CheckTime {

$tstats=@()

$timer = [diagnostics.stopwatch]::startnew()

foreach ($server in $npt){
$wdt = (Get-WmiObject -ComputerName $server -Query "select LocalDateTime from win32_operatingsystem").LocalDateTime
$dt = ([wmi]'').ConvertToDateTime($wdt) - $timer.elapsed

$tstat = "" |Select-Object Server,Timestamp
$tstat.server = $server
$tstat.timestamp = $dt
$tstats += $tstat
}

$enddate = (Get-Date).tostring("MM-dd-yyyy")
$tstats
}
# End function CheckTime

# Begin Disk Cleanup
function CleanDisk{

# Name: RemoteDiskCleanup.ps1                              
# Creator: Myrianthi                
# CreationDate: 11.26.2018                             
# LastModified: 2.1.2019                               
# Version: 2.3
# Doc: https://github.com/Myrianthi/remotediskcleanup
# Purpose: Remote-access temp file removal
# Requirements: Admin access, PS-Remoting enabled on remote devices
#

# --------------------------- Script begins here --------------------------- #

#Requires -RunAsAdministrator
Set-ExecutionPolicy RemoteSigned

# This will check all Disk Cleanup boxes by manually setting each key in the following registry path to 2.
# Comment out the files that you do not want Disk Cleanup to erase.
$SageSet = "StateFlags0099"
$StateFlags= "Stateflags0099"
$Base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\"
$VolCaches = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"

$Locations = @(
    "Active Setup Temp Folders"
    "BranchCache"
    "Downloaded Program Files"
    "GameNewsFiles"
    "GameStatisticsFiles"
    "GameUpdateFiles"
    "Internet Cache Files"
    "Memory Dump Files"
    "Offline Pages Files"
    "Old ChkDsk Files"
    "Previous Installations"

    # This is commented out because we already call a function in this script to wipe Recycle Bin contents older than 3 days.
    #"Recycle Bin"

    "Service Pack Cleanup"
    "Setup Log Files"
    "System error memory dump files"
    "System error minidump files"
    "Temporary Files"
    "Temporary Setup Files"
    "Temporary Sync Files"
    "Thumbnail Cache"
    "Update Cleanup"
    "Upgrade Discarded Files"
    "User file versions"
    "Windows Defender"
    "Windows Error Reporting Archive Files"
    "Windows Error Reporting Queue Files"
    "Windows Error Reporting System Archive Files"
    "Windows Error Reporting System Queue Files"
    "Windows ESD installation files"
    "Windows Upgrade Log Files"
)

Function Get-Recyclebin{
    [CmdletBinding()]
    Param
    (
        $ComputerOBJ,
        $RetentionTime = "3"
    )
    Write-Host "Empyting Recycle bin items older than $RetentionTime days" -ForegroundColor Yellow
    If($ComputerOBJ.PSRemoting -eq $true){
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock {
        
        Try{
            $Shell = New-Object -ComObject Shell.Application
            $Recycler = $Shell.NameSpace(0xa)
            $Recycler.Items() 

            foreach($item in $Recycler.Items())
            {
                $DeletedDate = $Recycler.GetDetailsOf($item,2) -replace "\u200f|\u200e","" #Invisible Unicode Characters
                $DeletedDatetime = Get-Date $DeletedDate 
                [Int]$DeletedDays = (New-TimeSpan -Start $DeletedDatetime -End $(Get-Date)).Days

                If($DeletedDays -ge $RetentionTime)
                {
                    Remove-Item -Path $item.Path -Confirm:$false -Force -Recurse
                }
            }
        }
        Catch [System.Exception]{
            $RecyclerError = $true
        }
        Finally{
            If($RecyclerError -eq $False){
                Write-output $True 
            }
            Else{
                Write-Output $False
            }
        }

        
    } -Credential $ComputerOBJ.Credential
        If($Result -eq $True){
            Write-Host "All recycle bin items older than $RetentionTime days were deleted" -ForegroundColor Green
        }
        Else{
            Write-Host "Unable to delete some items in the Recycle Bin." -ForegroundColor Red
        }
    }
    Else{
        Try{
            $Shell = New-Object -ComObject Shell.Application
            $Recycler = $Shell.NameSpace(0xa)
            $Recycler.Items() 

            foreach($item in $Recycler.Items())
            {
                $DeletedDate = $Recycler.GetDetailsOf($item,2) -replace "\u200f|\u200e","" #Invisible Unicode Characters
                $DeletedDatetime = Get-Date $DeletedDate 
                [Int]$DeletedDays = (New-TimeSpan -Start $DeletedDatetime -End $(Get-Date)).Days

                If($DeletedDays -ge $RetentionTime)
                {
                    Remove-Item -Path $item.Path -Confirm:$false -Force -Recurse
                }
            }
        }
        Catch [System.Exception]{
            $RecyclerError = $true
        }
        Finally{
            If($RecyclerError -eq $true){
                Write-Host "Unable to delete some items in the Recycle Bin." -ForegroundColor Red
            }
            Else{
                Write-Host "All recycle bin items older than $RetentionTime days were deleted" -ForegroundColor Green
            }
        }
    }    
}

Function Clean-Path{

    Param
    (
        [String]$Path,
        $ComputerOBJ
    )
    Write-Host "`t...Cleaning $Path"
    If($ComputerOBJ.PSRemoting -eq $True){

        Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock {

            If(Test-Path $Using:Path){

                Foreach($Item in $(Get-ChildItem -Path $Using:Path -Recurse)){
    
                    Try{
                        Remove-item -Path $item.FullName -Confirm:$False -Recurse -ErrorAction Stop
                    }
                    Catch [System.Exception]{
                        Write-verbose "$($Item.path) - $($_.Exception.Message)"
                    }
                }
            }

        } -Credential $ComputerOBJ.Credential
    }
    Else{

        If(Test-Path $Path){
        
        Foreach($Item in $(Get-ChildItem -Path $Path -Recurse)){
    
            Try{
                Remove-item -Path $item.FullName -Confirm:$False -Recurse -ErrorAction Stop
            }
            Catch [System.Exception]{
                Write-verbose "$($Item.path) - $($_.Exception.Message)"
            }
        }
    }



    }
}

Function Get-OrigFreeSpace{

    Param
    (
        $ComputerOBJ
    )

    Try{
        $RawFreespace = (Get-WmiObject Win32_logicaldisk -ComputerName $ComputerOBJ.ComputerName -Credential $ComputerOBJ.Credential -ErrorAction Stop | Where-Object {$_.DeviceID -eq 'C:'}).freespace
        $FreeSpaceGB = [decimal]("{0:N2}" -f($RawFreespace/1gb))
        Write-host "Current Free Space on the OS Drive : $FreeSpaceGB GB" -ForegroundColor Magenta
    }
    Catch [System.Exception]{
        $FreeSpaceGB = $False
        Write-Host "Unable to pull free space from OS drive. Press enter to Exit..." -ForegroundColor Red    
    }
    Finally{
        $ComputerOBJ | Add-Member -MemberType NoteProperty -Name OrigFreeSpace -Value $FreeSpaceGB
        Write-output $ComputerOBJ
    }
}

Function Get-FinalFreeSpace{

    Param
    (
        $ComputerOBJ
    )

    Try{
        $RawFreespace = (Get-WmiObject Win32_logicaldisk -ComputerName $ComputerOBJ.ComputerName -Credential $ComputerOBJ.Credential -ErrorAction Stop | Where-Object {$_.DeviceID -eq 'C:'}).freespace
        $FreeSpaceGB = [decimal]("{0:N2}" -f($RawFreespace/1gb))
        Write-host "Final Free Space on the OS Drive : $FreeSpaceGB GB" -ForegroundColor Magenta
    }

    Catch [System.Exception]{
        $FreeSpaceGB = $False
        Write-Host "Unable to pull free space from OS drive. Press enter to Exit..." -ForegroundColor Red    
    }
    Finally{
        $ComputerOBJ | Add-Member -MemberType NoteProperty -Name FinalFreeSpace -Value $FreeSpaceGB
        Write-output $ComputerOBJ
    }
} 

Function Get-Computername {

    Write-Host "Please enter the computername to connect to or just hit enter for localhost" -ForegroundColor Yellow
    $ComputerName = Read-Host

    if($ComputerName -eq '' -or $ComputerName -eq $null){
        $obj = New-object PSObject -Property @{
            ComputerName = $env:COMPUTERNAME
            Remote = $False
        }
    }
    else{
        $obj = New-object PSObject -Property @{
            ComputerName = $Computername
            Remote = $True
        }
    }

    Write-output $obj

}

Function Test-PSRemoting{

    Param
    (
        $ComputerOBJ
    )

    Write-Host "Please enter your credentials for the remote machine." -ForegroundColor Yellow
    $ComputerOBJ | Add-Member NoteProperty -Name Credential -Value (Get-Credential)

    $RemoteHostname = Invoke-command -ComputerName $ComputerOBJ.Computername -ScriptBlock {hostname} -Credential $ComputerOBJ.Credential -erroraction 'silentlycontinue'

    If($RemoteHostname -eq $ComputerOBJ.Computername){
        Write-Host "PowerShell Remoting was successful" -ForegroundColor Green
        $ComputerOBJ | Add-Member NoteProperty -Name PSRemoting -Value $True
    }
    Else {
        Write-host "PowerShell Remoting FAILED press enter to exit script." -ForegroundColor Red
        $ComputerOBJ | Add-Member NoteProperty -Name PSRemoting -Value $False
    }

    Write-output $ComputerOBJ
}

Function Run-CleanMGR{

    Param
    (
        $ComputerOBJ
    )

    If($ComputerOBJ.PSRemoting -eq $true){
        Write-Host "Attempting to Run Windows Disk Cleanup With Parameters..." -ForegroundColor Yellow
        Write-Host "`tApplying $sageset parameters to registry path:"
        Write-Host "`t$Base"
        $CleanMGR = Invoke-command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock {
                        $ErrorActionPreference = 'Stop'
                        Try{

                            # Set Sageset99/Stateflag0099 reg keys to 1 for a blank slate.
                            foreach($VC in $VolCaches){
                                New-ItemProperty -Path "$($VC.PSPath)" -Name $StateFlags -Value 1 -Type DWORD -Force | Out-Null
                            } 

                            ForEach($Location in $Locations) {
                                Set-ItemProperty -Path $($Base+$Location) -Name $SageSet -Type DWORD -Value 2 -ea silentlycontinue | Out-Null
                            }
                                                       
                            # Convert the Sageset number previously defined and Run the Disk Cleanup process with configured parameters.
                            $Args = "/sagerun:$([string]([int]$SageSet.Substring($SageSet.Length-4)))"
                            Start-Process -Wait "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList $Args -WindowStyle Hidden
                            
                            # Set Sageset99/Stateflag0099 reg keys back to 1 for a blank slate.
                            foreach($VC in $VolCaches){
                            New-ItemProperty -Path "$($VC.PSPath)" -Name $StateFlags -Value 1 -Type DWORD -Force | Out-Null
                            } 
                            $ErrorActionPreference = 'SilentlyContinue'
                            #Write-Output $true
                        }
                        Catch [System.Exception]{
                            $ErrorActionPreference = 'SilentlyContinue'
                            #Write-output $False
                        }
                    } -Credential $ComputerOBJ.Credential

        If($CleanMGR -eq $True){
            ForEach($Location in $Locations) {
                Write-Host "`t...Cleaning $Location"
            }
            Write-Host "Windows Disk Cleanup has been run successfully." -ForegroundColor Green
        }
        Else{
            Write-host "Cleanmgr is not installed! To use this portion of the script you must install the following windows features:" -ForegroundColor Red
            Write-host "Desktop-Experience, Ink-Handwriting" -ForegroundColor Red
        }
    }
    Else{

        Write-Host "Attempting to Run Windows Disk Cleanup With Parameters..." -ForegroundColor Yellow
        Write-Host "`tApplying $sageset parameters to registry path:"
        Write-Host "`t$Base"
        Echo ""
        $ErrorActionPreference = 'Stop'
        Try{

            # Set Sageset99/Stateflag0099 reg keys to 1 for a blank slate.
            foreach($VC in $VolCaches){
                New-ItemProperty -Path "$($VC.PSPath)" -Name $StateFlags -Value 1 -Type DWORD -Force | Out-Null
            } 

            ForEach($Location in $Locations) {
                Set-ItemProperty -Path $($Base+$Location) -Name $SageSet -Type DWORD -Value 2 -ea silentlycontinue | Out-Null
            }

            # Convert the Sageset number previously defined and Run the Disk Cleanup process with configured parameters.
            $Args = "/sagerun:$([string]([int]$SageSet.Substring($SageSet.Length-4)))"
            Start-Process -Wait "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList $Args -WindowStyle Hidden

            # Set Sageset99/Stateflag0099 reg keys back to 1 for a blank slate.
            foreach($VC in $VolCaches){
                New-ItemProperty -Path "$($VC.PSPath)" -Name $StateFlags -Value 1 -Type DWORD -Force | Out-Null
            } 
            $ErrorActionPreference = 'SilentlyContinue'
            #Write-Output $true 
            ForEach($Location in $Locations) {
                Write-Host "`t...Cleaning $Location"
            }
            Write-Host "Windows Disk Cleanup has been run successfully." -ForegroundColor Green
        }
        Catch [System.Exception]{
          Write-host "Cleanmgr is not installed! To use this portion of the script you must install the following windows features:" -ForegroundColor Red
          Write-host "Desktop-Experience, Ink-Handwriting" -ForegroundColor Red

        }
        $ErrorActionPreference = 'SilentlyContinue'
    }
}

Function Erase-IExplorerHistory{

    Param
    (
        $ComputerOBJ
    )

    If($ComputerOBJ.PSRemoting -eq $true){
        Write-Host "Attempting to Erase Internet Explorer temp data" -ForegroundColor Yellow
        $CleanIExplorer = Invoke-command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock {
                        $ErrorActionPreference = 'Stop'
                        Try{
                            Start-Process -FilePath rundll32.exe -ArgumentList 'inetcpl.cpl,ClearMyTracksByProcess 4351' -Wait -NoNewWindow
                            $ErrorActionPreference = 'SilentlyContinue'
                            #Write-Output $true
                        }
                        Catch [System.Exception]{
                            $ErrorActionPreference = 'SilentlyContinue'
                            #Write-output $False
                        }
                    } -Credential $ComputerOBJ.Credential

        If($CleanIExplorer -eq $True){
            Write-Host "Internet Explorer temp data has been successfully erased" -ForegroundColor Green
        }
        Else{
            Write-host "Failed to erase Internet Explorer temp data" -ForegroundColor Red
        }
    }
    Else{

        Write-Host "Attempting to Erase Internet Explorer temp data" -ForegroundColor Yellow
        $ErrorActionPreference = 'Stop'
        Try{
            Start-Process -FilePath rundll32.exe -ArgumentList 'inetcpl.cpl,ClearMyTracksByProcess 4351' -Wait -NoNewWindow
            Write-Host "Internet Explorer temp data has been successfully erased" -ForegroundColor Green
        }
        Catch [System.Exception]{
          Write-host "Failed to erase Internet Explorer temp data" -ForegroundColor Red
        }
        $ErrorActionPreference = 'SilentlyContinue'
    }
}


# Windows computer cleanup tool


Clear-Host

Echo "  **This tool will attempt to remove bloatware and erase temp files across all user profiles. Please use with caution.**"
Echo ""

$ComputerOBJ = Get-ComputerName

Echo "You have entered $ComputerOBJ. Is this correct?"
Pause
Echo ""

Start-Transcript -Path C:\Windows\System32\CleanupLogs\$ComputerOBJ.txt
[System.DateTime]::Now
Echo ""

Echo "********************************************************************************************************************"
If($ComputerOBJ.Remote -eq $true){
    $ComputerOBJ = Test-PSRemoting -ComputerOBJ $ComputerOBJ
    If($ComputerOBJ.PSRemoting -eq $False){
        Read-Host
        exit;
    }
}

$ComputerOBJ = Get-OrigFreeSpace -ComputerOBJ $ComputerOBJ

If($ComputerOBJ.OrigFreeSpace -eq $False){
    Read-host
    exit;
}
Echo "********************************************************************************************************************"
Echo ""

#======================================================================================================================================================


Write-Host "Cleaning temp directories across all user profiles" -ForegroundColor Yellow

Clean-path -Path 'C:\Temp\*' -Verbose -ComputerOBJ $ComputerOBJ
Clean-path -Path 'C:\Windows\Temp\*' -Verbose -ComputerOBJ $ComputerOBJ
Clean-Path -Path 'C:\Users\*\Documents\*tmp' -Verbose -ComputerOBJ $ComputerOBJ
Clean-path -Path 'C:\Documents and Settings\*\Local Settings\Temp\*' -ComputerOBJ $ComputerOBJ
Clean-path -Path 'C:\Users\*\Appdata\Local\Temp\*' -Verbose -ComputerOBJ $ComputerOBJ
Clean-path -Path 'C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*' -Verbose -ComputerOBJ $ComputerOBJ
Clean-path -Path 'C:\Users\*\AppData\Roaming\Microsoft\Windows\Cookies\*' -Verbose -ComputerOBJ $ComputerOBJ

#Clean-path -Path 'C:\ServiceProfiles\LocalService\AppData\Local\Temp\*' -Verbose -ComputerOBJ $ComputerOBJ

#####Only turned off because I don't want to hear complaints of programs taking too long to start#####
#Clean-path -Path 'C:\Windows\Prefetch' -Verbose -ComputerOBJ $ComputerOBJ

#####Internet Explorer Cache. Turned off Because I now have a function to clean it and therefore unneccesary to perform a hard reset#####
#Clean-path -Path 'C:\Users\*\AppData\Local\Microsoft\Windows\INetCache'-Verbose -ComputerOBJ $ComputerOBJ

#####Figured I would keep these because they don't take up too much space and some users might find their recent files convenient#####
#Clean-Path -Path 'C:\Users\*\AppData\Roaming\Microsoft\Windows\Recent' -Verbose -ComputerOBJ $ComputerOBJ
#Clean-Path -Path 'C:\AppData\Roaming\Microsoft\Windows\Recent' -Verbose -ComputerOBJ $ComputerOBJ

#####Some reports of this messing up Chrome by forcing a hard reset of its cache. It apparently still tries to read from cache when it's been manually cleared#####
#Clean-Path -Path 'C:\Users\*\AppData\Local\Google\Chrome\User Data\Default\Cache' -Verbose -ComputerOBJ $ComputerOBJ

#####Completely wiping Mozilla Firefoxes Cache. Hard reset not tested yet...#####
#Clean-Path -Path 'C:\Users\*\AppData\Local\Mozilla\Firefox\Profiles\*.default' -Verbose -ComputerOBJ $ComputerOBJ
#Clean-Path -Path 'C:\Users\*\AppData\Roaming\Mozilla\Firefox\Profiles\*.default' -Verbose -ComputerOBJ $ComputerOBJ

#####Error reporting and Debug information. Might come in handy to just keep this#####
#Clean-path -Path 'C:\ProgramData\Microsoft\Windows\WER\ReportArchive' -Verbose -ComputerOBJ $ComputerOBJ
#Clean-path -Path 'C:\ProgramData\Microsoft\Windows\WER\ReportQueue' -Verbose -ComputerOBJ $ComputerOBJ

Write-Host "All Temp Paths have been cleaned" -ForegroundColor Green
Echo ""

#======================================================================================================================================================

Run-CleanMGR -ComputerOBJ $ComputerOBJ
Echo ""
Erase-IExplorerHistory -ComputerOBJ $ComputerOBJ
Echo ""
Get-Recyclebin -ComputerOBJ $ComputerOBJ
Echo ""

# ADDING THIS SOON #
#Wipe-Freespace


Echo "********************************************************************************************************************"
$ComputerOBJ = Get-FinalFreeSpace -ComputerOBJ $ComputerOBJ
$SpaceRecovered = $($Computerobj.finalfreespace) - $($ComputerOBJ.OrigFreeSpace)

If($SpaceRecovered -lt 0){
    Write-Host "Less than a Gigabyte of Free Space was Recovered." -ForegroundColor Magenta
}
ElseIf($SpaceRecovered -eq 0){
    Write-host "No Space was Recovered" -ForegroundColor Magenta
}
Else{

    Write-host "Free Space Recovered : $SpaceRecovered GB" -ForegroundColor Magenta
}
Echo "********************************************************************************************************************"

Echo ""
Stop-Transcript

# --------------------------- Script ends here --------------------------- #

}

# Begin Get-DriveSpaceReport

function Get-DriveSpaceReport {
<#
.SYNOPSIS
    Gets drive space on specified computers using PSRemoting.
.DESCRIPTION
    Gets drive space information using PSRemoting or straight
    WMI calls, depending on specified switches.
.PARAMETER ComputerName
    Computer name to run the function against.
.PARAMETER DriveType
    Indicates the drive type to be analyzed. Defaults to 3, which
    is Local Disk. All drive types from the Win32_LogicalDisk class
    are valid.
.PARAMETER V2
    Uses RPC instead of WS-MAN to be PowerShell V2 compatible.
    Under the hood, it's using Get-WmiObject versus Get-CimInstance.
.PARAMETER NoSql
    Tells the module not to dump the data to the SQL database.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost
    Simple drive space check using only computer name.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost -DriveType 2
    Check drive space using computer name and specifying a DriveType of 2.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName dc2 -V2
    Check using PowerShell V2 compatibility.
.NOTES
    Version                 :  0.4
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>

[CmdletBinding()]
param(
    [parameter(mandatory=$true)][string[]]$ComputerName,
    [int]$DriveType=3,
    [switch]$V2=$false,
    [switch]$NoSql=$false,
    [string]$SqlConnectionString
)

Begin {

    Write-Verbose "Begin Block"
    
    # Simply enumerating the computers to run against. May remove this loop in the future.
    foreach ($c in $ComputerName) {
        Write-Verbose "Function will run against this computer  :  $c"
    } # End foreach loop
    
} # End Begin block

Process {
    # Enumerate each computer in $ComputerName and get the required info.
    Write-Verbose "Process Block"
    foreach ($computer in $ComputerName) {
        Write-Verbose "Processing $computer"
        try{
            if ( $V2 -eq $True ) {
                Write-Verbose "Using Get-WmiObject calls; -V2 switch was used."
                $os = Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem -ErrorAction Stop -ErrorVariable OSError
                $disk = Get-WmiObject -ComputerName $computer -Class Win32_LogicalDisk -Filter "drivetype=$DriveType" -ErrorAction Stop -ErrorVariable DiskError
            }
            else {
                Write-Verbose "Using Get-CimInstance via PSRemoting; -V2 switch was NOT used."
                $os = Invoke-Command -ComputerName $computer -ScriptBlock { Get-CimInstance -ClassName Win32_OperatingSystem } -ErrorAction Stop -ErrorVariable OSError
                $disk = Invoke-Command -ComputerName $computer -ScriptBlock { param($dt) Get-CimInstance -ClassName Win32_LogicalDisk -Filter "drivetype=$dt" } -ArgumentList $DriveType -ErrorAction Stop -ErrorVariable DiskError
            }
       
        # Enumerate each drive in $disk. Specifically, this allows for the details on each drive if a computer has more than one.
        foreach ($drive in $disk){
            Write-Verbose "Processing $drive"
            $prop = @{
                'ComputerName' = $computer
                'Drive' = $drive.DeviceID
                'PctFree' = $drive.FreeSpace / $drive.Size
                'Free' = $drive.FreeSpace
                'Size' = $drive.Size
                'OSName' = $os.Caption
                'Date' = (Get-Date)
            }
            $object = New-Object -TypeName PSObject -Property $prop
            $object.PSObject.TypeNames.Insert(0,'Report.DriveSpaceInfo')
            if (-not $NoSql) { Write-Output $object  | Save-ReportData -ConnectionString $SqlConnectionString }
            Write-Output $object
        }
        }
     catch{
            Write-Warning "You done screwed up.  $computer is no con permiso."
     } # End Catch block
    }

} # End Process Block

End { Write-Verbose "End Block" } # End End Block

}
# End Get-DriveSpaceReport
 

# Begin Send-DriveSpaceReport
function Send-DriveSpaceReport {
<#
.SYNOPSIS
    Creates and sends the drive space report to the specified user in HTML formatting.
.DESCRIPTION
    Takes objects generated from the Get-DriveSpaceReport function, formats
    it to HTML, and then sends it to the specified recipient.
.NOTES
    Version                 :  0.2
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true)]$Objects,
    [Parameter(Mandatory=$True)][string[]]$Recipient,
    [Parameter(Mandatory=$True)][string]$Sender,
    [Parameter(Mandatory=$True)][string]$EmailServer,
    [string]$SqlConnectionString,
    [switch]$AsAttachment = $false,
    $WarningThreshold=30,
    $CriticalThreshold=15
)



Begin{
    #Import-Module -Name C:\Scripts\Modules\SQLReporting
    Write-Verbose "Begin Block"
    Write-Verbose "Initializing object arrays"
    $NormalObjects = @()
    $WarningObjects = @()
    $CriticalObjects = @()
    if (-not $Objects) { $Objects = Get-ReportData -TypeName Report.DriveSpaceInfo -ConnectionString $SqlConnectionString }
}

Process{
    foreach ($Object in $Objects)
    {
      Write-Verbose "Processing Object: $Object"
      if($Object.Date.Date -eq (Get-Date).Date) {
        if ( $Object.PctFree -lt ($CriticalThreshold / 100) ) {
            Write-Verbose "Adding object to Critical Objects"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            Write-Verbose "Formatted object:  $FormattedObject"
            $CriticalObjects += $FormattedObject
        }
        elseif ( $Object.PctFree -lt ($WarningThreshold / 100) ) {
            Write-Verbose "Adding object to Warning Objects"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            $WarningObjects += $FormattedObject
        }
        else {
            Write-Verbose "Adding object to Normal Object"
            $FormattedObject = Select-Object -InputObject $Object -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f [System.Convert]::ToDouble($object.PctFree)}},@{n="Free";e={"{0:N2}" -f ($_.Free/1GB)}},@{n="Size";e={"{0:N2}" -f ($_.Size/1GB)}},@{n="DateCaptured";e={ $_.Date.ToString("MM/dd/yyyy")}}
            $NormalObjects += $FormattedObject
        }
      } # End foreach loop
    }
}

End{
    
    # CSS - Doesn't format well with Windows version of Outlook due to Word being used as rendering engine
    $css = '<style>
            table { width:98%; }
            td { text-align:center; padding:5px; }
            th { background-color:blue; color:white; }
            h3 { text-align:center }
            h6 { text-align:center }
            </style>'

    Write-Verbose "End Block"
    Write-Verbose "Building HTML report"

    $CriticalHTML = $CriticalObjects | ConvertTo-Html -Fragment -PreContent "<h3>CRITICAL - Less than $CriticalThreshold% free</h3>" | Out-String
    $WarningHTML = $WarningObjects | ConvertTo-Html -Fragment -PreContent "<h3>WARNING - Less than $WarningThreshold% free</h3>" | Out-String
    $NormalHTML = $NormalObjects | ConvertTo-Html -Fragment -PreContent "<h3>NORMAL - More than $WarningThreshold% free</h3>" | Out-String
    $FooterHtml = ConvertTo-Html -Fragment -PostContent "<h6>This report was run from:  $env:COMPUTERNAME on $(Get-Date)</h6>" | Out-String
    
    Write-Verbose "Sending Email:
          Recipient   : $Recipient
          Sender      : $Sender
          EmailServer : $EmailServer"

    if ($AsAttachment){
        $Report = ConvertTo-Html -Body "$CriticalHTML $WarningHTML $NormalHTML $FooterHtml $css" | Out-File $env:TMP\drivespace.html
        Write-Verbose "$Report"
        Send-MailMessage -to $Recipient -From $Sender -Subject "Drive Space Report" -Body "Please find the attached drive space report." -Attachments $env:TMP\drivespace.html -SmtpServer $EmailServer
    }
    else{
        $Report = ConvertTo-Html -Body "$CriticalHTML $WarningHTML $NormalHTML $FooterHtml $css" | Out-String
        Write-Verbose "$Report"
        Send-MailMessage -to $Recipient -From $Sender -Subject "Drive Space Report" -BodyAsHtml $Report -SmtpServer $EmailServer
    }

    
}

}
# End Send-DriveSpaceReport