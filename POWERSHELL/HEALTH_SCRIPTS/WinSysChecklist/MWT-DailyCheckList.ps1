﻿<#
Scripted WinSysCheckList
#Renamed 10/01/2021 MWT-DailyCheckList#
Contains various scripts and functions that will gather
the required information for the WinSys Daily Checklist

Written by: John Battista

V1.00, 9/18/2017 - Initial version
V1.01, 10/02/2021 - USON edits
     Renamed to MWT-DailyCheckList.ps1
#>
# Set the date the script was run to the end of each report
# Ex: $Report = '.\Report ' + $SaveDate + '.html'
## Functions for Script
## Out-File in UTF8 NonBom
function Out-FileUtf8NoBom {

  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position=0)] [string] $LiteralPath,
    [switch] $Append,
    [switch] $NoClobber,
    [AllowNull()] [int] $Width,
    [Parameter(ValueFromPipeline)] $InputObject
  )

  

  # Make sure that the .NET framework sees the same working dir. as PS
  # and resolve the input path to a full path.
  [System.IO.Directory]::SetCurrentDirectory($PWD) # Caveat: .NET Core doesn't support [Environment]::CurrentDirectory
  $LiteralPath = [IO.Path]::GetFullPath($LiteralPath)

  # If -NoClobber was specified, throw an exception if the target file already
  # exists.
  if ($NoClobber -and (Test-Path $LiteralPath)) {
    Throw [IO.IOException] "The file '$LiteralPath' already exists."
  }

  # Create a StreamWriter object.
  # Note that we take advantage of the fact that the StreamWriter class by default:
  # - uses UTF-8 encoding
  # - without a BOM.
  $sw = New-Object IO.StreamWriter $LiteralPath, $Append

  $htOutStringArgs = @{}
  if ($Width) {
    $htOutStringArgs += @{ Width = $Width }
  }

  # Note: By not using begin / process / end blocks, we're effectively running
  #       in the end block, which means that all pipeline input has already
  #       been collected in automatic variable $Input.
  #       We must use this approach, because using | Out-String individually
  #       in each iteration of a process block would format each input object
  #       with an indvidual header.
  try {
    $Input | Out-String -Stream @htOutStringArgs | % { $sw.WriteLine($_) }
  } finally {
    $sw.Dispose()
  }

}
## End Out-File in UTF8 NonBom



####******************* set email parameters ****************** ######
$from = "ADHealthCheck@optummso.com" 
$to="jbattista@optummso.com"
$smtpserver="mail.optummso.com"
####*******************######################****************** ######
$SaveDate = (Get-Date).tostring("MM-dd-yyyy")
$timeout = "60"
$AllSrvs = Get-Content -Path "$PSScriptRoot\Configs\Servers.txt"
$BatProc = Get-Content -Path "$PSScriptRoot\Configs\BatProcSrvs.txt"
$CapsMTSrvs = Get-Content -Path "$PSScriptRoot\Configs\CapsMTSrvs.txt"
$CapsSrvs = Get-Content -Path "$PSScriptRoot\Configs\CapsSrvs.txt"
$CASMTSrvs = Get-Content -Path "$PSScriptRoot\Configs\CASMTSrvs.txt"
$CASSrvs = Get-Content -Path "$PSScriptRoot\Configs\CASSrvs.txt"
$CollSrvs = Get-Content -Path "$PSScriptRoot\Configs\CollSrvs.txt"
$DCSrvs = Get-Content -Path "$PSScriptRoot\Configs\DCSrvs.txt"
$npt = $ntpServer

Get-ADComputer -Server MWTDC02 -Filter {Operatingsystem -Like 'Windows Server*' -and Enabled -eq 'true'} -Properties * |
    Select -ExpandProperty Name | Out-FileUtf8NoBom "$ResultsPath\USON\Servers-USON.txt"  -Append

Get-Content "$ResultsPath\USON\Servers-USON.txt" | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\USON\Servers-Alive.txt" -Append
  } else { 
  write-output "$_" | Out-FileUtf8NoBom "$ResultsPath\USON\Servers-Dead.txt" -Append }}




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

#End function CheckTime
}

$time = CheckTime


#$filenameTXT = '$PSScriptRoot\Logs\Time ' + $SaveDate + '.txt'
$filenameHTML = '$PSScriptRoot\Logs\Time ' + $SaveDate + '.html'
#Then pick your poison
#$tstats | Out-File $filenameTXT
$tstats | convertto-html | Out-File $filenameHTML
# Check required services are running

$filename = '$PSScriptRoot\Logs\Services.' + $SaveDate + '.html'
$filename2 = '$PSScriptRoot\Logs\Report.' + $SaveDate + '.html'
$filename3 = '$PSScriptRoot\Logs\DomainHealth.' + $SaveDate + '.html'
$filename4 = '$PSScriptRoot\Logs\ExchangeHealth.' + $SaveDate + '.html'

Get-Service -ComputerName $AllSrvs -Name W3SVC   | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File $filename

Get-Service -Computername LASSVC03 -Name CollectionsAgentTimeService,Contoso* | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername $AllSrvs -Name ContosoDataLayerService  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASMCE01, LASMCE02 -Name CreditEngine  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername $CapsMTSrvs -Name CreditPullService,Contoso*  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASCHAT01 -Name WhosOn*  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename
Get-Service -Computername LASCHAT02 -Name WhosOn*  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASPROCDB02 -Name MSSQLSERVER,SQLSERVERAGENT | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASPROCAPP01 -Name P360*  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASPROCAPP04 -Name Service1  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASRFAX01 -Name RF*  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASCODE02 -Name AccuRev*,JIRA050414101333  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASITS02 -DisplayName *Symantec*   | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASPRINT02 -Name Spooler  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename

Get-Service -Computername LASINFRA02 -Name Schedule  | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-File -append $filename


## Getting AD Health
 
$DomainName = "Contoso.corp"
$ADHealthReport =  '$PSScriptRoot\Logs\ADHealthReport.' + $SaveDate + '.html'

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

 
##################################### 
#        Forest and Domain Info        # 
##################################### 
Write-Host " ..... Forest and Domain Information ..... " -foregroundcolor green 
Import-Module ActiveDirectory 
$Domain = $DomainName 
$ADForestInfo = Get-ADforest -server $Domain  
$ADDomainInfo = Get-ADdomain -server $Domain  
$DCs = Get-ADDomainController -filter * -server "$Domain"  
$allDCs = $DCs | ForEach-Object {$_.hostname} 
$ADForestInfo.sites | ForEach-Object {$Sites += "$($_) "} 
Write-Host "Forest Information" 
$ForestResults = New-Object Object 
$ForestResults | Add-Member -Type NoteProperty -Name "Mode" -Value (($ADForestInfo.ForestMode) -replace "Windows", "") 
$ForestResults | Add-Member -Type NoteProperty -Name "DomainNamingMaster" -Value $ADForestInfo.DomainNamingMaster 
$ForestResults | Add-Member -Type NoteProperty -Name "SchemaMaster" -Value $ADForestInfo.SchemaMaster 
$ForestResults | Add-Member -Type NoteProperty -Name "ForestName" -Value $ADForestInfo.name 
$ForestResults | Add-Member -Type NoteProperty -Name "Sites" -Value $Sites 
Write-Host "Domain Information" 
$DomainResults = New-Object Object 
$DomainResults | Add-Member -Type NoteProperty -Name "Mode" -Value (($ADDomainInfo.DomainMode) -replace "Windows", "") 
$DomainResults | Add-Member -Type NoteProperty -Name "InfraMaster" -Value $ADDomainInfo.infrastructuremaster 
$DomainResults | Add-Member -Type NoteProperty -Name "Domain" -Value $ADDomainInfo.name 
$DomainResults | Add-Member -Type NoteProperty -Name "PDC" -Value $ADDomainInfo.pdcemulator 
$DomainResults | Add-Member -Type NoteProperty -Name "RID" -Value $ADDomainInfo.ridmaster 
 
##################### 
#        DCDIAG        # 
##################### 
Write-Host " ..... DCDiag ..... " -foregroundcolor green 
$AllDCDiags = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
    $Dcdiag = (Dcdiag.exe /c /v /skip:OutBoundSecureChannels /skip:VerifyEnterpriseReferences /s:$DC) -split ('[\r\n]') 
    $Results = New-Object Object 
    $Results | Add-Member -Type NoteProperty -Name "ServerName" -Value $DC 
        $Dcdiag | ForEach-Object{ 
        Switch -RegEx ($_) 
        { 
         "Starting"      { $TestName   = ($_ -Replace ".*Starting test: ").Trim() } 
         "passed test|failed test" { If ($_ -Match "passed test") {  
         $TestStatus = "Passed"  
         # $TestName 
         # $_ 
         }  
         Else  
         {  
         $TestStatus = "Failed"  
          # $TestName 
         # $_ 
         } } 
        } 
        If ($TestName -ne $Null -And $TestStatus -ne $Null) 
        { 
         $Results | Add-Member -Name $("$TestName".Trim()) -Value $TestStatus -Type NoteProperty -force 
         $TestName = $Null; $TestStatus = $Null 
        } 
        } 
$AllDCDiags += $Results 
} 
 $Results.ServerName 
 $Results.Connectivity 
 $Results.Advertising 
 $Results.FrsEvent 
 $Results.DFSREvent 
 $Results.SysVolCheck 
 $Results.KccEvent 
 $Results.KnowsOfRoleHolders 
 $Results.MachineAccount 
 $Results.NCSecDesc 
 $Results.NetLogons 
 $Results.ObjectsReplicated 
 $Results.Replications 
 $Results.RidManager 
 $Results.Services 
 $Results.SystemLog 
 $Results.VerifyReferences 
 $Results.CheckSDRefDom 
 $Results.CrossRefValidation 
 $Results.LocatorCheck 
 $Results.Intersite 
 
##################################### 
#        OS Info and Uptime            # 
##################################### 
Write-Host " ..... OS Info and Uptime ..... " -foregroundcolor green 
$AllDCInfo = @() 
foreach ($DC in $allDCs) 
    { 
        $Results = New-Object Object 
        $Results | Add-Member -Type NoteProperty -Name "ServerName" -Value $DC 
        Write-Host "Processing $DC" 
        $computers = Get-WMIObject -class Win32_OperatingSystem -computer $DC    
        foreach ($system in $computers)  
        {    
            if(-not $computers) 
            { 
                $uptime = "Server is not responding : !!!!!!!!!!!! : !!!!!!!!!!!! : !!!!!!!!!!!!" 
                "$DC : $uptime"  
                $TestStatus = "Failed" 
                $Results | Add-Member -Type NoteProperty -Name "Uptime" -Value $TestStatus 
                $Results | Add-Member -Type NoteProperty -Name "C-DriveFreeSpace" -Value $TestStatus 
                $Results | Add-Member -Type NoteProperty -Name "OSVersion" -Value $TestStatus 
            } 
            else 
            { 
                $disk = ([wmi]"\\$DC\root\cimv2:Win32_logicalDisk.DeviceID='c:'") | ForEach-Object {[math]::truncate($_.freespace /1gb)} 
                #$OSVersion = $system.caption -replace "MicrosoftÂ® Windows ", "" 
                $Bootup = $system.LastBootUpTime    
                $LastBootUpTime = WMIDateStringToDate $bootup    
                $now = Get-Date  
                $Uptime = $now - $lastBootUpTime    
                $d = $Uptime.Days    
                $h = $Uptime.Hours    
                $m = $uptime.Minutes    
                $ms= $uptime.Milliseconds 
                $DCUptime = "{0} days {1} hours" -f $d,$h 
                $Results | Add-Member -Type NoteProperty -Name "Uptime" -Value $DCUptime 
                $Results | Add-Member -Type NoteProperty -Name "C-DriveFreeSpace" -Value $disk 
                #$Results | Add-Member -Type NoteProperty -Name "OSVersion" -Value $osversion 
            } 
         }  
    $AllDCInfo += $Results 
    } 
 $Results.ServerName 
 $Results.DCUptime 
 $Results.disk 
 $Results.osversion 
 
 
##################################### 
#        OS Info and Uptime            # 
##################################### 
Write-Host " ..... Server Time ..... " -foregroundcolor green 
$Timeresults=@() 
foreach ($DC in $alldcs) 
{ 
Write-Host "$DC " 
$computers = Get-WMIObject -class Win32_OperatingSystem -computer $DC |Select-Object CSName,@{Name="LocalDateTime";Expression={$_.ConvertToDateTime($_.LocalDateTime)}}  
$domaintimes = New-Object object 
$domaintimes | add-member -type NoteProperty -Name "DCName" -Value $computers.csname 
$domaintimes | add-member -type NoteProperty -Name "LocalDateTime" -Value $computers.LocalDateTime  
$Timeresults += $domaintimes 
}   
 
##################################### 
#        DC Test Ports                # 
##################################### 
Write-Host " ..... Testing Ports ..... " -foregroundcolor green 
$AllPortResults = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
$389 = test-port -comp $DC -port 389 -tcp  
$ResultsPort = New-Object Object 
$ResultsPort | Add-Member -Type NoteProperty -Name "ServerName" -Value $DC 
$ResultsPort | Add-Member -Type NoteProperty -Name "LDAP389" -Value $389.open 
 
$3268 = test-port -comp $DC -port 3268 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "LDAP3268" -Value $3268.open 
 
$53 = test-port -comp $DC -port 53 -udp  
$ResultsPort | Add-Member -Type NoteProperty -Name "DNS53" -Value $53.open 
 
$135 = test-port -comp $DC -port 135 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "RPC135" -Value $135.open 
 
$445 = test-port -comp $DC -port 445 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "SMB445" -Value $445.open 
 
$AllPortResults += $ResultsPort 
} 
 
######################################### 
#        DC Repadmin ReplSum                # 
######################################### 
Write-Host " ..... Repadmin /Replsum ..... " -foregroundcolor green 
$myRepInfo = @(repadmin /replsum * /bysrc /bydest /sort:delta /homeserver:$Domain) 
# Initialize our array. 
$cleanRepInfo = @()  
   # Start @ #10 because all the previous lines are junk formatting 
   # and strip off the last 4 lines because they are not needed. 
    for ($i=10; $i -lt ($myRepInfo.Count-4); $i++) { 
            if($myRepInfo[$i] -ne ""){ 
            # Remove empty lines from our array. 
            $myRepInfo[$i] -replace '\s+', " "  | Out-Null           
            $cleanRepInfo += $myRepInfo[$i]              
            } 
            }             
$finalRepInfo = @()    
            foreach ($line in $cleanRepInfo) { 
            $splitRepInfo = $line -split '\s+',8 
            if ($splitRepInfo[0] -eq "Source") { $repType = "Source" } 
            if ($splitRepInfo[0] -eq "Destination") { $repType = "Destination" } 
            if ($splitRepInfo[1] -notmatch "DSA") {        
            # Create an Object and populate it with our values. 
           $objRepValues = New-Object System.Object  
               $objRepValues | Add-Member -type NoteProperty -name DSAType -value $repType # Source or Destination DSA 
               $objRepValues | Add-Member -type NoteProperty -name Hostname  -value $splitRepInfo[1] # Hostname 
               $objRepValues | Add-Member -type NoteProperty -name Delta  -value $splitRepInfo[2] # Largest Delta 
               $objRepValues | Add-Member -type NoteProperty -name Fails -value $splitRepInfo[3] # Failures 
               #$objRepValues | Add-Member -type NoteProperty -name Slash  -value $splitRepInfo[4] # Slash char 
               $objRepValues | Add-Member -type NoteProperty -name Total -value $splitRepInfo[5] # Totals 
               $objRepValues | Add-Member -type NoteProperty -name PctError  -value $splitRepInfo[6] # % errors    
               $objRepValues | Add-Member -type NoteProperty -name ErrorMsg  -value $splitRepInfo[7] # Error code 
            
            # Add the Object as a row to our array     
            $finalRepInfo += $objRepValues 
             
            } 
            } 
 
##################################### 
#        SEP Version                # 
##################################### 
Write-Host " ..... SEP Version ..... " -foregroundcolor green 
$AllSEPResults = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
$SEPresult = Get-SepVersion -ComputerName $DC -ErrorAction silentlycontinue | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate 
$AllSEPResults += $SEPresult 
} 

##################################### 
#        Exchange Health            # 
##################################### 
GoGo-PSExch

"$PSScriptRoot\ExchangeAnalyzer\Run-ExchangeAnalyzer.ps1 -FileName $ExchangeHealth"



##################################### 
#        Compile HTML               # 
##################################### 
$style = "BODY{font-family: Arial; font-size: 10pt;}" 
$style = "<style>BODY{color:#717D7D;background-color:#F5F5F5;font-size:10pt;font-family:'trebuchet ms', helvetica, sans-serif;font-weight:normal;padding-:0px;margin:0px;overflow:auto;}" 
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}" 
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }" 
$style = $style + "TD{font-weight: bold; border: 1px solid black; padding: 5px; }" 
$style = $style + "</style>" 
$HTML = "<h2>Forest Information</h2></br>" 
$HTML += $ForestResults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Domain Information</h2></br>" 
$HTML += $DomainResults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Domain Controller Information</h2></br>" 
$HTML += $DCs | Select-Object HostName,Site,Ipv4Address,OperatingSystem,OperatingSystemServicePack,IsGlobalCatalog | ConvertTo-HTML -head $style 
$HTML += "</br>" 
$HTML += $AllDCInfo | ConvertTo-HTML -head $style 
$HTML += "</br>" 
$HTML += $Timeresults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>DCDiag Results</h2></br>" 
$HTML += $AllDCDiags | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Replication Information</h2></br>" 
$HTML += $finalRepInfo | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Port Tests</h2></br>" 
$HTML += $AllPortResults| ConvertTo-HTML -head $style 
$HTML += "</br><h2>SEP Versions</h2></br>" 
$HTML += $AllSEPResults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Exchange Health Report</h2></br>"
$HTML += $ExchangeHealth 
$HTML = $HTML -Replace ('failed', '<font color="red">Failed</font>') 
$HTML = $HTML -Replace ('passed', '<font color="green">Passed</font>') 
$HTML | Out-File $ADHealthReport

#Invoke-Item $filename
#Invoke-Item $filename2
#Invoke-Item $filename3

########################################################################################
#############################################Send Email#################################


$Body = (Get-Content $filename), (Get-Content $time), (Get-Content $ADHealthReport) | Out-String

Send-MailMessage -To $to -From $from -Subject "Active Directory Health Monitor" -Body $Body -BodyAsHtml -SmtpServer $smtpserver

########################################################################################

########################################################################################