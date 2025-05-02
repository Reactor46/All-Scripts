<#
Scripted DailyCheckList
Contains various scripts and functions that will gather
the required information for the Daily Checklist

Written by: John Battista

V1.00, 9/18/2017 - Initial version
V1.2, 6/24/2019 - Updates and changes for USON
#>
# Set the date the script was run to the end of each report
# Ex: $Report = '.\Report ' + $SaveDate + '.html'
####******************* set email parameters ****************** ######
$from = "DailyCheckList@uson.local" 
$to="jbattista@optummso.com"
$smtpserver="192.168.1.8"
####*******************######################****************** ######
$SaveDate = (Get-Date).tostring("MM-dd-yyyy")
$timeout = "60"
$DCSrvs = "USONVSVRDC01","USONVSVRDC02","USONVSVRDC03","MSODC04"
$FaxSrv = "USONVSVRFAX01"
$FSSrvs = "USONPSVRFPF","USONVSVRFS01"
$Email = "USONVSVREX01","USONVSVRDAG01","USONVSVRDAG02"
$VarianSrvs = "USONPSVRRO1","USONVSVRAPP01","USONVSVRAPP02","USONVSVRAPP03","USONVSVRAPP04"
$SepSrv = "USONVSVRSEP"
$VoIPSrvs = "USONVSVRVOIP01","USONVSVRVOIP02","USONVSVRVOIP03"
$BakSrvs = "USONVSVRBK01","USONVSVRBK2"
$SSRSSrv = "USONVSVRSQL01"
$npt = $DCSrvs


$time = CheckTime

$filenameHTML = 'E:\Scripts\Logs\Time ' + $SaveDate + '.html'

$tstats | convertto-html | Out-File $filenameHTML
# Check required services are running

$filename = 'E:\Scripts\Logs\Services.' + $SaveDate + '.html'
$filename2 = 'E:\Scripts\Logs\Report.' + $SaveDate + '.html'
$filename3 = 'E:\Scripts\Logs\DomainHealth.' + $SaveDate + '.html'
$filename4 = 'E:\Scripts\Logs\ExchangeHealth.' + $SaveDate + '.html'


$MSExcServices = "MSExchangeAB","MSExchangeADTopology","MSExchangeAntispamUpdate","MSExchangeEdgeSync","MSExchangeFBA","MSExchangeFDS","MSExchangeImap4","MSExchangeMailboxReplication","MSExchangeMonitoring","MSExchangePop3","MSExchangeProtectedServiceHost","MSExchangeRPC","MSExchangeServiceHost","MSExchangeTransport","MSExchangeTransportLogSearch"
$ADDCServices = "ADWS","DFS","DFSR","DHCP","DNS","MpsSvc","Netlogon","NTDS","NtFrs","ProfSvc","RemoteRegistry","SamSs","TermService","W32Time","Winmgmt","WinRM","KDC"

$DCSrvs |
ForEach {
Get-Service -ComputerName $_ -Name $ADDCServices | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-FileUtf8NoBom -Append $filename
}
$Email |
ForEach {
Get-Service -ComputerName $_ -Name $MSExcServices | Select-Object MachineName, Status, Displayname | ConvertTo-Html | Out-FileUtf8NoBom -Append $filename
}



## Getting AD Health
 
$DomainName = "USON.LOCAL"
$ADHealthReport =  'E:\Scripts\Results\ADHealthReport.' + $SaveDate + '.html'

 
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
#        Report Gathering           # 
##################################### 
$ExchangeRpt = \\USONVSVREX01\E$\Scripts\


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