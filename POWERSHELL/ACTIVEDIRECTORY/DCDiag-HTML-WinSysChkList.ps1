# Example : .\chklst_DCDiag-HTML.ps1 -DomainName "Contoso.corp" -HTMLFileName "C:\Scripts\Report.html" 
 
 
[CmdletBinding()]Param( 
[Parameter(Mandatory=$True,Position=1)][string]$DomainName, 
[Parameter(Mandatory=$True,Position=2)][string]$HTMLFileName 
) 

##################################### 
#        Forest and Domain Info     # 
##################################### 
Get-ChildItem "$PSScriptRoot\Functions\*.ps1" | %{.$_} -ErrorAction SilentlyContinue
Remove-Module Carbon -ErrorAction SilentlyContinue
Write-Host " ..... Forest and Domain Information ..... " -foregroundcolor green 
#Import-Module ActiveDirectory 
$Domain = $DomainName 
$ADForestInfo = Get-ADforest -server $Domain  
$ADDomainInfo = Get-ADdomain -server $Domain  
$DCs = Get-ADDomainController -filter * -server "$Domain"  
$allDCs = $DCs | foreach {$_.hostname} 
$ADForestInfo.sites | foreach {$Sites += "$($_) "} 
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
#        DCDIAG     # 
##################### 
Write-Host " ..... DCDiag ..... " -foregroundcolor green 
$AllDCDiags = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
    $Dcdiag = (Dcdiag.exe /c /v /skip:OutBoundSecureChannels /skip:VerifyEnterpriseReferences /s:$DC) -split ('[\r\n]') 
    $Results = New-Object Object 
    $Results | Add-Member -Type NoteProperty -Name "ServerName" -Value $DC 
        $Dcdiag | %{ 
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
#        OS Info and Uptime         # 
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
                #$OSVersion = $system.caption -replace "Microsoft® Windows ", "" 
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
#        OS Info and Uptime         # 
##################################### 
Write-Host " ..... Server Time ..... " -foregroundcolor green 
$Timeresults=@() 
foreach ($DC in $allDCs) 
{ 
Write-Host "$DC " 
$computers = Get-WMIObject -class Win32_OperatingSystem -computer $DC |Select-Object CSName,@{Name="LocalDateTime";Expression={$_.ConvertToDateTime($_.LocalDateTime)}}  
$domaintimes = New-Object object 
$domaintimes | Add-Member -type NoteProperty -Name "DCName" -Value $computers.csname 
$domaintimes | Add-Member -type NoteProperty -Name "LocalDateTime" -Value $computers.LocalDateTime  
$Timeresults += $domaintimes 
}   
 
##################################### 
#        DC Test Ports              # 
##################################### 
Write-Host " ..... Testing Ports ..... " -foregroundcolor green 
$AllPortResults = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
$389 = Test-Port -comp $DC -port 389 -tcp  
$ResultsPort = New-Object Object 
$ResultsPort | Add-Member -Type NoteProperty -Name "ServerName" -Value $DC 
$ResultsPort | Add-Member -Type NoteProperty -Name "LDAP389" -Value $389.open 
 
$3268 = Test-Port -comp $DC -port 3268 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "LDAP3268" -Value $3268.open 
 
$53 = Test-Port -comp $DC -port 53 -udp  
$ResultsPort | Add-Member -Type NoteProperty -Name "DNS53" -Value $53.open 
 
$135 = Test-Port -comp $DC -port 135 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "RPC135" -Value $135.open 
 
$445 = Test-Port -comp $DC -port 445 -tcp  
$ResultsPort | Add-Member -Type NoteProperty -Name "SMB445" -Value $445.open 
 
$AllPortResults += $ResultsPort 
} 
 
######################################### 
#        DC Repadmin ReplSum            # 
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
 


############### 
# SEP Version # 
############### 
Write-Host " ..... SEP Version ..... " -foregroundcolor green 
$AllSEPResults = @() 
foreach ($DC in $allDCs) 
{ 
Write-Host "Processing $DC" 
$SEPresult = $DC | Get-SEPVersion | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate 
$AllSEPResults += $SEPresult 
} 
 
##################################### 
#        Compile HTML               # 
##################################### 
#$style = "BODY{font-family: Arial; font-size: 10pt;}" 
#$style = "<style>BODY{color:#717D7D;background-color:#F5F5F5;font-size:10pt;font-family:'trebuchet ms', helvetica, sans-serif;font-weight:normal;padding-:0px;margin:0px;overflow:auto;}" 
#$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}" 
#$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }" 
#$style = $style + "TD{font-weight: bold; border: 1px solid black; padding: 5px; }" 
#$style = $style + "</style>" 

$Style = @"
 <header>
</header>
<style>
BODY{font-family:Calibri;font-size:12pt;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:black;background-color:#0BC68D;text-align:center;}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align:center;}
</style>
"@

$HTML = "<h2>Forest Information</h2></br>" 
$HTML += $ForestResults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Domain Information</h2></br>" 
$HTML += $DomainResults | ConvertTo-HTML -head $style 
$HTML += "</br><h2>Domain Controller Information</h2></br>" 
$HTML += $DCs | Select HostName,Site,Ipv4Address,OperatingSystem,OperatingSystemServicePack,IsGlobalCatalog | ConvertTo-HTML -head $style 
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
$HTML = $HTML -Replace ('failed', '<font color="red">Failed</font>') 
$HTML = $HTML -Replace ('passed', '<font color="green">Passed</font>') 
$HTML | Out-File $HTMLFileName