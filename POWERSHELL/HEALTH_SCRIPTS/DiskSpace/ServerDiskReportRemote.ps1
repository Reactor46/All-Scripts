param (
[CmdletBinding()]
[switch]$GetADComputers,
[int]$DiskWarnThresh = 20,
[switch]$ShowOnlyWarnings,
[string]$FilePath = (Get-Location).Path
)


#region Variables and Arguments    
$ErrLogPath =  ".\ServerDiskReport-Errorlog.txt"
$date= Get-Date -format g
$version = "4.0 (Remote Edition)" #<<<<<<CHANGE THIS ANYTIME MODIFIED!
$RedBang = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAA2FBMVEX///++MjLnalHTJyDVKy3veGf8p6Tpd2HWYl/UKiPvemncTU/leF3fWTf2kYjZPkDWLjDUXlnwhHXXOjTUSyrYaGf1iH7NS0ncYj/USy/eVjTPJxPMJwDMJwb1hnzYOz7ONAX8o6DVQxnhbG7NTE3QKhfZPkHWVS7SXlfic3XnbFTZU0P+/v7PVle1IiHMR0ieGwD1qqD3xLyzSTLLf3DPNBaQHwL99fTcalDjeHTCb2uOFQ7KJhL21MyGJxC3LhWoHQTvuKnkycKAHgjsrqb01tPmmYLSnpF1j2u3AAAAAXRSTlMAQObYZgAAAKxJREFUeF5lz0WWw0AMRVEXmZkdZmZmbNj/jlpKO6PcyT/1RioJkZyUI5FcB3KUF6LacQXEtvoqxPN1IeYLIXTfg0J6yoBSuh0tKR0oPSIR0yqA8+gCY5kQ0ga6P544KYRuDf38bnC6EMIS2q1mOCGErIzW1xtOBmHYLoKv7xNMewghCTTG2P5wZEwLEjxk6ric88mEu870/9SW0amCjtF6f2bcb4L+GN8f3/8DQgYUGINoSt8AAAAASUVORK5CYII='
$RedBangHTML = "<img src=data:image/png;base64,$($RedBang) alt='Error' style='vertical-align:middle;padding-left:5px;padding-right:5px;'/>"
$YellowBang = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAACdklEQVR42mVSTU8TURR9JCQsoJKIq0oUookpBkJnphqXtD+AX8BfcMOWD2PbtCaaYCKRIhRFi1q1Frvwk6ihKMFFi9ICg0BRISIFtbQKKe3xvjcd08LizH25953zzr13mNVqZS0tLQcw55TaVLcE1W3GvEtp0+/ZbDb9juCWkCghQOQKTt55XS8ghFxKhS7Co35mxcq6esyhdK8P1wGfmoCPTVgfrkfMLnVT7YBjPVFMbvh8uRHZD8eBGVkgO3lMtDJjlxr2t8qJJarc7s/Ro0C8GblcDvl8HtuTjdjw10K9JGO/C1bc16xDal2+ZiLrJ4A5CblsSohER05hc6wOiasmzDqV1uKZ6Q5EnCebmZdGIK4I7GxNCQf+C7X4GjqNdMhIGzGjRIA+hd6l89/66eXpBiJbgFipQCJA+XcmrPadRJzulmwhelGuWaD+dl/VAlEztdBM0yc3q0H8WHkvBNaCJhKg1ZILPqdpu1Lz/z+IO2TXhpesjzHgLWGcEGbILPViJTKAqSsMSR/lgoTHdO43IuZUXJzLB3eO97U3Wg48Z5rIG44y/H1RjRVfNZY8DNt3KfeQcJ9hz1cuNhInLqMfxPfbe0RTDxGeMk3oGcOftSB2kxEk+hjStyl3j8Cd3GH4df0w6J/xMepdKzzQ7GG0IPSEWljuxZfIoJjB1kCBzIVuEW6WgW+NcfsY0azBX7AZIDyiV4mw0MOg0gxSXv6yTmbIDzExTOEgM2BAZsiAtLcS6aEq7UxIDVYh6anEJiF1w0D3DiHDo4fqhO89dWCB9jPuiQ5ZDXdaFie6LIsiEsY7ZC12KovhLp7T60qhpqiB9rPuf+YeJOq16/K9AAAAAElFTkSuQmCC'
$YellowBangHTML = "<img src=data:image/png;base64,$($YellowBang) alt='Warning' style='vertical-align:middle;padding-right:5px;'/>"

#endregion Variables and Arguments

#region Admin Check
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] “Administrator”))
{
    Write-Warning “Administrator rights are required to run this script! Please re-run this script as an Administrator!”
    Write-Output ("Time: " + (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " ---- Error: Script Requires Administrative Rights!") |Out-File -FilePath .\EventReport-Errorlog.txt -Append
    Start-Sleep -Seconds 5
    Break
}
#endregion Admin Check

if ($GetADComputers)
    {
        try
            {
                Import-Module ActiveDirectory -ErrorAction Stop
                $computers = Get-ADComputer -filter 'enabled -eq $True' -ErrorAction Stop|Sort Name |Select -expand Name
            }
      
        catch
            {
                Write-Output ("Time: " + (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " ---- Error: Get-ADComputer")|Set-Variable ADErrMsg
                "GetADComputer Error - ExceptionMessage: " + $_.Exception.Message
                $ADErrMsg| Out-File -FilePath $ErrLogPath -Append
                "                             ExceptionMessage: " + $_.Exception.Message| Out-File $ErrLogPath -Append
                break
            }  
    }  
else
{
        try
            {
                $computers = get-content .\list.txt -ErrorAction Stop
                if ((Get-Item .\list.txt).Length -eq 0)
                    {
                       Write-Output ("Time: " + (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " ---- Error: 'list.txt' in - " + "'" + (Get-Location).Path +"'" + " is empty! ---- Please place desired computer names in list.txt file, or use the -GetADComputer parameter")|
                       Set-Variable ListErrMsg
                       $ListErrMsg
                       $ListErrMsg | Out-File -FilePath $ErrLogPath -Append
                       break
                    }
            }
     
        catch
            {
                Write-Output ("Time: " + (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " ---- Error: No 'list.txt' found."+ " ---- Please place desired computer names in list.txt file, or use the -GetADComputer parameter")|
                Set-Variable ListErrMsg
                $ListErrMsg
                Write-Warning "ExceptionMessage: $($_.Exception.Message)"
                $ListErrMsg | Out-File -FilePath $ErrLogPath -Append
                "                             ExceptionMessage: $($_.Exception.Message)"|
                Out-File -FilePath $ErrLogPath -Append
                break             
            }
    }
#>

cls
""
Write-Host -ForegroundColor Yellow "Checking Computers for connectivity before running script..."
Write-Host -ForegroundColor Yellow "See Errorlog.txt for servers that failed to connect"
Start-Sleep -Seconds 1
""
#region Remote Check
$CompPlaceHolder = @()
$i = 0
foreach ($computer in $computers) {
$i ++
Write-Progress -Activity "Testing Computer Connections" -CurrentOperation $computer -PercentComplete (($i / $computers.count) * 100 )       
try
{
    Write-Host -ForegroundColor Yellow “Verifying Remote RPC/WMI Connectivity on" $computer "...”
    $TestConnect = Test-Connection -ComputerName $computer -Count 1 -WarningAction Stop -EA Stop
    $WMIConnect= Get-WmiObject Win32_ComputerSystem -computername $computer -EA Stop
    Write-Host -ForegroundColor Green "Remote RPC/WMI connection successful!"
    ""
    $Connections = New-Object -Type PSObject -Property @{
            "Name"          = $computer
            }
            $CompPlaceHolder += $Connections
}
       
catch                
      {
     
            Write-Host -ForegroundColor Red "Remote RPC/WMI on server: $computer is unreachable. See $ErrLogPath"
            Write-Output ("Time: " + (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " ---- Error: Connecting to RPC/WMI on " + $computer)  |Out-File -FilePath $ErrLogPath -Append
            ("                             ExceptionMessage: " + $($_.Exception.Message)) | Out-File -FilePath $ErrLogPath -Append
          
          }
                   
}
$computers = $CompPlaceHolder|Select -ExpandProperty Name #Use of $CompPlaceHolder.name doesn't work right in PS 2.0
#endregion Remote Check

#region Functions
Function Get-LogicalDiskInformation {
    param(
          [string[]]
        $ComputerName=$env:computername
          )
   
 
$ObjHolder = @()   
foreach ($computer in $ComputerName)
        {
        $LogicalDiskTable = @()
        $LogicalDiskInfo  = Get-WMIObject -ComputerName $computer -Query "Select SystemName,VolumeName,Name,Size,FreeSpace FROM Win32_LogicalDisk Where DriveType=3"
       
        foreach ($dsk in $LogicalDiskInfo){
      
        $Disk = New-Object -Type PSObject -Property @{
              
                SystemName    = $dsk.SystemName
                VolumeName    = $dsk.VolumeName
                Name          = $dsk.Name
                SizeInGB      = if ($dsk.Size){[math]::Round(($dsk.Size/1gb),2)}else{0}
                FreeSpaceInGB = if ($dsk.Size){[math]::Round(($dsk.FreeSpace/1gb),2)}else{0} 
                PercentFree   = if ($dsk.Size){[math]::Round((($dsk.freespace/$dsk.size)*100),0)}  else{0}         
                } 
                $LogicalDiskTable +=$Disk
            }
 
        $ObjHolder += $LogicalDiskTable
        }
    $ObjHolder|Sort SystemName   
}
#endregion Functions

#* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
$HTMLMiddle = @()

$i = 0
foreach ($computer in $computers) {
$i ++
Write-Progress -Activity "Gathering Disk Information" -CurrentOperation $computer -PercentComplete (($i / ($computers).count) * 100 ) 
	    
#region Disks
$physDiskinfo       = Get-WmiObject -ComputerName $computer -query "SELECT Model,FirmwareRevision,DeviceID,Size,InterfaceType,BytesPerSector,Status FROM Win32_DiskDrive" |
                      Select-Object Model,FirmwareRevision,DeviceID, @{n='Size (GB)';e={"{0:n2}" -f ($_.size/1gb)}},InterfaceType,BytesPerSector,Status

$physDiskinfoWarnCt = ($physDiskinfo|Where {$_.Status -ne "OK"}).Count

$diskInfo           = Get-LogicalDiskInformation -ComputerName $computer | 
                      Select SystemName,VolumeName,Name,SizeInGB,FreeSpaceInGB,PercentFree|Sort Name

$diskInfoWarnCt     = ($diskInfo|Where {$_.PercentFree -le $DiskWarnThresh}|Measure).Count

$logicalDiskResult  = @()

foreach ($dsk in $diskInfo) {
                
                $var = New-Object -Type PSObject -Property @{
                       
                        SystemName    = $Dsk.SystemName
                        VolumeName    = $Dsk.VolumeName
                        Name          = $Dsk.Name
                        SizeInGB      = $Dsk.SizeInGB
                        FreeSpaceInGB = $Dsk.FreeSpaceInGB
                        PercentFree   = if ($Dsk.PercentFree -le $DiskWarnThresh){
                                                "<p class=`"DiskWarn`">" + $Dsk.PercentFree +"%" + "</p>"
                                                }
                                           
                                            else
                                                {
                                                    "$($Dsk.PercentFree)" + "%"
                                                }
                        }
                     $LogicalDiskResult += $var

        }

$diskInfoHtml       = $logicalDiskResult|Select SystemName,VolumeName,Name,SizeInGB,FreeSpaceInGB,PercentFree|
                      ConvertTo-Html -Fragment

$physDiskinfoHtml   = $physDiskinfo|Sort DeviceID|Select Model,FirmwareRevision,DeviceID,InterfaceType,BytesPerSector,"Size (GB)",@{l="Status";e={if ($_.Status -ne "OK") {("<p class=`"DiskWarn`">" + $_.Status + "</p>")} else {$_.Status} }}| ConvertTo-Html -Fragment

$totalDiskWarnCt    = $diskInfoWarnCt + $physDiskinfoWarnCt

$totalDiskWarnCtHtml= if ($totalDiskWarnCt -gt 0) {("<font color=`"red`">" + ($totalDiskWarnCt.ToString()) + "</font>" + " " + $RedBangHTML)} else {("<font color=`"green`">" + $totalDiskWarnCt + "</font>")}

#endregion Disks


#Create HTML Report for the current System being looped through
$CurrentSystemHTML = @"
	
	<input class="toggle-box" id="identifier-$computer-1" type="checkbox" name="grouped"><label for="identifier-$computer-1"> Disk Report | Disk Warnings: $totalDiskWarnCtHtml - $computer</label>  
	<div>
    
	 
    <h3>Logical Disks</h3>
    <p>List of logical disks on $computer. Free space percentage less than or equal to 5% are highlighted.</p>
	$diskInfoHtml
    <h3>Physical Disks</h3>
    <p>List of physical disks on $computer. Error states are highlighted.</p>
    $physDiskinfoHtml
    
    

    </div> 
    


"@

if ($ShowOnlyWarnings)
    { 
        if ($totalDiskWarnCt -lt 1) 
            {
            (Get-Date -Format "MM.d.yyyy - hh:mm:ss") + " --- Disks OK - $($computer)"
            } 
        
        else {
        $HTMLMiddle += $CurrentSystemHTML
        }                   
    }

else {
        $HTMLMiddle += $CurrentSystemHTML
    }
}

#region Assemble the HTML Header and CSS for Report

$HTMLHeader = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>Server Disk Report</title>
<style type="text/css">
<!--
body {
        font: 80% "Proxima Nova",Helvetica,Arial,Tahoma,Verdana,sans-serif;
        margin: 0;
        padding: 0;
        
        }     

h1{
        background-color: inherit;
        color: #606060;
        font-size: 1.8em;
        font-weight: normal;
        margin: 0 0 15px;
        padding: 0;
        letter-spacing: 1px;
		
        }

h2{
        background-color: inherit;
        color: #606060;
        font-size: 1.6em;
        font-weight: normal;
        margin: 0 0 15px;
        padding: 0;
        letter-spacing: 1px;
		
        }
h3 {

        color: #606060;
        font-size: 1.4em;
        font-weight: normal;
        margin: 0 0 15px;
        padding: 0;
        letter-spacing: 1px;
        }

p{
        margin-left: 10px;    
        }

p.DiskWarn {

        background-color:#FB7171;
        text-align:left;
        font-weight:bold;
        padding:0px;
        margin:0px;
        
        }

table {
        font-family: inherit;
        border: 1px solid #CCCCCC;
        border-radius: 3px;
        margin-bottom: 15px;
		
        }

th {			
	    background: #f3f3f3;
        border-bottom: 1px solid #CCCCCC;
        font-size: 14px;
        font-weight: normal;
        padding: 10px;
        text-align: center;
        vertical-align: middle;	
	    }

tr {
        border-bottom: 1px solid #CCCCCC;
        }

td {
		padding: 10px 8px 10px 8px;
        font-size: 12px;
        text-align: left;
        vertical-align: middle;
        }
		
tr:hover td
        {
		background-color: #007cbb;
        Color: #F5FFFA;
	    }
		
tr:nth-child(odd) 
        {
		background: #F9F9F9;
	    } 

table.list {
        
        width:100%;
        float:left;
        }
        
.toggle-box {
display: none;
}

.toggle-box + label {
padding: 10px;
cursor: pointer;
display: block;
clear: both;
font-size: 16px;
margin-right: auto;
margin-bottom:5px;
text-align: left;
}

.toggle-box + label:hover {
text-shadow:1px 1px 1px rgba(0,0,0,0.1)
}

.toggle-box + label + div {
display: none;
margin-left: 0px;
margin-right: auto;
}

.toggle-box:checked + label {

font-style: italic;
}

.toggle-box:checked + label + div {
display: block;
margin-right: auto;

}

.toggle-box + label:before {
content: "";
display: block;
float: left;
border-right: 2px solid;
border-bottom: 2px solid;
width: 5px;
height: 5px;
transform: rotate(-45deg);
margin-top: 6px;
margin-right: 20px;
margin-left: auto;
text-align: left;
-webkit-transition: all 0.5s;
transition: all 0.5s;
}

.toggle-box:checked + label:before {
border-right: 2px solid;
border-bottom: 2px solid;
width: 5px;
height: 5px;
transform: rotate(45deg);
-webkit-transition: all 0.5s;
transition: all 0.5s;
}	

  
-->
</style>
</head>
<body>
<h2>Server Disk Report</h2>
<p><i>Generated by Server Disk Report Tool version $version on $date.</i></p>
<p>The servers indicated below have free disk space percentage less than or equal to $DiskWarnThresh %</p>
"@

#endregion Assemble the HTML Header and CSS for Report

# Assemble the closing HTML for our report.
$HTMLEnd = @"
</body>
</html>
"@

# Assemble the final report from all our HTML sections
""
Write-Host -ForegroundColor Yellow "Assembling HTML for final report..."
$HTMLmessage = $HTMLHeader + $HTMLMiddle + $HTMLEnd
# Save the report out to a file in the current path
Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($HTMLmessage) | Out-File ($FilePath + "\ServerDiskReport.html")
""
Write-Host -ForegroundColor Green "Complete!"
""
""
Write-Host -ForegroundColor DarkYellow ("File Written: " + $FilePath + "\ServerDiskReport.html")

# Email our report out
# send-mailmessage -from $fromemail -to $users -subject "Systems Report" -Attachments $ListOfAttachments -BodyAsHTML -body $HTMLmessage -priority Normal -smtpServer $server

