#Set-ExecutionPolicy Unrestricted -Force -ErrorAction SilentlyContinue
################################################################################
# Author          : Syed Abdul Khader 
# Description     : Get server inventory and send email 
################################################################################
Function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML>
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>
<meta name="save" content="history">
<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 8px; COLOR: #ffffff; FONT-FAMILY: Arial; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
.dspcont { display:none; BORDER-RIGHT: #B1BABF 1px solid; BORDER-TOP: #B1BABF 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; WIDTH: 95%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #CC0000}
a.dsphead1 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead2 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}
td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}
th {VERTICAL-ALIGN: TOP; COLOR: #CC0000; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt} 
</style>
<script type="text/javascript">
function dsp(loc){
   if(document.getElementById){
      var foc=loc.firstChild;
      foc=loc.firstChild.innerHTML?
         loc.firstChild:
         loc.firstChild.nextSibling;
      foc.innerHTML=foc.innerHTML=='hide'?'show':'hide';
      foc=loc.parentNode.nextSibling.style?
         loc.parentNode.nextSibling:
         loc.parentNode.nextSibling.nextSibling;
      foc.style.display=foc.style.display=='block'?'none':'block';}}  
if(!document.getElementById)
   document.write('<style type="text/css">\n'+'.dspcont{display:block;}\n'+ '</style>');
</script>
</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#CC0000">
<font face="Arial" size="1">Report created on $(Get-Date)</br></font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
Return $Report
}

Function Get-CustomHeader0 ($Title){
$Report = @"
		<h1><a class="dsphead0">$($Title)</a></h1>
	<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader ($Num, $Title){
$Report = @"
	<h2><a href="javascript:void(0)" class="dsphead$($Num)" onclick="dsp(this)">
	<span class="expando">show</span>$($Title)</a></h2>
	<div class="dspcont">
"@
Return $Report
}

Function Get-CustomHeaderClose{

	$Report = @"
		</DIV>
		<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader0Close{

	$Report = @"
</DIV>
"@
Return $Report
}

Function Get-CustomHTMLClose{

	$Report = @"
</div>
</body>
</html>
"@
Return $Report
}

Function Get-HTMLTable{
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE html>', ""
	$HTMLTable = $HTMLTable -replace '<head>', ""
	$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
	$HTMLTable = $HTMLTable -replace '</head><body>', ""
	$HTMLTable = $HTMLTable -replace '</body></html>', ""
	Return $HTMLTable
}

Function Get-HTMLDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}

Function Get-HTMLRolesDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
   	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}

$target = $env:computername
	$ComputerSystem = Get-WmiObject -computername $Target Win32_ComputerSystem
	switch ($ComputerSystem.DomainRole){
		0 { $ComputerRole = "Standalone Workstation" }
		1 { $ComputerRole = "Member Workstation" }
		2 { $ComputerRole = "Standalone Server" }
		3 { $ComputerRole = "Member Server" }
		4 { $ComputerRole = "Domain Controller" }
		5 { $ComputerRole = "Domain Controller" }
		default { $ComputerRole = "Information not available" }
	}
	
	$OperatingSystems = Get-WmiObject -computername $Target Win32_OperatingSystem
	$TimeZone = Get-WmiObject -computername $Target Win32_Timezone
	$Keyboards = Get-WmiObject -computername $Target Win32_Keyboard
	$SchedTasks = Get-WmiObject -computername $Target Win32_ScheduledJob
	$BootINI = $OperatingSystems.SystemDrive + "boot.ini"
	$RecoveryOptions = Get-WmiObject -computername $Target Win32_OSRecoveryConfiguration
	
	switch ($ComputerRole){
		"Member Workstation" { $CompType = "Computer Domain"; break }
		"Domain Controller" { $CompType = "Computer Domain"; break }
		"Member Server" { $CompType = "Computer Domain"; break }
		default { $CompType = "Computer Domain"; break }
	}

	$LBTime=$OperatingSystems.ConvertToDateTime($OperatingSystems.Lastbootuptime)
	####################################################################################
    # General Settings
    ####################################################################################
	$MyReport = Get-CustomHTML "$Target System Info"
	$MyReport += Get-CustomHeader0  "$Target Details"
	$MyReport += Get-CustomHeader "2" "General"
		$MyReport += Get-HTMLDetail "Computer Name" ($ComputerSystem.Name)
		#$MyReport += Get-HTMLDetail "Computer Role" ($ComputerRole)
		$MyReport += Get-HTMLDetail $CompType ($ComputerSystem.Domain)
		$MyReport += Get-HTMLDetail "Operating System" ($OperatingSystems.Caption)
		$MyReport += Get-HTMLDetail "Service Pack" ($OperatingSystems.CSDVersion)
		$MyReport += Get-HTMLDetail "System Root" ($OperatingSystems.SystemDrive)
		$MyReport += Get-HTMLDetail "Manufacturer" ($ComputerSystem.Manufacturer)
		$MyReport += Get-HTMLDetail "Model" ($ComputerSystem.Model)
		$MyReport += Get-HTMLDetail "Number of Processors" ($ComputerSystem.NumberOfLogicalProcessors)
		$MyReport += Get-HTMLDetail "Memory" ([Math]::Round($ComputerSystem.TotalPhysicalMemory/1GB,2))
		#$MyReport += Get-HTMLDetail "Last System Boot" ($LBTime)
		$MyReport += Get-CustomHeaderClose
    ################################ End of General Settings############################
    ####################################################################################
    # Local Users
    ####################################################################################
    $localusers = Get-WmiObject win32_UserAccount -ComputerName $Target -Filter "LocalAccount='$True'" | select Name,PasswordExpires,Disabled,Lockout
    $MyReport += Get-CustomHeader "2" "Local Users"
	$MyReport += Get-HTMLTable ($localusers)
	$MyReport += Get-CustomHeaderClose
    ################################ End of Users  ####################################
    ####################################################################################
    # Local Group and Members
    ####################################################################################
    If ($OperatingSystems -contains 'Microsoft Windows Server 2012*')
    {
    	$GroupName = get-wmiobject win32_group -Filter “LocalAccount=True"
	$groups = $GroupName.name
    	$rep = @()
	$MyReport += Get-CustomHeader "2" "Local Group and Members"
	$Rep = @()
    	foreach ($Group in $groups)
    	{
        	$members = net localgroup $Group | where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 4
        	$groupmem= $null
        	If ($members -ne $null)
        	{
            		foreach ($member in $members)
            		{
                		if($groupmem -ne $null)
                		{
                    			$groupmem = -join($groupmem,", ", $member)
                		}
                		Else
                		{
                    			$groupmem = -join($member)
                		}
            		}
        	}
		$Details=""|Select "Group","Members"
        	$Details.Group= $Group
        	$Details.Members = $groupmem
       		$rep +=$Details
        	$Details = $null
    	}
    	$MyReport += Get-HTMLTable ($rep)
	$MyReport += Get-CustomHeaderClose
    }
    Else
    {
	$Groups = net localgroup| where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 2
	$Groups = $Groups -replace '[*]',''
    	$rep = @()
	$MyReport += Get-CustomHeader "2" "Local Group and Members"
	$Rep = @()
    	foreach ($Group in $groups)
    	{
        	$members = net localgroup $Group | where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 4
        	$groupmem= $null
        	If ($members -ne $null)
        	{
            		foreach ($member in $members)
            		{
                		if($groupmem -ne $null)
                		{
                    			$groupmem = -join($groupmem,", ", $member)
                		}
                		Else
                		{
                    			$groupmem = -join($member)
                		}
            		}
        	}
		$Details=""|Select "Group","Members"
        	$Details.Group= $Group
        	$Details.Members = $groupmem
        	$rep +=$Details
        	$Details = $null
    	}
    	$MyReport += Get-HTMLTable ($rep)
	$MyReport += Get-CustomHeaderClose
    }
    ####################################################################################
    # Software
    ####################################################################################
	If ((get-wmiobject -ComputerName $Target -namespace "root/cimv2" -list) | Where-Object {$_.name -match "Win32_Product"})
	{
		$MyReport += Get-CustomHeader "2" "Software"
			$MyReport += Get-HTMLTable (get-wmiobject -ComputerName $Target Win32_Product| sort-object Name | Select Name,Version,Vendor)
		$MyReport += Get-CustomHeaderClose
	}
    ################################ End of Software ####################################
	####################################################################################
    # Logical Disks
    ####################################################################################
	$Disks = Get-WmiObject -ComputerName $Target Win32_LogicalDisk
	$MyReport += Get-CustomHeader "2" "Logical Disk Configuration"
		$LogicalDrives = @()
		Foreach ($LDrive in ($Disks | Where {$_.DriveType -eq 3}))
        {
			$Details = "" | Select "Drive Letter", Label, "File System", "Disk Size (GB)", "Disk Free Space (GB)"
			$Details."Drive Letter" = $LDrive.DeviceID
			$Details.Label = $LDrive.VolumeName
			$Details."File System" = $LDrive.FileSystem
			$Details."Disk Size (GB)" = [math]::Round($LDrive.size / 1GB,2)
			$Details."Disk Free Space (GB)" = [math]::round($LDrive.FreeSpace / 1GB,2)
			$LogicalDrives += $Details
		}
		$MyReport += Get-HTMLTable ($LogicalDrives)
	    $MyReport += Get-CustomHeaderClose
    ################################ End of Disks#####################################
	####################################################################################
    # Network Configuration
    ####################################################################################
	$Adapters = Get-WmiObject -ComputerName $Target Win32_NetworkAdapterConfiguration
	$MyReport += Get-CustomHeader "2" "NIC Configuration"
	$IPInfo = @()
	Foreach ($Adapter in ($Adapters | Where {$_.IPEnabled -eq $True})) 
    {
		$Details = "" | Select Description, "Physical address", "IP Address / Subnet Mask", "Default Gateway", "DHCP Enabled", DNS
		$Details.Description = "$($Adapter.Description)"
		$Details."Physical address" = "$($Adapter.MACaddress)"
		If ($Adapter.IPAddress -ne $Null) {
		$Details."IP Address / Subnet Mask" = "$($Adapter.IPAddress)/$($Adapter.IPSubnet)"
			$Details."Default Gateway" = "$($Adapter.DefaultIPGateway)"
		}
		If ($Adapter.DHCPEnabled -eq "True")	{
			$Details."DHCP Enabled" = "Yes"
		}
		Else {
			$Details."DHCP Enabled" = "No"
		}
		If ($Adapter.DNSServerSearchOrder -ne $Null)	{
			$Details.DNS =  "$($Adapter.DNSServerSearchOrder)"
		}
		$IPInfo += $Details
	}
	$MyReport += Get-HTMLTable ($IPInfo)
	$MyReport += Get-CustomHeaderClose
    ################################ End of Network#####################################
	####################################################################################
    # Service 
    ####################################################################################
	$ListOfServices = Get-WmiObject -ComputerName $Target Win32_Service
	$MyReport += Get-CustomHeader "2" "Services"
		$Services = @()
		Foreach ($Service in $ListOfServices)
        {
			$Details = "" | Select Name,Account,"Start Mode",State
			$Details.Name = $Service.Caption
			$Details.Account = $Service.Startname
			$Details."Start Mode" = $Service.StartMode
            $Details.State = $Service.State
            $Services += $Details
		}
        $MyReport += Get-HTMLTable ($Services)
	$MyReport += Get-CustomHeaderClose
	#$MyReport += Get-CustomHeaderClose
    ################################ End of Service ####################################
    ####################################################################################
    # Windows Roles and Features 
    ####################################################################################
    $MyReport += Get-CustomHeader "2" "Roles and Features"
    try
    {
        Get-WindowsFeature | ? { $_.Installed -eq "Installed"} > RnF.txt
        $rnfs=gc rnf.txt
        $Roles = @()
        foreach ($rnf in $rnfs)
        {
            $Roles += "<font color=`"##17202A`"; font size=`"2`"><pre>" +$rnf + "</pre></font>"
        }
        $MyReport += Get-HTMLRolesDetail ($Roles)
    }
    Catch
    {
	Try
	{
		Import-Module servermanager
		Get-WindowsFeature | ? { $_.Installed -eq "Installed"} > RnF.txt
		$rnfs=gc rnf.txt
	}
	Catch
	{
		Get-WmiObject Win32_ServerFeature | select Name >RnF.txt
		$rnfs=gc rnf.txt
	}
        $Roles = @()
        foreach ($rnf in $rnfs)
        {
            $Roles += "<font color=`"##17202A`"; font size=`"2`"><pre>" +$rnf + "</pre></font>"
        }
        $MyReport += Get-HTMLRolesDetail ($Roles)
    }
    $MyReport += Get-CustomHeaderClose
    Remove-Item rnf.txt -Force -ErrorAction SilentlyContinue
    ################################ End of Roles and Features###############################>

    	$MyReport += Get-CustomHeader0Close
	$MyReport += Get-CustomHTMLClose
	$MyReport += Get-CustomHTMLClose
    
    $reportime=get-date -uFormat "%d-%m-%Y %H:%M"
    $Path="C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\Inventory\HTML-Results"
    If ((Test-Path $Path) -eq $false)
    {
        New-Item $Path -type directory
    }
    else
    {
        Remove-Item -Path "C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\Inventory\HTML-Results\*.*" -Recurse -ErrorAction SilentlyContinue
    }    
    $fdate=get-Date -format "dd-MM-yy_HHmm"
    $Filename = $Path+"$($Target)_$fdate.html"
	$MyReport | out-file -encoding ASCII -filepath $Filename
    ii $Filename

